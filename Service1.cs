using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;
using System.Reflection;
using System.IO;
using System.Net;

namespace FTPDownloadService
{
    public partial class Service1 : ServiceBase
    {
        private bool _cancel;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            _cancel = false;

            //Task.Run(Processing());
            Task.Run(() => Processing());

            EventLog.WriteEntry("Service to download files from FTP CB is start", EventLogEntryType.Information);
        }

        private void Processing()
        {            
            try
            {
                while (!_cancel)
                {
                    Logger.WriteLine("**********************Старт процесса проверки поступления новых файлов**********************");

                    String server = ReadSetting("Адрес сервера"); ;
                    String folderForDBF = ReadSetting("Папка для справочников");
                    String folderForAll = ReadSetting("Папка для сохранения всех файлов");
                    String timeToDownloadString = ReadSetting("Интервал синхронизации");
                    int timeToDownload = Convert.ToInt32(timeToDownloadString) * 60000;

                    if (!Directory.Exists(@"C:\UCBApps")) Directory.CreateDirectory(@"C:\UCBApps");
                    if (!Directory.Exists(@"C:\UCBApps\Temp")) Directory.CreateDirectory(@"C:\UCBApps\Temp");
                    if (!Directory.Exists(@"C:\UCBApps\Txt")) Directory.CreateDirectory(@"C:\UCBApps\Txt");

                    DirectoryInfo destinDBF = new DirectoryInfo(folderForDBF + "\\");
                    DirectoryInfo sourceTemp = new DirectoryInfo(@"Temp");

                    string currentDate = DateTime.Now.Date.ToString("dd-MM-yy") + ".txt";
                    string remoteUri = server + currentDate; //DD-MM-YY.txt         
                    currentDate = @"C:\UCBApps\Txt\" + currentDate;

                    if (!File.Exists(currentDate))
                    {
                        Logger.WriteLine("Текущий файл рассылки " + remoteUri);
                        if (!downloadFile(remoteUri, currentDate)) return;

                        //Удаление всех файлов из папки Temp            
                        Logger.WriteLine("Начался процесс очистки папки Temp");

                        deletTempFiles();

                        //Скачивание всех файлов, указанных в файле-рассылке
                        Logger.WriteLine("Запуск процесса скачивания файлов согласно файлу рассылки");

                        downloadAllFils();
                        Thread.Sleep(1000);

                        //Копирование всех файлов из Temp в папку сохранения всех файлов
                        Logger.WriteLine("Запуск процесса копирования всех файлов в общую папку");
                        copyAllFiles();
                        Thread.Sleep(5000);

                        //Разархивирование всех файлов в папке Temp 

                        unArjAllFiles();
                        Thread.Sleep(10000);

                        //Копирование файлов *.dbf в папку справочников
                        Logger.WriteLine("Начался процесс копирования всех файлов справочников в папку " + folderForDBF);

                        copyToDBF();
                    }
                    else
                    {
                        Logger.WriteLine("За текущую дату" + DateTime.Now.Date.ToString("dd-MM-yy") + " файл рассылки уже был скачан и обработан");
                    }

                    Logger.WriteLine("**********************Стоп процесса проверки поступления новых файлов**********************");
                    Thread.Sleep(timeToDownload);

                }
                

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
                Logger.WriteLine(ex.Message);
            }

            //throw new NotImplementedException();
        }

        static void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error writing app settings");
            }
        }

        static String ReadSetting(string key)
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                string result = appSettings[key] ?? "Not Found";
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                return "Error reading app settings";
            }
        }

        static class Logger
        {
            //----------------------------------------------------------
            // Статический метод записи строки в файл лога без переноса
            //----------------------------------------------------------
            public static void Write(string text)
            {
                using (StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\log.txt", true))
                {
                    sw.Write(text);
                }
            }

            //---------------------------------------------------------
            // Статический метод записи строки в файл лога с переносом
            //---------------------------------------------------------
            public static void WriteLine(string message)
            {
                using (StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\log.txt", true))
                {
                    sw.WriteLine(String.Format("{0,-23} {1}", DateTime.Now.ToString() + ":", message));
                }
            }
        }

        bool downloadFile(String uri, String path)
        {
            try
            {
                WebClient myWebClient = new WebClient();
                myWebClient.DownloadFile(uri, path);
                Logger.WriteLine("Файл " + uri + " скачался в файл " + path);
                return true;
            }
            catch
            {
                Logger.WriteLine("Ошибка при скачивании файла " + uri);
                return false;
            }
        }

        void deletTempFiles()
        {
            DirectoryInfo sourceTemp = new DirectoryInfo(@"C:\UCBApps\Temp");
            foreach (var item in sourceTemp.GetFiles())
            {
                try
                {
                    item.Delete();
                    Logger.WriteLine("Файл " + item.Name + " удален из папки Temp");
                }
                catch
                {
                    Logger.WriteLine("Ошибка удаления файла " + item.Name + " из папки Temp");
                }
            }
        }

        void downloadAllFils()
        {
            string currentDate = @"C:\UCBApps\Txt\" + DateTime.Now.Date.ToString("dd-MM-yy") + ".txt";
            string server = ReadSetting("Адрес сервера");

            var lines = File.ReadAllLines(currentDate);
            int i = 0;
            foreach (var line in lines)
            {
                i++;
                string[] words = line.Split(new char[] { ' ' });
                if (!words[0].StartsWith("!!!") & i > 2)
                {
                    string fullPathToFile = server + words[0];
                    fullPathToFile = fullPathToFile.Replace("\\", "/");
                    string[] filePathMass = words[0].Split(new char[] { '\\' });
                    string fileNameFTP = @"C:\UCBApps\Temp" + "\\" + filePathMass[filePathMass.Length - 1].ToString();
                    Logger.WriteLine("Попытка скачать файл " + fullPathToFile + " в " + fileNameFTP);
                    downloadFile(fullPathToFile, fileNameFTP);
                }
            }

        }

        void copyAllFiles()
        {
            string folderForAll = ReadSetting("Папка для сохранения всех файлов");
            DirectoryInfo sourceTemp = new DirectoryInfo(@"c:\UCBApps\Temp\");

            string currentMonth = DateTime.Now.Date.ToString("MM");
            string currentYear = DateTime.Now.Date.ToString("yyyy");

            if (!Directory.Exists(folderForAll + "\\" + currentYear))
            {
                Directory.CreateDirectory(folderForAll + "\\" + currentYear);
                Logger.WriteLine("Создана папка " + folderForAll + "\\" + currentYear);
            }

            if (!Directory.Exists(folderForAll + "\\" + currentYear + "\\" + currentMonth))
            {
                Directory.CreateDirectory(folderForAll + "\\" + currentYear + "\\" + currentMonth);
                Logger.WriteLine("Создана папка " + folderForAll + "\\" + currentYear + "\\" + currentMonth);
            }

            Logger.WriteLine("Начался процесс копирования всех файлов в папку " + folderForAll + "\\" + currentYear + "\\" + currentMonth);

            DirectoryInfo destin = new DirectoryInfo(folderForAll + "\\" + currentYear + "\\" + currentMonth + "\\");

            foreach (var item in sourceTemp.GetFiles())
            {
                item.CopyTo(destin + item.Name, true);
                Logger.WriteLine("Файл " + item.Name + " скопирован в " + folderForAll + "\\" + currentYear + "\\" + currentMonth);
            }

        }

        void unArjAllFiles()
        {
            String pathTo7Zip = ReadSetting("Путь к архиватору 7-Zip");
            try
            {
                Process.Start(pathTo7Zip, " x " + "-y " + @" -oC:\UCBApps\Temp " + @"C:\UCBApps\Temp\*.arj");
                Logger.WriteLine("Начался процесс разархивирования файлов в папке Temp");
            }
            catch
            {
                Logger.WriteLine("Неудачное разархивирование в папке Temp");
            }
        }

        void copyToDBF()
        {
            string folderForDBF = ReadSetting("Папка для справочников");
            DirectoryInfo sourceTemp = new DirectoryInfo(@"C:\UCBApps\Temp");
            DirectoryInfo destinDBF = new DirectoryInfo(folderForDBF + "\\");

            foreach (var item in sourceTemp.GetFiles())
            {
                Logger.WriteLine("Обработка файла для копирования в DBF " + item.Name);
                if (item.Extension.ToString().ToLower() == ".dbf")
                {
                    item.CopyTo(destinDBF + item.Name, true);
                    Logger.WriteLine("Файл " + item.Name + " скопирован в " + folderForDBF);
                }
            }
        }

        protected override void OnStop()
        {
            _cancel = true;
            EventLog.WriteEntry("Service to download files from FTP CB is stopped", EventLogEntryType.Information);
        }
    }
}
