using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WeighingSystem
{
    public static class Log
    {
        #region Properities

        /// <summary>
        /// 프로세스 정보를 반환한다.
        /// </summary>
        public static Process Process
        {
            get { return Process.GetCurrentProcess(); }
        }

        /// <summary>
        /// 프로세스명을 반환한다.
        /// </summary>
        public static string ProcessName
        {
            get { return Process.ProcessName; }
        }

        /// <summary>
        /// 실행파일의 위치를 반환한다.
        /// </summary>
        public static string ProcessPath
        {
            get
            {
                string path = Log.Process.MainModule.FileName;
                path = path.Substring(0, path.LastIndexOf(@"\"));

                return path;
            }
        }

        /// <summary>
        /// 로그를 기록할 디렉토리 위치를 반환한다.
        /// </summary>
        public static string LogDir
        {
            get { return string.Concat(ProcessPath, @"\log"); }
        }

        /// <summary>
        /// 로그를 기록할 기본 경로를 반환한다.
        /// </summary>
        public static string DefaultPath
        {
            get
            {
                string timeString = string.Empty;
                timeString = DateTime.Now.ToString("yyyyMMdd_HH");

                string path = string.Empty;
                path = string.Concat(ProcessPath, @"\log", @"\log_", timeString, ".log");

                return path;

            }
        }

        private static Object _syncObject = new Object();

        #endregion

        #region Constructor

        static Log()
        {
            // 로그 디렉토리를 생성한다
            if (!Directory.Exists(LogDir))
            {
                Directory.CreateDirectory(LogDir);
                AddLog(string.Concat("CreateDirectory - ", LogDir));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// 프로세스가 단독적으로 실행하는지에 대한 여부를 반환한다.
        /// </summary>
        /// <returns>단독 실행 여부</returns>
        public static bool IsProcessAlone()
        {
            Process[] proc = Process.GetProcessesByName(ProcessName);

            if (proc == null || proc.Length <= 1)
                return true;

            return false;
        }

        public static void AddLog(string log)
        {
            WriteLogData(DefaultPath, log);
        }

        public static void AddLog(string path, string log)
        {
            string timeString = string.Empty;
            timeString = DateTime.Now.ToString("yyyyMMdd_HH");

            string newPath = string.Empty;
            newPath = string.Concat(path, @"\log_", timeString, ".log");

            // 2018.02.09
            // - 해당 경로가 존재하지 않을 경우 생성한다
            // 로그 디렉토리를 생성한다
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                AddLog(string.Concat("CreateDirectory - ", path));
            }

            WriteLogData(newPath, log);
        }

        public static void DeleteLogFile(int day)
        {
            DirectoryInfo di = new DirectoryInfo(LogDir);

            DeleteFile(di, ".log", day);
        }

        public static void DeleteLogFile(string path, int day)
        {
            DirectoryInfo di = new DirectoryInfo(path);

            DeleteFile(di, ".log", day);
        }

        public static void DeleteOldFile(string path, string extension, int day)
        {
            DirectoryInfo di = new DirectoryInfo(path);

            DeleteFile(di, extension, day);
        }

        /// <summary>
        /// 오래된 디렉토리를 삭제함 (히스토리 데이터 삭제 용도)
        /// - 하위디렉토리를 포함하여 모두 삭제시킨다
        /// </summary>
        /// <param name="path"></param>
        /// <param name="day"></param>
        public static void DeleteOldDirectory(string path, int day)
        {
            DirectoryInfo di = new DirectoryInfo(path);

            DeleteDirectory(di, day, true);
        }

        #endregion

        #region Implementations

        private static void WriteLogData(string path, string log)
        {
            // string logData = string.Empty;
            // logData = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ☞ " + log + "\r\n";

            StringBuilder sb = new StringBuilder();

            sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.Append(" ☞ ");
            sb.Append(log);
            sb.Append("\r\n");

            try
            {
                lock (_syncObject)
                {
                    FileInfo fi = new FileInfo(path);

                    StreamWriter sw = fi.AppendText();
                    // sw.WriteLine(logData);
                    sw.WriteLine(sb.ToString());
                    sw.Flush();
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void DeleteFile(DirectoryInfo dirInfo, string extension, int day)
        {
            foreach (FileInfo fi in dirInfo.GetFiles())
            {
                if (fi.Extension != extension)
                    continue;

                if (fi.LastWriteTime <= DateTime.Now.AddDays(-day))
                {
                    WriteLogData(DefaultPath, "파일 삭제 - " + fi.Name);
                    fi.Delete();
                }
            }
        }

        private static void DeleteDirectory(DirectoryInfo dirInfo, int day, bool deleteSubDirectory = false)
        {
            foreach (DirectoryInfo di in dirInfo.GetDirectories())
            {
                if (di.LastWriteTime <= DateTime.Now.AddDays(-day))
                {
                    WriteLogData(DefaultPath, "디렉토리 삭제 - " + di.Name);
                    di.Delete(deleteSubDirectory);
                }
            }
        }

        #endregion
    }
}
