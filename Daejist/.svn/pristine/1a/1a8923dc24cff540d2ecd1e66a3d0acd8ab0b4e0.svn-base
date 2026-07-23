using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace WeighingSystem
{
    public class IniFile
    {
        #region Properties

        /// <summary>
        /// Ini File Path
        /// </summary>
        public string IniPath
        {
            get { return _path; }
        }
        private string _path;

        #endregion

        #region Constructor

        public IniFile(string path)
        {
            if (File.Exists(path))
            {
                if (Path.GetExtension(path).ToLower() == ".ini")
                    _path = path;
                else
                    _path = "ini file 이 지정되지 않았습니다 (" + path + ")";
            }
            else
            {
                _path = "파일이 존재하지 않습니다 (" + path + ")";
            }
        }

        #endregion

        #region Methods

        public string GetIniValue(string section, string key, bool decryption = false)
        {
            StringBuilder sb = new StringBuilder(512);
            int bytes = GetPrivateProfileString(section, key, "", sb, 512, IniPath);

            if (decryption)
                return DataEncryption.Decrypt(sb.ToString());
            else
                return sb.ToString();

        }

        public bool SetIniValue(string section, string key, string value, bool encryption = false)
        {
            string encryptValue = string.Empty;
            if (encryption)
                encryptValue = DataEncryption.Encrypt(value);
            else
                encryptValue = value;

            return WritePrivateProfileString(section, key, encryptValue, IniPath);
        }

        public string[] GetSectionValues(string section)
        {
            byte[] buffer = new byte[4096];
            uint bytes = GetPrivateProfileSection(section, buffer, 4096, IniPath);

            return Encoding.Default.GetString(buffer).Split(new char[1] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public string[] GetIniSections()
        {
            byte[] buffer = new byte[4096];
            uint bytes = GetPrivateProfileSectionNames(buffer, 4096, IniPath);

            return Encoding.Default.GetString(buffer).Split(new char[1] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public bool DeleteIniValue(string section, string key)
        {
            return WritePrivateProfileString(section, key, null, IniPath);
        }

        public bool DeleteIniSection(string section)
        {
            return WritePrivateProfileString(section, null, null, IniPath);
        }

        #endregion

        #region Implementations

        /// <summary>
        /// 섹션과 키로 검색하여 값을 문자열형으로 읽어옵니다.
        /// </summary>
        /// <param name="lpAppName">섹션명</param>
        /// <param name="lpKeyName">키값</param>
        /// <param name="lpDefault">기본값</param>
        /// <param name="lpReturnedString">가져온문자열</param>
        /// <param name="nSize">문자열버퍼크기</param>
        /// <param name="lpFileName">파일이름</param>
        /// <returns>가져온문자열의크기</returns>
        [DllImport("kernel32")]
        public static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        /// <summary>
        /// 섹션과 키로 검색하여 값을 저장합니다.
        /// </summary>
        /// <param name="lpAppName">섹션명</param>
        /// <param name="lpKeyName">키값</param>
        /// <param name="lpString">저장할문자열</param>
        /// <param name="lpFileName">파일이름</param>
        /// <returns>저장성공여부</returns>
        [DllImport("kernel32")]
        public static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        /// <summary>
        /// 섹션과 키로 검색하여 값을 Inteager 형으로 불러옵니다.
        /// </summary>
        /// <param name="lpAppName">섹션명</param>
        /// <param name="lpKeyName">키값</param>
        /// <param name="nDefault">기본값</param>
        /// <param name="lpFileName">파일이름</param>
        /// <returns>검색된값, 해당키로 검색실패시 기본값으로 대체됨</returns>
        [DllImport("kernel32")]
        public static extern uint GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

        /// <summary>
        /// 섹션으로 검색하여 키와 값을 Pair 형태로 가져옵니다.
        /// </summary>
        /// <param name="IpAppName">섹션명</param>
        /// <param name="IpPairValues">Pair한 키와 값을 담을 배열</param>
        /// <param name="nSize">배열의크기</param>
        /// <param name="IpFileName">파일이름</param>
        /// <returns>읽어온바이트수</returns>
        [DllImport("kernel32.dll")]
        public static extern uint GetPrivateProfileSection(string IpAppName, byte[] IpPairValues, uint nSize, string IpFileName);

        /// <summary>
        /// 섹션을 가져옵니다.
        /// </summary>
        /// <param name="IpSections">섹션의 리스트를 직렬화하여 담을 배열</param>
        /// <param name="nSize">배열의크기</param>
        /// <param name="IpFileName">파일이름</param>
        /// <returns>읽어온바이트수</returns>
        [DllImport("kernel32.dll")]
        public static extern uint GetPrivateProfileSectionNames(byte[] IpSections, uint nSize, string IpFileName);

        #endregion
    }
}
