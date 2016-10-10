using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TMC
{
    public class INI
    {
        private string iniFilePath;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        public INI()
        {
        }

        public INI(string fullPath)
        {
            iniFilePath = fullPath;
        }

        public void Write(string Section, string Key, string Value, string iniFilePath)
        {
            this.iniFilePath = iniFilePath;
            Write(Section, Key, Value);
        }

        public void Write(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, iniFilePath);
        }

        public string Read(string Section, string Key, string iniFilePath)
        {
            this.iniFilePath = iniFilePath;
            return Read(Section, Key);
        }

        public string Read(string Section, string Key)
        {
            StringBuilder sb = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, string.Empty, sb, 255, iniFilePath);
            return sb.ToString();
        }

        public void DeleteKey(string Section, string Key)
        {
            Write(Section, Key, null);
        }

        public void DeleteSection(string Section)
        {
            Write(Section, null, null);
        }

        public bool KeyExists(string Key, string Section)
        {
            return Read(Key, Section).Length > 0;
        }
    }
}
