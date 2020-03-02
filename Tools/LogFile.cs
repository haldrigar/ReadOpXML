using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace ReadOpXML.Tools
{
    public static class LogFile
    {
        public static void SaveMessage(string message, [CallerMemberName] string callerName = "")
        {
            using (StreamWriter str = new StreamWriter(new FileStream("ReadOpXML.log", FileMode.Append), Encoding.UTF8))
            {
                str.WriteLine($"{DateTime.Now}\t{callerName}:\t{message}");
            }
        }
    }
}
