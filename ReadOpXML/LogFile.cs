using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace ReadOpXML
{ 
    public static class LogFile
    {
        public static void SaveMessage(string fileName, string message, [CallerMemberName] string callerName = "")
        {
            string appdataPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string cfgfilePath = Path.Combine(appdataPath ?? throw new InvalidOperationException(), fileName);

            using (StreamWriter str = new StreamWriter(new FileStream(cfgfilePath, FileMode.Append), Encoding.UTF8))
            {
                str.WriteLine($"{DateTime.Now}\t{callerName.PadRight(25, ' ')}:\t{message}");
            }
        }
    }
}
