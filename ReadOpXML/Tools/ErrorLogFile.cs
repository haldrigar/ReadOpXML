using System.IO;
using System.Text;

namespace ReadOpXML.Tools
{ 
    public static class ErrorLogFile
    {
        public static void SaveMessage(string fileName, string message)
        {
            using (StreamWriter str = new StreamWriter(new FileStream(fileName, FileMode.Append), Encoding.UTF8))
            {
                str.WriteLine(message);
            }
        }
    }
}
