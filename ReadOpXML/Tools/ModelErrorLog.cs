namespace ReadOpXML.Tools
{
    public class ModelErrorLog
    {
        public int IdFile { get; set; } 
        public string FileName { get; set; }
        public string Typ { get; set; }
        public string Atrybut { get; set; }
        public string Wartosc { get; set; }
        public string Info { get; set; }

        public ModelErrorLog(int idFile, string fileName, string typ, string atrybut, string wartosc, string info)
        {
            IdFile = idFile;
            FileName = fileName;
            Typ = typ;
            Atrybut = atrybut;
            Wartosc = wartosc;
            Info = info;
        }
    }
}
