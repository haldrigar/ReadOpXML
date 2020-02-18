using System.Collections.Generic;

namespace ReadOpXML.Schemat
{
    internal class PzgMaterialZasobu
    {
        public int IdFile { get; set; }
        public string XmlPath { get; set; }
        public string IdMaterialu { get; set; }
        public string IdMaterialuPierwszyCzlon { get; set; }
        public string IdMaterialuDrugiCzlon { get; set; }
        public string IdMaterialuTrzeciCzlon { get; set; }
        public string IdMaterialuCzwartyCzlon { get; set; }
        public string PzgDataPrzyjecia { get; set; }
        public string DataWplywu { get; set; }
        public string PzgNazwa { get; set; }
        public string PzgPolozenieObszaru { get; set; }
        public string Obreb { get; set; }
        public string PzgTworcaOsobaId { get; set; }
        public string PzgTworcaNazwa { get; set; }
        public string PzgTworcaRegon { get; set; }
        public string PzgTworcaPesel { get; set; }
        public string PzgSposobPozyskania { get; set; }
        public string PzgPostacMaterialu { get; set; }
        public string PzgRodzNosnika { get; set; }
        public string PzgDostep { get; set; }
        public string PzgPrzyczynyOgraniczen { get; set; }
        public string PzgTypMaterialu { get; set; }
        public string PzgKatArchiwalna { get; set; }
        public string PzgJezyk { get; set; }
        public string PzgOpis { get; set; }
        public string PzgOznMaterialuZasobu { get; set; }
        public string OznMaterialuZasobuTyp { get; set; }
        public string OznMaterialuZasobuJedn { get; set; }
        public string OznMaterialuZasobuNr { get; set; }
        public string OznMaterialuZasobuRok { get; set; }
        public string OznMaterialuZasobuTom { get; set; }
        public string OznMaterialuZasobuSepJednNr { get; set; }
        public string OznMaterialuZasobuSepNrRok { get; set; }
        public string PzgDokumentWyl { get; set; }
        public string PzgDataWyl { get; set; }
        public string PzgDataArchLubBrak { get; set; }
        public List<string> PzgCelList { get; set; }
        public List<string> CelArchiwalnyList { get; set; }
        public List<string> DzialkaPrzedList { get; set; }
        public List<string> DzialkaPoList { get; set; }
        public string Opis2 { get; set; }
    }
}

