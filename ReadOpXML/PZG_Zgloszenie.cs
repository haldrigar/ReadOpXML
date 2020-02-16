﻿using System.Collections.Generic;

namespace ReadOpXML
{
    class PzgZgloszenie
    {
        public int IdFile { get; set; }
        public string XmlPath { get; set; }
        public string IdMaterialu { get; set; }
        public string PzgIdZgloszenia { get; set; }
        public string PzgPolozenieObszaru { get; set; }
        public string IdZgloszeniaJedn { get; set; }
        public string IdZgloszeniaNr { get; set; }
        public string IdZgloszeniaRok { get; set; }
        public string IdZgloszeniaEtap { get; set; }
        public string IdZgloszeniaSepJednNr { get; set; }
        public string IdZgloszeniaSepNrRok { get; set; }
        public string PzgDataZgloszenia { get; set; }
        public string Obreb { get; set; }
        public string PzgPodmiotZglaszajacyNazwa { get; set; }
        public string PzgPodmiotZglaszajacyRegon { get; set; }
        public string PzgPodmiotZglaszajacyPesel { get; set; }
        public List<string> OsobaUprawnionaList { get; set; }
        public List<string> PzgCelList { get; set; }
        public List<string> CelArchiwalnyList { get; set; }
        public string PzgRodzaj { get; set; }

    }
}
