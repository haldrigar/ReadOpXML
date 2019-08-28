using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml;

namespace ReadOpXML
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            string starupPath = args.Length > 0 ? args[0] : Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            List<string> xmlFiles = Directory.EnumerateFiles(starupPath ?? throw new InvalidOperationException(), "*.xml", SearchOption.AllDirectories).ToList();

            List<PzgMaterialZasobu> pzgMaterialZasobuList = new List<PzgMaterialZasobu>();
            List<PzgZgloszenie> pzgZgloszenieList = new List<PzgZgloszenie>();

            foreach (string xmlFile in xmlFiles)
            {
                Console.Write(@".");

                XmlDocument doc = new XmlDocument();

                try
                {
                    doc.Load(xmlFile);
                }
                catch (Exception e)
                {
                    Console.WriteLine($@"{xmlFile}: {e.Message}");
                    continue;
                }

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("xmls", "http://www.w3.org/2001/XMLSchema");

                PzgMaterialZasobu pzgMaterialZasobu = new PzgMaterialZasobu
                {
                    XmlFile = xmlFile,

                    IdMaterialuPierwszyCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "pierwszyCzlon"),
                    IdMaterialuDrugiCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "drugiCzlon"),
                    IdMaterialuTrzeciCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "trzeciCzlon"),
                    IdMaterialuCzwartyCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "czwartyCzlon"),
                    PzgDataPrzyjecia = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataPrzyjecia"),
                    PzgDataWplywu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataWplywu"),
                    PzgNazwa = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_nazwa"),
                    Obreb = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "obreb"),
                    PzgTworcaNazwa = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_tworca", "nazwa"),
                    PzgTworcaRegon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_tworca", "REGON"),
                    PzgTworcaPesel = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_tworca", "PESEL"),
                    PzgSposobPozyskania = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_sposobPozyskania"),
                    PzgPostacMaterialu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_postacMaterialu"),
                    PzgRodzNosnika = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_rodzNosnika"),
                    PzgDostep = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dostep"),
                    PzgPrzyczynyOgraniczen = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_przyczynyOgraniczen"),
                    PzgTypMaterialu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_typMaterialu"),
                    PzgKatArchiwalna = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_katArchiwalna"),
                    PzgJezyk = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_jezyk"),
                    PzgOpis = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_opis"),
                    PzgOznMaterialuZasobu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_oznMaterialuZasobu"),
                    OznMaterialuZasobuTyp = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuTyp"),
                    OznMaterialuZasobuJedn = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuJedn"),
                    OznMaterialuZasobuNr = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuNr"),
                    OznMaterialuZasobuRok = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuRok"),
                    OznMaterialuZasobuTom = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuTom"),
                    OznMaterialuZasobuSepJednNr = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuSepJednNr"),
                    OznMaterialuZasobuSepNrRok = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuSepNrRok"),
                    PzgDokumentWyl = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dokumentWyl"),
                    PzgDataWyl = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataWyl"),
                    PzgDataArchLubBrak = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataArchLubBrak"),
                    PzgCel = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_cel"),
                    CelArchiwalny = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "celArchiwalny"),
                    DzialkaPrzed = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "dzialkaPrzed"),
                    DzialkaPo = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "dzialkaPo"),
                    Opis2 = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "opis2")
                };

                pzgMaterialZasobuList.Add(pzgMaterialZasobu);

                PzgZgloszenie pzgZgloszenie = new PzgZgloszenie()
                {
                    XmlFile = xmlFile,
                    PzgIdZgloszenia = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_idZgloszenia"),
                    IdZgloszeniaJedn = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaJedn"),
                    IdZgloszeniaNr = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaNr"),
                    IdZgloszeniaRok = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaRok"),
                    IdZgloszeniaEtap = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaEtap"),
                    IdZgloszeniaSepJednNr = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaSepJednNr"),
                    IdZgloszeniaSepNrRok = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaSepNrRok"),
                    PzgDataZgloszenia = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_dataZgloszenia"),
                    Obreb = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "obreb"),
                    PzgPodmiotZglaszajacyNazwa = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "nazwa"),
                    PzgPodmiotZglaszajacyRegon = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "REGON"),
                    PzgPodmiotZglaszajacyPesel = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "PESEL"),
                    OsobaUprawnionaImie = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "osobaUprawniona", "imie"),
                    OsobaUprawnionaNazwisko = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "osobaUprawniona", "nazwisko"),
                    OsobaUprawnionaNumerUprawnien = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "osobaUprawniona", "numer_uprawnien"),
                    PzgCel = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_cel"),
                    CelArchiwalny = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "celArchiwalny"),
                    PzgRodzaj = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_rodzaj")
                };

                pzgZgloszenieList.Add(pzgZgloszenie);
            }

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                Console.Write('\n');
                
                //  -------------------------------------------------------------------------------

                Console.WriteLine(@"Eksport danych do XLS [operaty]...");

                ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("operaty");

                sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuList, true);

                int rowsCount = sheet.Dimension.Rows;
                int columnsCount = sheet.Dimension.Columns;

                sheet.View.FreezePanes(2, 1);

                sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                sheet.Cells.AutoFitColumns(10, 50);

                //  -------------------------------------------------------------------------------

                Console.WriteLine(@"Eksport danych do XLS [zgłoszenia]...");

                sheet = excelPackage.Workbook.Worksheets.Add("zgłoszenia");

                sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieList, true);

                rowsCount = sheet.Dimension.Rows;
                columnsCount = sheet.Dimension.Columns;

                sheet.View.FreezePanes(2, 1);

                sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                sheet.Cells.AutoFitColumns(10, 50);

                //  -------------------------------------------------------------------------------

                using (FileStream fileStream = new FileStream(Path.Combine(starupPath, "xml.xlsx"), FileMode.Create))
                {
                    Console.WriteLine(@"Zapis pliku XLS...");
                    excelPackage.SaveAs(fileStream);
                }
            }
            
            Console.WriteLine(@"Koniec");
            Console.ReadKey(true);
        }

        private static string GetXmlValue(this XmlDocument doc, XmlNamespaceManager nsmgr, params string[] nodes)
        {
            string xPath = nodes.Aggregate("/xmls:schema", (current, t) => current + "/xmls:" + t);

            XmlNodeList nodeList = doc.DocumentElement?.SelectNodes(xPath, nsmgr);

            if (nodeList == null || nodeList.Count == 0)
            {
                return "--brak znacznika--";
            }

            string values = "";

            foreach (XmlNode node in nodeList)
            {
                if (node.InnerText.Contains("|")) throw new Exception("Błąd separatora!");

                values = values + "|" + node.InnerText;
            }

            return values.TrimStart('|');
        }
    }
}
