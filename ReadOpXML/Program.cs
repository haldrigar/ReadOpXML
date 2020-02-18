using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Schema;
using CommandLine;
using OfficeOpenXml;
using ReadOpXML.Schemat;
using ReadOpXML.Tools;
using static ReadOpXML.Tools.Functions;

namespace ReadOpXML
{
    public class Program
    {
        private static readonly List<WalidationLog> WalidationLogList = new List<WalidationLog>();
        private static readonly List<ModelErrorLog> ModelErrorLogList = new List<ModelErrorLog>();

        public static void Main(string[] args)
        {
            string startupPath = string.Empty;
            bool poprawa = false;
            bool walidacja = false;

            Parser.Default.ParseArguments<Options>(args).WithParsed(RunOptions).WithNotParsed(HandleParseError);

            void RunOptions(Options opts)
            {
                startupPath = opts.StarupPath;
                poprawa = opts.Poprawa;
                walidacja = opts.Walidacja;
            }

            void HandleParseError(IEnumerable<Error> errs)
            {
                Console.ReadKey(false);
                Environment.Exit(0);
            }

            Console.WriteLine("Wyszukiwanie plików XML w katalogu {0}...", startupPath);

            List<string> xmlFiles = Directory.EnumerateFiles(startupPath ?? throw new InvalidOperationException(), "*.xml", SearchOption.AllDirectories).ToList();
            
            Console.WriteLine("Znaleziono {0} plików XML.\n", xmlFiles.Count);

            bool isError = false;

            Console.WriteLine("Walidacja syntaktyczna plików XML...\n");

            foreach (string xmlFile in xmlFiles)
            {
                try
                {
                    XmlDocument doc = new XmlDocument {PreserveWhitespace = true};
                    doc.Load(xmlFile);
                }
                catch (XmlException e)
                {
                    Console.WriteLine($@"{xmlFile}: {e.Message}");
                    LogFile.SaveMessage(Path.Combine(startupPath, "ReadOpXML.log"), xmlFile + " => " + e.Message);

                    isError = true;
                }
            }

            if (isError)
            {
                Console.WriteLine("Musisz ręcznie poprawić wskazane błedy by zaczytać pliki XML!!!");
                Console.ReadKey(true);
                Environment.Exit(0);
            }

            if (poprawa)
            {
                Console.WriteLine("Automatyczna poprawa plików XML...\n");

                foreach (string xmlFile in xmlFiles)
                {
                    XmlDocument doc = new XmlDocument {PreserveWhitespace = true};

                    doc.Load(xmlFile);
                    doc.FixErrors();
                    doc.Save(xmlFile);
                }
            }

            Console.WriteLine(walidacja ? "Wczytywanie i walidacja plików XML...\n" : "Wczytywanie plików XML...\n");

            Dictionary<int, PzgMaterialZasobu> pzgMaterialZasobuDict = new Dictionary<int, PzgMaterialZasobu>();
            List<PzgCel> pzgMaterialZasobuCelList = new List<PzgCel>();
            List<CelArchiwalny> pzgMaterialZasobuCelArchList = new List<CelArchiwalny>();
            List<DzialkaPrzed> pzgMaterialZasobuDzialkaPrzedList = new List<DzialkaPrzed>();
            List<DzialkaPo> pzgMaterialZasobuDzialkaPoList = new List<DzialkaPo>();

            Dictionary<int, PzgZgloszenie> pzgZgloszenieDict = new Dictionary<int, PzgZgloszenie>();
            List<PzgCel> pzgZgloszenieCelList = new List<PzgCel>();
            List<CelArchiwalny> pzgZgloszenieCelArchList = new List<CelArchiwalny>();
            List<OsobaUprawniona> pzgZgloszenieOsobaUprawnionaList = new List<OsobaUprawniona>();

            XmlReaderSettings settings = new XmlReaderSettings();

            if (walidacja)
            {
                settings.Schemas.Add("http://www.w3.org/2001/XMLSchema", startupPath + "\\xsd\\schemat.xsd");
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationFlags = XmlSchemaValidationFlags.AllowXmlAttributes | XmlSchemaValidationFlags.ProcessIdentityConstraints | XmlSchemaValidationFlags.ProcessInlineSchema | XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += ValidationEventHandler;
            }

            int idFile = 0;

            foreach (string xmlFile in xmlFiles)
            {
                idFile++;

                XmlReader reader = XmlReader.Create(xmlFile, settings);

                XmlDocument doc = new XmlDocument {PreserveWhitespace = true};

                doc.Load(reader);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("xmls", "http://www.w3.org/2001/XMLSchema");

                PzgMaterialZasobu pzgMaterialZasobu = new PzgMaterialZasobu
                {
                    IdFile = idFile,
                    XmlPath = xmlFile,
                    IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile)

                };

                pzgMaterialZasobu.IdMaterialuPierwszyCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "pierwszyCzlon");
                pzgMaterialZasobu.IdMaterialuDrugiCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "drugiCzlon");
                pzgMaterialZasobu.IdMaterialuTrzeciCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "trzeciCzlon");
                pzgMaterialZasobu.IdMaterialuCzwartyCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "czwartyCzlon");
                pzgMaterialZasobu.PzgDataPrzyjecia = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataPrzyjecia");
                pzgMaterialZasobu.DataWplywu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "dataWplywu");
                pzgMaterialZasobu.PzgNazwa = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_nazwa");
                pzgMaterialZasobu.PzgPolozenieObszaru = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_polozenieObszaru");
                pzgMaterialZasobu.Obreb = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "obreb");
                pzgMaterialZasobu.PzgTworcaOsobaId = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_tworca", "osoba_id");
                pzgMaterialZasobu.PzgTworcaNazwa = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_tworca", "nazwa");
                pzgMaterialZasobu.PzgTworcaRegon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_tworca", "REGON");
                pzgMaterialZasobu.PzgTworcaPesel = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_tworca", "PESEL");
                pzgMaterialZasobu.PzgSposobPozyskania = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_sposobPozyskania");
                pzgMaterialZasobu.PzgPostacMaterialu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_postacMaterialu");
                pzgMaterialZasobu.PzgRodzNosnika = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_rodzNosnika");
                pzgMaterialZasobu.PzgDostep = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dostep");
                pzgMaterialZasobu.PzgPrzyczynyOgraniczen = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_przyczynyOgraniczen");
                pzgMaterialZasobu.PzgTypMaterialu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_typMaterialu");
                pzgMaterialZasobu.PzgKatArchiwalna = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_katArchiwalna");
                pzgMaterialZasobu.PzgJezyk = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_jezyk");
                pzgMaterialZasobu.PzgOpis = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_opis");
                pzgMaterialZasobu.PzgOznMaterialuZasobu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_oznMaterialuZasobu");
                pzgMaterialZasobu.OznMaterialuZasobuTyp = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuTyp");
                pzgMaterialZasobu.OznMaterialuZasobuJedn = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuJedn");
                pzgMaterialZasobu.OznMaterialuZasobuNr = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuNr");
                pzgMaterialZasobu.OznMaterialuZasobuRok = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuRok");
                pzgMaterialZasobu.OznMaterialuZasobuTom = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuTom");
                pzgMaterialZasobu.OznMaterialuZasobuSepJednNr = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuSepJednNr");
                pzgMaterialZasobu.OznMaterialuZasobuSepNrRok = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "oznMaterialuZasobuSepNrRok");
                pzgMaterialZasobu.PzgDokumentWyl = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dokumentWyl");
                pzgMaterialZasobu.PzgDataWyl = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataWyl");
                pzgMaterialZasobu.PzgDataArchLubBrak = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataArchLubBrak");
                pzgMaterialZasobu.PzgCelList = doc.GetXmlValueList(nsmgr, "PZG_MaterialZasobu", "pzg_cel");
                pzgMaterialZasobu.CelArchiwalnyList = doc.GetXmlValueList(nsmgr, "PZG_MaterialZasobu", "celArchiwalny");
                pzgMaterialZasobu.DzialkaPrzedList = doc.GetXmlValueList(nsmgr, "PZG_MaterialZasobu", "dzialkaPrzed");
                pzgMaterialZasobu.DzialkaPoList = doc.GetXmlValueList(nsmgr, "PZG_MaterialZasobu", "dzialkaPo");
                pzgMaterialZasobu.Opis2 = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "opis2");

                pzgMaterialZasobuDict.Add(pzgMaterialZasobu.IdFile, pzgMaterialZasobu);

                XmlNodeList nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:pzg_cel", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuCelList.AddRange(from XmlNode node in nodeList
                        select new PzgCel
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = string.IsNullOrEmpty(node.InnerText) ? "--brak wartosci--" : node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:celArchiwalny", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuCelArchList.AddRange(from XmlNode node in nodeList
                        select new CelArchiwalny
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = string.IsNullOrEmpty(node.InnerText) ? "--brak wartosci--" : node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPrzed", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuDzialkaPrzedList.AddRange(from XmlNode node in nodeList
                        select new DzialkaPrzed
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = string.IsNullOrEmpty(node.InnerText) ? "--brak wartosci--" : node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPo", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuDzialkaPoList.AddRange(from XmlNode node in nodeList
                        select new DzialkaPo
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = string.IsNullOrEmpty(node.InnerText) ? "--brak wartosci--" : node.InnerText
                        });

                PzgZgloszenie pzgZgloszenie = new PzgZgloszenie
                {
                    IdFile = idFile,
                    XmlPath = xmlFile,
                    IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile)
                };

                pzgZgloszenie.PzgIdZgloszenia = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_idZgloszenia");
                pzgZgloszenie.IdZgloszeniaJedn = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaJedn");
                pzgZgloszenie.IdZgloszeniaNr = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaNr");
                pzgZgloszenie.IdZgloszeniaRok = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaRok");
                pzgZgloszenie.IdZgloszeniaEtap = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaEtap");
                pzgZgloszenie.IdZgloszeniaSepJednNr = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaSepJednNr");
                pzgZgloszenie.IdZgloszeniaSepNrRok = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaSepNrRok");
                pzgZgloszenie.PzgDataZgloszenia = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_dataZgloszenia");
                pzgZgloszenie.PzgPolozenieObszaru = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_polozenieObszaru");
                pzgZgloszenie.Obreb = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "obreb");
                pzgZgloszenie.PzgPodmiotZglaszajacyOsobaId = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "osoba_id");
                pzgZgloszenie.PzgPodmiotZglaszajacyNazwa = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "nazwa");
                pzgZgloszenie.PzgPodmiotZglaszajacyRegon = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "REGON");
                pzgZgloszenie.PzgPodmiotZglaszajacyPesel = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "PESEL");
                pzgZgloszenie.OsobaUprawnionaList = doc.GetXmlValueList(nsmgr, "PZG_Zgloszenie", "osobaUprawniona");
                pzgZgloszenie.PzgCelList = doc.GetXmlValueList(nsmgr, "PZG_Zgloszenie", "pzg_cel");
                pzgZgloszenie.CelArchiwalnyList = doc.GetXmlValueList(nsmgr, "PZG_Zgloszenie", "celArchiwalny");
                pzgZgloszenie.PzgRodzaj = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_rodzaj");

                pzgZgloszenieDict.Add(pzgZgloszenie.IdFile, pzgZgloszenie);

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_Zgloszenie/xmls:pzg_cel", nsmgr);

                if (nodeList != null) 
                    pzgZgloszenieCelList.AddRange(from XmlNode node in nodeList
                        select new PzgCel
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = string.IsNullOrEmpty(node.InnerText) ? "--brak wartosci--" : node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_Zgloszenie/xmls:celArchiwalny", nsmgr);

                if (nodeList != null) 
                    pzgZgloszenieCelArchList.AddRange(from XmlNode node in nodeList
                        select new CelArchiwalny
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = string.IsNullOrEmpty(node.InnerText) ? "--brak wartosci--" : node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_Zgloszenie/xmls:osobaUprawniona", nsmgr);

                if (nodeList != null)
                {
                    foreach (XmlNode node in nodeList)
                    {
                        XmlDocument docosoba = new XmlDocument();

                        docosoba.LoadXml(node.OuterXml);

                        OsobaUprawniona osobaUprawniona = new OsobaUprawniona
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile,
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Imie = docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:imie", nsmgr)?.InnerText,
                            Nazwisko = docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:nazwisko", nsmgr)?.InnerText,
                            NumerUprawnien = docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:numer_uprawnien", nsmgr)?.InnerText
                        };

                        pzgZgloszenieOsobaUprawnionaList.Add(osobaUprawniona);
                    }
                }
            }

            // -------------------------------------------------------------------------------------------
            // Badanie zgodności operatów i zgłoszeń z modelem danych

            if (walidacja)
            {
                if (WalidationLogList.Count > 0) Console.WriteLine("");

                Console.WriteLine("Weryfikacja poprawności operatów z modelem danych...");

                foreach (PzgMaterialZasobu operat in pzgMaterialZasobuDict.Values)
                {
                    string oznMaterialuZasobu = GetOperatZgloszenieName(operat.OznMaterialuZasobuJedn, operat.OznMaterialuZasobuNr, operat.OznMaterialuZasobuRok, operat.OznMaterialuZasobuTom, true);

                    if (operat.PzgOznMaterialuZasobu != oznMaterialuZasobu)
                    {
                        Console.WriteLine($"Błąd: {operat.XmlPath} - Nazwa operatu nie pasuje do jego składowych lub zły separator.");
                        ModelErrorLogList.Add(new ModelErrorLog(operat.IdFile, operat.XmlPath, "Błąd", "OznMaterialuZasobu", operat.PzgOznMaterialuZasobu,"Nazwa operatu nie pasuje do jego składowych lub zły separator."));
                    }

                    int operatCount = pzgMaterialZasobuDict.Values.Count(o => o.PzgOznMaterialuZasobu == operat.PzgOznMaterialuZasobu && o.Obreb == operat.Obreb);

                    if (operatCount > 1)
                    {
                        Console.WriteLine($"Błąd: {operat.XmlPath} - Duplikat nazwy operatu.");
                        ModelErrorLogList.Add(new ModelErrorLog(operat.IdFile, operat.XmlPath, "Błąd", "OznMaterialuZasobu", operat.PzgOznMaterialuZasobu,"Duplikat nazwy operatu."));
                    }
                }

                Console.WriteLine("");

                // -------------------------------------------------------------------------------------------

                Console.WriteLine("Weryfikacja poprawności zgłoszeń z modelem danych...");

                foreach (PzgZgloszenie zgloszenie in pzgZgloszenieDict.Values)
                {
                    string pzgIdZgloszenia = GetOperatZgloszenieName(zgloszenie.IdZgloszeniaJedn, zgloszenie.IdZgloszeniaNr, zgloszenie.IdZgloszeniaRok, zgloszenie.IdZgloszeniaEtap, false);                    

                    if (zgloszenie.PzgIdZgloszenia != pzgIdZgloszenia)
                    {
                        Console.WriteLine($"Błąd: {zgloszenie.XmlPath} - Nazwa zgłoszenia nie pasuje do jego składowych lub zły separator.");
                        ModelErrorLogList.Add(new ModelErrorLog(zgloszenie.IdFile, zgloszenie.XmlPath, "Błąd", "PzgIdZgloszenia", zgloszenie.PzgIdZgloszenia,"Nazwa zgłoszenia nie pasuje do jego składowych lub zły separator."));
                    }

                    int zgloszenieCount = pzgZgloszenieDict.Values.Count(o => o.PzgIdZgloszenia == zgloszenie.PzgIdZgloszenia);

                    if (zgloszenieCount > 1)
                    {
                        int zgloszenieMultiAttributesCount = pzgZgloszenieDict.Values.Count(z =>
                            z.PzgIdZgloszenia == zgloszenie.PzgIdZgloszenia &&
                            z.PzgDataZgloszenia == zgloszenie.PzgDataZgloszenia &&
                            z.Obreb == zgloszenie.Obreb);

                        if (zgloszenieMultiAttributesCount == zgloszenieCount)
                        {
                            Console.WriteLine($"Ostrzeżenie: {zgloszenie.XmlPath} - Prawdopodobny duplikat nazwy zgłoszenia.");
                            ModelErrorLogList.Add(new ModelErrorLog(zgloszenie.IdFile, zgloszenie.XmlPath, "Ostrzeżenie", "PzgIdZgloszenia", zgloszenie.PzgIdZgloszenia,"Prawdopodobny duplikat nazwu zgłoszenia."));
                        }
                        else
                        {
                            Console.WriteLine($"Błąd: {zgloszenie.XmlPath} - Duplikat nazwy zgłoszenia.");
                            ModelErrorLogList.Add(new ModelErrorLog(zgloszenie.IdFile, zgloszenie.XmlPath, "Błąd", "PzgIdZgloszenia", zgloszenie.PzgIdZgloszenia,"Duplikat nazwu zgłoszenia."));
                        }
                    }
                }
            }

            // -------------------------------------------------------------------------------------------

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                string[] arkusze = {"operaty", "operaty_cel", "operaty_cel_arch", "operaty_dzialka_przed", "operaty_dzialka_po", "zgłoszenia", "zgłoszenia_cel", "zgłoszenia_cel_arch", "zgłoszenia_osoba_uprawniona", "walidacja", "model"};

                if (ModelErrorLogList.Count > 0 || WalidationLogList.Count > 0) Console.WriteLine("");

                foreach (string arkusz in arkusze)
                {
                    Console.WriteLine($"Eksport danych do XLS [{arkusz}]...");    

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add(arkusz);

                    switch (arkusz)
                    {
                        case "operaty" :

                            sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuDict.Values, true);

                            Console.WriteLine("\nUzupełnianie pzg_cel, celArchiwalny, dzialkaPrzed, dzialkaPo...");

                            for (int i = 2; i < sheet.Dimension.End.Row; i++)
                            {
                                int idFIle = (int)sheet.Cells[$"A{i}"].Value;

                                sheet.Cells[$"AK{i}"].Value = pzgMaterialZasobuDict[idFIle].PzgCelList
                                    .Aggregate(string.Empty, (current, wartosc) => current + "|" + wartosc)
                                    .TrimStart('|');

                                sheet.Cells[$"AL{i}"].Value = pzgMaterialZasobuDict[idFIle].CelArchiwalnyList
                                    .Aggregate(string.Empty, (current, wartosc) => current + "|" + wartosc)
                                    .TrimStart('|');

                                sheet.Cells[$"AM{i}"].Value = pzgMaterialZasobuDict[idFIle].DzialkaPrzedList
                                    .Aggregate(string.Empty, (current, wartosc) => current + "|" + wartosc)
                                    .TrimStart('|');

                                sheet.Cells[$"AN{i}"].Value = pzgMaterialZasobuDict[idFIle].DzialkaPoList
                                    .Aggregate(string.Empty, (current, wartosc) => current + "|" + wartosc)
                                    .TrimStart('|');
                            }

                            Console.WriteLine("\nWeryfikacja pzg_polozenieObszaru...");

                            for (int i = 2; i < sheet.Dimension.End.Row; i++)
                            {
                                if (sheet.Cells[$"K{i}"].Text.Length > 32767)
                                {
                                    sheet.Cells[$"K{i}"].Value = sheet.Cells[$"K{i}"].Text.Substring(0, 32766);

                                    Console.WriteLine(sheet.Cells[$"B{i}"].Text + "\tpzg_polozenieObszaru jest zbyt duży!");
                                }
                            }

                            Console.WriteLine("\nWeryfikacja dzialkaPrzed...");

                            for (int i = 2; i < sheet.Dimension.End.Row; i++)
                            {
                                if (sheet.Cells[$"AM{i}"].Text.Length > 32767)
                                {
                                    sheet.Cells[$"AM{i}"].Value = sheet.Cells[$"AM{i}"].Text.Substring(0, 32766);

                                    Console.WriteLine(sheet.Cells[$"B{i}"].Text + "\tdzialkaPrzed jest zbyt duży!");
                                }
                            }

                            Console.WriteLine("\nWeryfikacja dzialkaPo...");

                            for (int i = 2; i < sheet.Dimension.End.Row; i++)
                            {
                                if (sheet.Cells[$"AN{i}"].Text.Length > 32767)
                                {
                                    sheet.Cells[$"AN{i}"].Value = sheet.Cells[$"AN{i}"].Text.Substring(0, 32766);

                                    Console.WriteLine(sheet.Cells[$"B{i}"].Text + "\tdzialkaPo jest zbyt duży!");
                                }
                            }

                            Console.WriteLine("");

                            break;

                        case "operaty_cel":
                            sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuCelList, true);
                            break;

                        case "operaty_cel_arch":
                            sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuCelArchList, true);
                            break;

                        case "operaty_dzialka_przed":
                            sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuDzialkaPrzedList, true);
                            break;

                        case "operaty_dzialka_po":
                            sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuDzialkaPoList, true);
                            break;

                        case "zgłoszenia":
                            sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieDict.Values, true);

                            Console.WriteLine("\nUzupełnianie osobaUprawniona, pzg_cel, celArchiwalny...");

                            for (int i = 2; i < sheet.Dimension.End.Row; i++)
                            {
                                int idFIle = (int)sheet.Cells[$"A{i}"].Value;

                                sheet.Cells[$"R{i}"].Value = pzgZgloszenieDict[idFIle].OsobaUprawnionaList
                                    .Aggregate(string.Empty, (current, wartosc) => current + "|" + wartosc)
                                    .TrimStart('|');

                                sheet.Cells[$"S{i}"].Value = pzgZgloszenieDict[idFIle].PzgCelList
                                    .Aggregate(string.Empty, (current, wartosc) => current + "|" + wartosc)
                                    .TrimStart('|');

                                sheet.Cells[$"T{i}"].Value = pzgZgloszenieDict[idFIle].CelArchiwalnyList
                                    .Aggregate(string.Empty, (current, wartosc) => current + "|" + wartosc)
                                    .TrimStart('|');
                            }

                            Console.WriteLine("\nWeryfikacja pzg_polozenieObszaru...");

                            for (int i = 2; i < sheet.Dimension.End.Row; i++)
                            {
                                if (sheet.Cells[$"E{i}"].Text.Length > 32767)
                                {
                                    sheet.Cells[$"E{i}"].Value = sheet.Cells[$"E{i}"].Text.Substring(0, 32766);

                                    string xmlPath = sheet.Cells[$"B{i}"].Text;

                                    Console.WriteLine(xmlPath + "\tpzg_polozenieObszaru jest zbyt duży!");
                                }
                            }

                            Console.WriteLine("");

                            break;

                        case "zgłoszenia_cel":
                            sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieCelList, true);
                            break;

                        case "zgłoszenia_cel_arch":
                            sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieCelArchList, true);
                            break;

                        case "zgłoszenia_osoba_uprawniona":
                            sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieOsobaUprawnionaList, true);
                            break;

                        case "walidacja":
                            sheet.Cells[1, 1].LoadFromCollection(WalidationLogList, true);
                            break;

                        case "model":
                            sheet.Cells[1, 1].LoadFromCollection(ModelErrorLogList, true);
                            break;
                    }

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    ExcelRange range = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                    sheet.Tables.Add(range, arkusz);

                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                using (FileStream fileStream = new FileStream(Path.Combine(startupPath, "xml.xlsx"), FileMode.Create))
                {
                    Console.WriteLine("\nZapis pliku XLS...");

                    excelPackage.SaveAs(fileStream);
                }
            }
            
            Console.WriteLine("\nKoniec");

            Console.ReadKey(true);
        }

        private static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            XmlReader localSender = (XmlReader) sender;

            WalidationLog err = new WalidationLog
            {
                FileName = localSender.BaseURI,
                Element = localSender.LocalName,
                Line = e.Exception.LineNumber + ":" + e.Exception.LinePosition,
                Message = e.Message
            };

            WalidationLogList.Add(err);

            Console.WriteLine($"Błąd walidacji: {localSender.BaseURI} - {localSender.LocalName}");
        }
    }
}
