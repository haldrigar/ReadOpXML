using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Schema;
using OfficeOpenXml;

namespace ReadOpXML
{
    public class Program
    {
        private static readonly List<ErrorLog> ErrorLogsList = new List<ErrorLog>();

        public static void Main(string[] args)
        {
            bool operatyExport = true;
            bool operatyCelExport = true;
            bool operatyCelArchExport = true;
            bool operatyDzialkaPrzedExport = true;
            bool operatyDzialkaPoExport = true;

            bool zgloszeniaExport = true;
            bool zgloszeniaCelExport = true;
            bool zgloszeniaCelArchExport = true;
            bool zgloszeniaOsobaUprawnionaExport = true;

            bool walidacjaExport = true;

            bool poprawa = true;

            string starupPath = args.Length > 0 ? args[0] : Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            List<PzgMaterialZasobu> pzgMaterialZasobuList = new List<PzgMaterialZasobu>();
            List<PzgCel> pzgMaterialZasobuCelList = new List<PzgCel>();
            List<CelArchiwalny> pzgMaterialZasobuCelArchList = new List<CelArchiwalny>();
            List<DzialkaPrzed> pzgMaterialZasobuDzialkaPrzedList = new List<DzialkaPrzed>();
            List<DzialkaPo> pzgMaterialZasobuDzialkaPoList = new List<DzialkaPo>();

            List<PzgZgloszenie> pzgZgloszenieList = new List<PzgZgloszenie>();
            List<PzgCel> pzgZgloszenieCelList = new List<PzgCel>();
            List<CelArchiwalny> pzgZgloszenieCelArchList = new List<CelArchiwalny>();
            List<OsobaUprawniona> pzgZgloszenieOsobaUprawnionaList = new List<OsobaUprawniona>();

            Console.WriteLine("Wyszukiwanie plików XML...");
            List<string> xmlFiles = Directory.EnumerateFiles(starupPath ?? throw new InvalidOperationException(), "*.xml", SearchOption.AllDirectories).ToList();
            Console.WriteLine("Znaleziono {0} plików XML.\n", xmlFiles.Count);

            bool isError = false;

            Console.WriteLine("Test i wstępna poprawa plików XML...\n");

            foreach (string xmlFile in xmlFiles)
            {
                try
                {
                    XmlDocument doc = new XmlDocument {PreserveWhitespace = true};

                    doc.Load(xmlFile);

                    if (poprawa)
                    {
                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                    nsmgr.AddNamespace("xmls", "http://www.w3.org/2001/XMLSchema");

                    int rok = 0;   

                    string fileName = Path.GetFileName(xmlFile);

                    if (fileName != null && Regex.IsMatch(fileName.ToUpper(), @"^P\.[0-9]{4}\.[0-9]{4}\.[0-9]+\.XML$"))
                    {
                        rok = Convert.ToInt32(fileName.Substring(7, 4));
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA oznMaterialuZasobuSepJednNr z NULL i z pustego
                    
                    XmlNode xmlNode = doc.DocumentElement?.SelectSingleNode("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:oznMaterialuZasobuSepJednNr", nsmgr);

                    if (xmlNode != null && (xmlNode.InnerText == "null" || xmlNode.InnerText == "" || xmlNode.InnerText == "/") && rok > 2013)
                    {
                        xmlNode.InnerText = ".";
                    }

                    if (xmlNode != null && (xmlNode.InnerText == "null" || xmlNode.InnerText == "" || xmlNode.InnerText == "/") && rok > 0 && rok <=2013)
                    {
                        xmlNode.InnerText = "-";
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA idZgloszeniaSepJednNr z NULL i z pustego
                    
                    xmlNode = doc.DocumentElement?.SelectSingleNode("/xmls:schema/xmls:PZG_Zgloszenie/xmls:idZgloszeniaSepJednNr", nsmgr);

                    if (xmlNode != null && (xmlNode.InnerText == "null" || xmlNode.InnerText == "" || xmlNode.InnerText == "/") && rok > 2013)
                    {
                        xmlNode.InnerText = ".";
                    }

                    if (xmlNode != null && (xmlNode.InnerText == "null" || xmlNode.InnerText == "" || xmlNode.InnerText == "/") && rok > 0 && rok <=2013)
                    {
                        xmlNode.InnerText = "-";
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA pzg_nazwa z 'operat techniczny' na 'operatTechniczny'

                    xmlNode = doc.DocumentElement?.SelectSingleNode("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:pzg_nazwa", nsmgr);

                    if (xmlNode != null && (xmlNode.InnerText == "operat techniczny" || xmlNode.InnerText == ""))
                    {
                        xmlNode.InnerText = "operatTechniczny";
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA pzg_tworca/nazwa z pustego na 'TWÓRCA NIEZNANY'

                    xmlNode = doc.DocumentElement?.SelectSingleNode("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:pzg_tworca/xmls:nazwa", nsmgr);

                    if (xmlNode != null && xmlNode.InnerText == "")
                    {
                        xmlNode.InnerText = "TWÓRCA NIEZNANY";
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA pzg_podmiotZglaszajacy/nazwa z pustego na 'TWÓRCA NIEZNANY'

                    xmlNode = doc.DocumentElement?.SelectSingleNode("/xmls:schema/xmls:PZG_Zgloszenie/xmls:pzg_podmiotZglaszajacy/xmls:nazwa", nsmgr);

                    if (xmlNode != null && xmlNode.InnerText == "")
                    {
                        xmlNode.InnerText = "TWÓRCA NIEZNANY";
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA REGON dla PZG_MaterialZasobu/tworca

                    xmlNode = doc.DocumentElement?.SelectSingleNode("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:pzg_tworca/xmls:REGON", nsmgr);

                    if (xmlNode != null && xmlNode.InnerText.Length == 8)
                    {
                        xmlNode.InnerText = "0" + xmlNode.InnerText;
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA REGON dla PZG_Zgloszenie/pzg_podmiotZglaszajacy

                    xmlNode = doc.DocumentElement?.SelectSingleNode("/xmls:schema/xmls:PZG_Zgloszenie/xmls:pzg_podmiotZglaszajacy/xmls:REGON", nsmgr);

                    if (xmlNode != null && xmlNode.InnerText.Length == 8)
                    {
                        xmlNode.InnerText = "0" + xmlNode.InnerText;
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA spacji w numerze działki PRZED
                    
                    XmlNodeList xmlNodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPrzed", nsmgr);

                    foreach (XmlNode node in xmlNodeList)
                    {
                        node.InnerText = node.InnerText.Replace(" ", "");
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA spacji w numerze działki PO
                    
                    xmlNodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPo", nsmgr);

                    foreach (XmlNode node in xmlNodeList)
                    {
                        node.InnerText = node.InnerText.Replace(" ", "");
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA numeru w dzialkaPrzed
                    
                    xmlNodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPrzed", nsmgr);

                    foreach (XmlNode node in xmlNodeList)
                    {
                        // jedna litera na końcu numeru działki 040903_5.0006.871a
                        if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+[a-z]{1}$"))
                        {
                            node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" + node.InnerText.Substring(node.InnerText.Length - 1, 1);
                        }

                        // jedna litera na końcu numeru działki 040903_5.0006.871/11a
                        if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+\/[0-9]+[a-z]{1}$"))
                        {
                            node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" + node.InnerText.Substring(node.InnerText.Length - 1, 1);
                        }
                    }

                    //  ---------------------------------------------------------------------------
                    //  POPRAWA numeru w dzialkaPo
                    
                    xmlNodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPo", nsmgr);

                    foreach (XmlNode node in xmlNodeList)
                    {
                        // jedna litera na końcu numeru działki 040903_5.0006.871a
                        if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+[a-z]{1}$"))
                        {
                            node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" + node.InnerText.Substring(node.InnerText.Length - 1, 1);
                        }

                        // jedna litera na końcu numeru działki 040903_5.0006.871/11a
                        if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+\/[0-9]+[a-z]{1}$"))
                        {
                            node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" + node.InnerText.Substring(node.InnerText.Length - 1, 1);
                        }
                    }

                    //  ---------------------------------------------------------------------------

                    doc.Save(xmlFile);  // Zapisanie dokumentu
                    }
                }
                catch (XmlException e)
                {
                    Console.WriteLine($@"{xmlFile}: {e.Message}");
                    LogFile.SaveMessage("ReadOpXML.log", xmlFile + " => " + e.Message);

                    isError = true;
                }
            }

            if (isError)
            {
                Console.WriteLine("\nMusisz ręcznie poprawić wskazane błedy by zaczytać pliki XML!!!");
                Console.ReadKey(true);
                return;
            }

            Console.WriteLine("Wczytywanie i walidacja plików XML...");

            int idFile = 0;

            foreach (string xmlFile in xmlFiles)
            {
                idFile++;

                XmlReaderSettings settings = new XmlReaderSettings();
                settings.Schemas.Add("http://www.w3.org/2001/XMLSchema", starupPath + "\\xsd\\schemat.xsd");
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationFlags = XmlSchemaValidationFlags.AllowXmlAttributes | XmlSchemaValidationFlags.ProcessIdentityConstraints | XmlSchemaValidationFlags.ProcessInlineSchema | XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += ValidationEventHandler;

                XmlReader reader = XmlReader.Create(xmlFile, settings);

                XmlDocument doc = new XmlDocument {PreserveWhitespace = true};

                doc.Load(reader);

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("xmls", "http://www.w3.org/2001/XMLSchema");

                PzgMaterialZasobu pzgMaterialZasobu = new PzgMaterialZasobu
                {
                IdFile = idFile,
                    XmlPath = xmlFile,
                    IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                    IdMaterialuPierwszyCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "pierwszyCzlon"),
                    IdMaterialuDrugiCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "drugiCzlon"),
                    IdMaterialuTrzeciCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "trzeciCzlon"),
                    IdMaterialuCzwartyCzlon = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_IdMaterialu", "czwartyCzlon"),
                    PzgDataPrzyjecia = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_dataPrzyjecia"),
                    DataWplywu = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "dataWplywu"),
                    PzgNazwa = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_nazwa"),
                    PzgPolozenieObszaru = string.Concat(doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_polozenieObszaru").Take(32767)),
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
                    DzialkaPrzed = string.Concat(doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "dzialkaPrzed").Take(32767)),
                    DzialkaPo = string.Concat(doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "dzialkaPo").Take(32767)),
                    Opis2 = doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "opis2")
                };

                pzgMaterialZasobuList.Add(pzgMaterialZasobu);

                XmlNodeList nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:pzg_cel", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuCelList.AddRange(from XmlNode node in nodeList
                        select new PzgCel
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:celArchiwalny", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuCelArchList.AddRange(from XmlNode node in nodeList
                        select new CelArchiwalny
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPrzed", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuDzialkaPrzedList.AddRange(from XmlNode node in nodeList
                        select new DzialkaPrzed
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPo", nsmgr);

                if (nodeList != null) 
                    pzgMaterialZasobuDzialkaPoList.AddRange(from XmlNode node in nodeList
                        select new DzialkaPo
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = node.InnerText
                        });

                PzgZgloszenie pzgZgloszenie = new PzgZgloszenie
                {
                    IdFile = idFile,
                    XmlPath = xmlFile,
                    IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                    PzgIdZgloszenia = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_idZgloszenia"),
                    IdZgloszeniaJedn = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaJedn"),
                    IdZgloszeniaNr = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaNr"),
                    IdZgloszeniaRok = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaRok"),
                    IdZgloszeniaEtap = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaEtap"),
                    IdZgloszeniaSepJednNr = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaSepJednNr"),
                    IdZgloszeniaSepNrRok = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "idZgloszeniaSepNrRok"),
                    PzgDataZgloszenia = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_dataZgloszenia"),
                    PzgPolozenieObszaru = string.Concat(doc.GetXmlValue(nsmgr, "PZG_MaterialZasobu", "pzg_polozenieObszaru").Take(32767)),
                    Obreb = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "obreb"),
                    PzgPodmiotZglaszajacyNazwa = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "nazwa"),
                    PzgPodmiotZglaszajacyRegon = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "REGON"),
                    PzgPodmiotZglaszajacyPesel = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_podmiotZglaszajacy", "PESEL"),
                    OsobaUprawniona = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "osobaUprawniona"),
                    PzgCel = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_cel"),
                    CelArchiwalny = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "celArchiwalny"),
                    PzgRodzaj = doc.GetXmlValue(nsmgr, "PZG_Zgloszenie", "pzg_rodzaj")
                };

                pzgZgloszenieList.Add(pzgZgloszenie);

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_Zgloszenie/xmls:pzg_cel", nsmgr);

                if (nodeList != null) 
                    pzgZgloszenieCelList.AddRange(from XmlNode node in nodeList
                        select new PzgCel
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = node.InnerText
                        });

                nodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_Zgloszenie/xmls:celArchiwalny", nsmgr);

                if (nodeList != null) 
                    pzgZgloszenieCelArchList.AddRange(from XmlNode node in nodeList
                        select new CelArchiwalny
                        {
                            IdFile = idFile,
                            XmlPath = xmlFile, 
                            IdMaterialu = Path.GetFileNameWithoutExtension(xmlFile),
                            Value = node.InnerText
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

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                Console.Write('\n');

                //  -------------------------------------------------------------------------------

                if (operatyExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [operaty]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("operaty");

                    sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    ExcelRange range = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                    sheet.Tables.Add(range, "operaty");

                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                if (operatyCelExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [operaty_cel]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("operaty_cel");

                    sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuCelList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                if (operatyCelArchExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [operaty_cel_arch]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("operaty_cel_arch");

                    sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuCelArchList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                if (operatyDzialkaPrzedExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [operaty_dzialka_przed]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("operaty_dzialka_przed");

                    sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuDzialkaPrzedList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                if (operatyDzialkaPoExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [operaty_dzialka_po]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("operaty_dzialka_po");

                    sheet.Cells[1, 1].LoadFromCollection(pzgMaterialZasobuDzialkaPoList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                if (zgloszeniaExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [zgłoszenia]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("zgłoszenia");

                    sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    ExcelRange range = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                    sheet.Tables.Add(range, "zgłoszenia");

                    sheet.Cells.AutoFitColumns(10, 50);
                }
                
                //  -------------------------------------------------------------------------------

                if (zgloszeniaCelExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [zgłoszenia_cel]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("zgłoszenia_cel");

                    sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieCelList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }
                
                //  -------------------------------------------------------------------------------

                if (zgloszeniaCelArchExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [zgłoszenia_cel_arch]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("zgłoszenia_cel_arch");

                    sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieCelArchList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                if (zgloszeniaOsobaUprawnionaExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [zgłoszenia_osoba_uprawniona]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("zgłoszenia_osoba_uprawniona");

                    sheet.Cells[1, 1].LoadFromCollection(pzgZgloszenieOsobaUprawnionaList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].Style.Numberformat.Format = "@";

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }

                //  -------------------------------------------------------------------------------

                if (walidacjaExport)
                {
                    Console.WriteLine(@"Eksport danych do XLS [błędy walidacji]...");

                    ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.Add("walidacja");

                    sheet.Cells[1, 1].LoadFromCollection(ErrorLogsList, true);

                    int rowsCount = sheet.Dimension.Rows;
                    int columnsCount = sheet.Dimension.Columns;

                    sheet.View.FreezePanes(2, 1);

                    sheet.Cells[1, 1, rowsCount, columnsCount].AutoFilter = true;
                    sheet.Cells.AutoFitColumns(10, 50);
                }

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


        private static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            XmlReader localSender = (XmlReader) sender;

            ErrorLog err = new ErrorLog
            {
                FileName = localSender.BaseURI,
                Element = localSender.LocalName,
                Line = e.Exception.LineNumber + ":" + e.Exception.LinePosition,
                Message = e.Message
            };

            ErrorLogsList.Add(err);
        }
    }
}
