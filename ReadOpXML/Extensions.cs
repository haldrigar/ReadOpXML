using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace ReadOpXML
{
    public static class Extensions
    {
        public static void FixErrors(this XmlDocument doc)
        {
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("xmls", "http://www.w3.org/2001/XMLSchema");
            
            int rok = 0;

            string fileName = Path.GetFileName(doc.BaseURI);

            if (Regex.IsMatch(fileName.ToUpper(), @"^P\.[0-9]{4}\.[0-9]{4}\.[0-9]+\.XML$"))
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

            if (xmlNodeList != null)
                foreach (XmlNode node in xmlNodeList)
                {
                    node.InnerText = node.InnerText.Replace(" ", "");
                }

            //  ---------------------------------------------------------------------------
            //  POPRAWA spacji w numerze działki PO
            
            xmlNodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPo", nsmgr);

            if (xmlNodeList != null)
                foreach (XmlNode node in xmlNodeList)
                {
                    node.InnerText = node.InnerText.Replace(" ", "");
                }

            //  ---------------------------------------------------------------------------
            //  POPRAWA numeru w dzialkaPrzed
            
            xmlNodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPrzed", nsmgr);

            if (xmlNodeList != null)
                foreach (XmlNode node in xmlNodeList)
                {
                    // jedna litera na końcu numeru działki 040903_5.0006.871a
                    if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+[a-z]{1}$"))
                    {
                        node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" +
                                         node.InnerText.Substring(node.InnerText.Length - 1, 1);
                    }

                    // jedna litera na końcu numeru działki 040903_5.0006.871/11a
                    if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+\/[0-9]+[a-z]{1}$"))
                    {
                        node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" +
                                         node.InnerText.Substring(node.InnerText.Length - 1, 1);
                    }
                }

            //  ---------------------------------------------------------------------------
            //  POPRAWA numeru w dzialkaPo
            
            xmlNodeList = doc.DocumentElement?.SelectNodes("/xmls:schema/xmls:PZG_MaterialZasobu/xmls:dzialkaPo", nsmgr);

            if (xmlNodeList != null)
                foreach (XmlNode node in xmlNodeList)
                {
                    // jedna litera na końcu numeru działki 040903_5.0006.871a
                    if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+[a-z]{1}$"))
                    {
                        node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" +
                                         node.InnerText.Substring(node.InnerText.Length - 1, 1);
                    }

                    // jedna litera na końcu numeru działki 040903_5.0006.871/11a
                    if (Regex.IsMatch(node.InnerText, @"^[0-9]{6}_.\.[0-9]{4}\.[0-9]+\/[0-9]+[a-z]{1}$"))
                    {
                        node.InnerText = node.InnerText.Substring(0, node.InnerText.Length - 1) + "-" +
                                         node.InnerText.Substring(node.InnerText.Length - 1, 1);
                    }
                }

            //  ---------------------------------------------------------------------------
        }
        
        public static string GetXmlValue(this XmlDocument doc, XmlNamespaceManager nsmgr, params string[] nodes)
        {
            string xPath = nodes.Aggregate("/xmls:schema", (current, t) => current + "/xmls:" + t);

            XmlNodeList nodeList = doc.DocumentElement?.SelectNodes(xPath, nsmgr);

            if (nodeList == null || nodeList.Count == 0)
            {
                return "--brak znacznika--";
            }

            List<string> valuesList = new List<string>();

            if (nodes.Contains("osobaUprawniona"))
            {
                foreach (XmlNode node in nodeList)
                {
                    XmlDocument docosoba = new XmlDocument();

                    docosoba.LoadXml(node.OuterXml);

                    string imie = docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:imie", nsmgr)?.InnerText;
                    string nazwisko = docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:nazwisko", nsmgr)?.InnerText;
                    string numerUprawnien = docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:numer_uprawnien", nsmgr)?.InnerText;

                    if (string.IsNullOrEmpty(imie)) imie = "--brak wartosci--";
                    if (string.IsNullOrEmpty(nazwisko)) nazwisko = "--brak wartosci--";
                    if (string.IsNullOrEmpty(numerUprawnien)) numerUprawnien = "--brak wartosci--";

                    if (imie.Contains('|') || nazwisko.Contains('|') ||  numerUprawnien.Contains('|')) throw new Exception("Błąd separatora!");

                    valuesList.Add(imie + "_" + nazwisko + "_" + numerUprawnien);
                }
            }
            else
            {
                foreach (XmlNode node in nodeList)
                {
                    string wartosc = node.InnerText;

                    if (string.IsNullOrEmpty(wartosc)) wartosc = "--brak wartosci--";

                    if (node.InnerText.Contains("|")) throw new Exception("Błąd separatora!");

                    valuesList.Add(wartosc);
                }
            }

            valuesList.Sort();

            return valuesList.Aggregate(string.Empty, (current, value) => current + '|' + value).TrimStart('|');
        }
    }
}
