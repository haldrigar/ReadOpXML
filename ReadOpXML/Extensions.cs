using System;
using System.Linq;
using System.Xml;

namespace ReadOpXML
{
    public static class Extensions
    {
        public static string GetXmlValue(this XmlDocument doc, XmlNamespaceManager nsmgr, params string[] nodes)
        {
            string xPath = nodes.Aggregate("/xmls:schema", (current, t) => current + "/xmls:" + t);

            XmlNodeList nodeList = doc.DocumentElement?.SelectNodes(xPath, nsmgr);

            if (nodeList == null || nodeList.Count == 0)
            {
                return "--brak znacznika--";
            }

            string values = "";

            if (nodes.Contains("osobaUprawniona"))
            {
                foreach (XmlNode node in nodeList)
                {
                    XmlDocument docosoba = new XmlDocument();

                    docosoba.LoadXml(node.OuterXml);

                    values = values + "|" + docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:imie", nsmgr)?.InnerText + "_" + 
                             docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:nazwisko", nsmgr)?.InnerText + "_" + 
                             docosoba.SelectSingleNode("/xmls:osobaUprawniona/xmls:numer_uprawnien", nsmgr)?.InnerText;
                }
            }
            else
            {
                foreach (XmlNode node in nodeList)
                {
                    if (node.InnerText.Contains("|")) throw new Exception("Błąd separatora!");

                    values = values + "|" + node.InnerText;
                }
            }

            return values.TrimStart('|').Trim();
        }
    }
}
