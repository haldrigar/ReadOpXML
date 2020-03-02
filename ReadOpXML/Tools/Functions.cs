using ReadOpXML.Schemat;

namespace ReadOpXML.Tools
{
    public static class Functions
    {
        public static string GetPzgOznMaterialuZasobu(PzgMaterialZasobu pzgMaterial)
        {
            string wynik = string.Empty;

            wynik = pzgMaterial.OznMaterialuZasobuTyp == "--brak wartosci--" ? string.Empty : wynik + pzgMaterial.OznMaterialuZasobuTyp + ".";
            wynik = pzgMaterial.OznMaterialuZasobuJedn == "--brak wartosci--" ? wynik : wynik + pzgMaterial.OznMaterialuZasobuJedn + pzgMaterial.OznMaterialuZasobuSepJednNr;
            wynik = wynik + pzgMaterial.OznMaterialuZasobuNr + pzgMaterial.OznMaterialuZasobuSepNrRok + pzgMaterial.OznMaterialuZasobuRok;

            if (pzgMaterial.OznMaterialuZasobuTom != "--brak wartosci--") wynik = wynik + " t." + pzgMaterial.OznMaterialuZasobuTom;

            return wynik;
        }

        public static string GetPzgIdZgloszenia(PzgZgloszenie pzgZgloszenie)
        {
            string wynik = string.Empty;

            wynik = pzgZgloszenie.IdZgloszeniaJedn == "--brak wartosci--" ? wynik : wynik + pzgZgloszenie.IdZgloszeniaJedn + pzgZgloszenie.IdZgloszeniaSepJednNr;
            wynik = wynik + pzgZgloszenie.IdZgloszeniaNr + pzgZgloszenie.IdZgloszeniaSepNrRok + pzgZgloszenie.IdZgloszeniaRok;

            if (pzgZgloszenie.IdZgloszeniaEtap != "--brak wartosci--") wynik = wynik + " et." + pzgZgloszenie.IdZgloszeniaEtap;

            return wynik;
        }
    }
}
