namespace ReadOpXML.Tools
{
    public static class Functions
    {
        public static string GetOperatZgloszenieName(string jedn, string nr, string rok, string tomEtap, bool isOperat)
        {
            if (jedn == "--brak wartosci--") jedn = string.Empty;
            if (nr == "--brak wartosci--") nr = string.Empty;
            if (rok == "--brak wartosci--") rok = string.Empty;
            if (tomEtap == "--brak wartosci--") tomEtap = string.Empty;

            char sepJednNr;
            char sepNrRok;

            bool status = int.TryParse(rok, out int rokInt);

            if (!status) return "--błąd w roku--";

            if (rokInt <=2013)
            {
                sepJednNr = '-';
                sepNrRok = '/';
            }
            else
            {
                sepJednNr = '.';
                sepNrRok = '.';
            }

            if (!string.IsNullOrEmpty(jedn) && !string.IsNullOrEmpty(nr) && !string.IsNullOrEmpty(rok) && !string.IsNullOrEmpty(tomEtap))
            {
                return jedn + sepJednNr + nr + sepNrRok + rok + (isOperat ? " t." : " e.") + tomEtap;
            }

            if (!string.IsNullOrEmpty(jedn) && !string.IsNullOrEmpty(nr) && !string.IsNullOrEmpty(rok) && string.IsNullOrEmpty(tomEtap))
            {
                return jedn + sepJednNr + nr + sepNrRok + rok;
            }

            if (string.IsNullOrEmpty(jedn) && !string.IsNullOrEmpty(nr) && !string.IsNullOrEmpty(rok) && !string.IsNullOrEmpty(tomEtap))
            {
                return nr + sepNrRok + rok + (isOperat ? " t." : " e.") + tomEtap;
            }

            if (string.IsNullOrEmpty(jedn) && !string.IsNullOrEmpty(nr) && !string.IsNullOrEmpty(rok) && string.IsNullOrEmpty(tomEtap))
            {
                return nr + sepNrRok + rok;
            }

            return string.Empty;
        }
    }
}
