using CommandLine;

namespace ReadOpXML
{
    public class Options
    {
        [Option('s', "startupPath", Required = true, HelpText = "Katalog z danymi.")]
        public string StarupPath { get; set; }

        [Option('p', "poprawa", Required = false, Default = false, HelpText = "Czy wykonać automatyczną poprawę plików.")]
        public bool Poprawa { get; set; }

        [Option('w', "walidacja", Required = false, Default = false, HelpText = "Czy wykonać walidację plików.")]
        public bool Walidacja { get; set; }
    }
}
