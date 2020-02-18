using CommandLine;

namespace ReadOpXML.Tools
{
    public class Options
    {
        [Option('s', "startupPath", Required = true, HelpText = "Katalog z danymi.")]
        public string StarupPath { get; set; }

        [Option('p', "poprawa", Required = false, HelpText = "Wykonaj automatyczną poprawę plików.")]
        public bool Poprawa { get; set; }

        [Option('w', "walidacja", Required = false, HelpText = "Wykonaj walidację plików zgodnie ze schematem XSD.")]
        public bool Walidacja { get; set; }
    }
}
