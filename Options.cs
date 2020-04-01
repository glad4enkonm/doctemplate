using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;

namespace doctemplate
{
    public class Options
    {
        [Option('t', "template", Required = true, HelpText = "Templates files to be processed.")]
        public IEnumerable<string> TemplateFiles { get; set; }

        [Option('s', "save", HelpText = "Save values to a yaml file.")]
        public bool SaveYAML { get; set; }
    }
}
