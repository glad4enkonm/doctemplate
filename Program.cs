using CommandLine;
using DocumentFormat.OpenXml.Packaging;
using Jint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace doctemplate
{
    class Program
    {
        public class Options
        {
            [Option('t', "template", Required = true, HelpText = "Templates files to be processed.")]
            public IEnumerable<string> TemplateFiles { get; set; }
        }

        static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed(ProcessDocx);
        }

        static void ProcessDocx(Options opts)
        {

            string templatePath = opts.TemplateFiles.First();
            string documentPath = templatePath.Replace("_", "output\\");
            string documentText;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templatePath, false))
            {
                using (StreamReader reader = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    documentText = reader.ReadToEnd();
                }
            }

            var regex = new Regex(@"!([^!<>]+)!");
            var values = regex.Matches(documentText).Select(match => match.Groups.Values.Last().Value);
            var valueSet = new HashSet<string>(values);

            IEnumerable<string> constValues = valueSet.Where(m => !m.StartsWith("=")); // all not formulas

            Engine jsEngine = new Engine();
            jsEngine = jsEngine.Execute( // define function today
                @"function today(days) {
                    const daysToAdd = isNaN(days) ? 0 : days;
                    var date = new Date();
                    date.setDate(date.getDate() + daysToAdd);
                    return date.toLocaleDateString('de-DE');
                }");

            jsEngine = jsEngine.Execute( // define function EUR
                @"function $(amount) {
                    return amount.toLocaleString('de - DE', {minimumFractionDigits: 2});
                }");

            Console.WriteLine("Enter template values:");
            foreach (var constant in constValues)
            {
                Console.Write($"{constant}=");
                string value = Console.ReadLine();
                jsEngine = jsEngine.SetValue(constant, value);
                documentText = documentText.Replace($"!{constant}!", value);
            }

            IEnumerable<string> funcValues = valueSet.Where(m => m.StartsWith("=")); // all not formulas
            foreach (var func in funcValues)
            {
                ProcessFunction(ref documentText, regex, ref jsEngine, func);
            }


            // Create output directory and files
            FileInfo fileInfo = new FileInfo(templatePath);
            new FileInfo(documentPath).Directory.Create(); // create path
            fileInfo.CopyTo(documentPath, true);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(documentPath, true))
            {
                using (StreamWriter writer = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    writer.Write(documentText);
                }
            }
        }

        private static void ProcessFunction(ref string documentText, Regex regex, ref Engine jsEngine, string func)
        {
            try
            {
                string code = func.StartsWith("=") ? func.Substring(1) : func; // remove leading = sign
                string value = jsEngine.Execute(code).GetCompletionValue().ToObject().ToString();
                documentText = documentText.Replace($"!{func}!", value);
            }
            catch (Exception ex)
            {
                // E.g. amount is not defined
                var notDefinedRegex = new Regex(@"(.+) is not defined");
                var match = notDefinedRegex.Match(ex.Message);

                if (!match.Success) throw;

                var whatToDefine = match.Groups.Values.Last().Value;

                Console.Write($"{whatToDefine}=");
                string value = Console.ReadLine();
                jsEngine = jsEngine.SetValue(whatToDefine, value);

                ProcessFunction(ref documentText, regex, ref jsEngine, func); // call recursively
            }
        }
    }
}
