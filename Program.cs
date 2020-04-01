using CommandLine;
using DocumentFormat.OpenXml.Packaging;
using Jint;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using YamlDotNet.Serialization;

namespace doctemplate
{
    class Program
    {
        private static Dictionary<string, string> EnteredKeyValuePairs = new Dictionary<string, string>();

        static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed(ProcessDocx);
        }

        static void ProcessDocx(Options opts)
        {
            string templatePath, documentPath, documentText;
            LoadDocxFile(opts, out templatePath, out documentPath, out documentText);

            var regex = new Regex(@"!([^!<>]+)!");
            var values = regex.Matches(documentText).Select(match => match.Groups.Values.Last().Value);
            var valueSet = new HashSet<string>(values);

            IEnumerable<string> constValues = valueSet.Where(m => !m.StartsWith("=")); // all not formulas

            Engine jsEngine = new Engine();

            IEnumerable<string> loadedKeys = LoadValuesFromFileIfExist(ref documentText, templatePath, ref jsEngine);

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

            IEnumerable<string> keysToEnter = constValues.Except(loadedKeys);
            if (keysToEnter.Any())
            {
                Console.WriteLine("Enter template values:");
                foreach (var constant in keysToEnter)
                {
                    Console.Write($"{constant}=");
                    string value = Console.ReadLine();
                    EnteredKeyValuePairs.Add(constant, value);
                    jsEngine = jsEngine.SetValue(constant, value);
                    documentText = documentText.Replace($"!{constant}!", value);
                }
            }

            IEnumerable<string> funcValues = valueSet.Where(m => m.StartsWith("=")); // all not formulas
            foreach (var func in funcValues)
                ProcessFunction(ref documentText, regex, ref jsEngine, func);

            SaveDocxFile(templatePath, documentPath, documentText);

            if (opts.SaveYAML)
            {
                var serializer = new SerializerBuilder().Build();
                var yaml = serializer.Serialize(EnteredKeyValuePairs);
                File.WriteAllText(templatePath + ".yaml", yaml);
            }
        }

        private static void SaveDocxFile(string templatePath, string documentPath, string documentText)
        {
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

        private static void LoadDocxFile(Options opts, out string templatePath, out string documentPath, out string documentText)
        {
            templatePath = opts.TemplateFiles.First();
            documentPath = templatePath.Replace("_", "output\\");
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templatePath, false))
            {
                using (StreamReader reader = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    documentText = reader.ReadToEnd();
                }
            }
        }

        private static IEnumerable<string> LoadValuesFromFileIfExist(ref string documentText, string templatePath, ref Engine jsEngine)
        {
            string valuesFilePath = templatePath + ".yaml";
            if (!File.Exists(valuesFilePath))
                return new string[] { }; // return empty string array if no file
            
            string yaml = File.ReadAllText(valuesFilePath);

            var deserializer = new Deserializer();
            var yamlValues = deserializer.Deserialize<Dictionary<string, string>>(new StringReader(yaml));
            foreach (var pair in yamlValues)
            {
                jsEngine = jsEngine.SetValue(pair.Key, pair.Value);
                Console.WriteLine($"{pair.Key}={pair.Value}");
                documentText = documentText.Replace($"!{pair.Key}!", pair.Value);
            }
            return yamlValues.Keys;
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
                EnteredKeyValuePairs.Add(whatToDefine, value);

                jsEngine = jsEngine.SetValue(whatToDefine, value);

                ProcessFunction(ref documentText, regex, ref jsEngine, func); // call recursively
            }
        }
    }
}
