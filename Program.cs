using System;
using System.Collections.Generic;
using System.IO;
using CommandLine;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;

namespace MyTools.CLI
{
    public class Program
    {
        private static string _licenseKey;

        [Verb("section", isDefault: false, HelpText = "Get a few pages from a PDF document as a new PDF document.")]
        public class SectionOptions
        {
            [Option('p', "path", Required = true, HelpText = "FilePath")]
            public string FilePath { get; set; }

            [Option('f', "from", Required = true, HelpText = "Starting page nummer.")]
            public int FromPage { get; set; }

            [Option('t', "till", Required = true, HelpText = "Ending page nummer.")]
            public int TillPage { get; set; }
        }

        static void Main(string[] args)
        {            
            Console.WriteLine("Welcome to PDF tools v0.0.1");

            var result = Parser.Default.ParseArguments<SectionOptions>(args)
                .MapResult((SectionOptions opts) => RunSection(opts), HandleParseError);

            Console.WriteLine(result);
        }

        static string RunSection(SectionOptions opts)
        {
            //handle options

            ensureSyncfusionLicense();

            var destinationFilePath = opts.FilePath.Replace(".pdf", "_output.pdf");
            using (var fileStream = new FileStream(opts.FilePath, FileMode.Open, FileAccess.Read))
            using (var loadedDocument = new PdfLoadedDocument(fileStream))
            {
                var document = new PdfDocument();

                for (int i = 0; i < loadedDocument.PageCount; i++)
                {
                    var pageNr = i + 1;
                    if (pageNr >= opts.FromPage && pageNr <= opts.TillPage)
                    {
                        document.ImportPage(loadedDocument, i);
                    }
                }

                using (var stream = new MemoryStream())
                using (var outputFileStream = new FileStream(destinationFilePath, FileMode.Create, FileAccess.Write))
                {
                    document.Save(stream);
                    stream.Position = 0;
                    document.Close(true);
                    byte[] bytes = stream.ToArray();
                    outputFileStream.Write(bytes, 0, (int)bytes.Length);

                }
            }

            return $"Saved the output file to {destinationFilePath}";
        }

        private static void ensureSyncfusionLicense()
        {
            _licenseKey = Environment.GetEnvironmentVariable("SyncfusionLicenseKey", EnvironmentVariableTarget.User);

            if (_licenseKey == null)
            {
                Console.WriteLine("Paste your Syncfusion License key");
                Environment.SetEnvironmentVariable("SyncfusionLicenseKey", Console.ReadLine(), EnvironmentVariableTarget.User);
            }

            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(_licenseKey);
        }

        static string HandleParseError(IEnumerable<Error> errs)
        {
            //handle errors
            return "error!";
        }
    }
}
