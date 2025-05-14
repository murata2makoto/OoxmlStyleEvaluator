using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OoxmlStyleEvaluator;

class Program
{
    static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        if (args.Length != 1)
        {
            Console.Error.WriteLine("Usage: Example.CSharpConsole <document.docx>");
            return 1;
        }

        string docxPath = args[0];
        if (!File.Exists(docxPath))
        {
            Console.Error.WriteLine($"File not found: {docxPath}");
            return 1;
        }

        using var archive = ZipFile.OpenRead(docxPath);
        var documentEntry = archive.GetEntry("word/document.xml")
            ?? throw new Exception("document.xml not found in the DOCX file.");

        XDocument documentXml;
        using (var stream = documentEntry.Open())
        {
            documentXml = XDocument.Load(stream);
        }
        var evaluator = new StyleEvaluator(archive, documentXml);

        var paras = documentXml.Root?
            .Descendants(OoxmlStyleEvaluator.XmlHelpers.w + "p")
            .ToList() ?? new();

        foreach (var para in paras)
        {
            var level = evaluator.GetHeadingLevel(para);
            var label = evaluator.GetHeadingNumberLabel(para);

            if (level >= 0 && label != "")
            {
                Console.WriteLine($"Heading (Level {level}): {label}");
            }
            else if (level >= 0)
            {
                Console.WriteLine($"Heading (Level {level}): (no label)");
            }
        }

        return 0;
    }
}
