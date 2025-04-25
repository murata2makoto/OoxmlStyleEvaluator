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
            var level = evaluator.GetHeadingLevelNullable(para);
            var label = evaluator.GetHeadingNumberLabelNullable(para);

            if (level.HasValue && label != null)
            {
                Console.WriteLine($"Heading (Level {level.Value}): {label}");
            }
            else if (level.HasValue)
            {
                Console.WriteLine($"Heading (Level {level.Value}): (no label)");
            }
            else
            {
                var bulletLevel = evaluator.GetBulletLevelNullable(para);
                var bulletLabel = evaluator.GetBulletLabelNullable(para);

                if (bulletLevel.HasValue && bulletLabel != null)
                {
                    Console.WriteLine($"Bullet (Level {bulletLevel.Value}): {bulletLabel}");
                }
                else if (bulletLevel.HasValue)
                {
                    Console.WriteLine($"Bullet (Level {bulletLevel.Value}): (no label)");
                }
            }
        }

        return 0;
    }
}
