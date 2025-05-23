using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OoxmlStyleEvaluator;
using static OoxmlStyleEvaluator.ParagraphPropertyAccessors;

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

        var evaluator = new StyleEvaluator(archive);

        var paras = evaluator.DocumentRoot
            .Descendants(OoxmlStyleEvaluator.XmlHelpers.w + "p")
            .ToList() ;

        foreach (var para in paras)
        {
            if (evaluator.IsHeadingParagraph(para)) { 
                var level = evaluator.ParagraphStyleResolver
                var label = evaluator.GetHeadingNumberLabel(para);
            }

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
