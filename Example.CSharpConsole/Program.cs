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

        var evaluator = new StyleEvaluator(archive);

        var paras = evaluator.DocumentRoot
            .Descendants(OoxmlStyleEvaluator.XmlHelpers.w + "p")
            .ToList() ;

        foreach (var para in paras)
        {
            if (evaluator.IsHeadingParagraph(para))
            {
                int level = evaluator.GetHeadingLevel(para);
                switch (level)
                {
                    case -1:
                        Console.WriteLine("shouldn't happen");
                        break;
                    case int n when n >= 0 && n <= 9:
                        Console.WriteLine($"Level {level} Heading");
                        break;
                }

                Console.WriteLine($"Num level {evaluator.GetNumLevel(para)}");
                Console.WriteLine($"Num id {evaluator.GetNumId(para)}");
                Console.WriteLine($"{para.Value}");
            }
            else if (evaluator.IsBulletParagraph(para))
            {
                int numLevel = evaluator.GetNumLevel(para);
                switch (numLevel)
                {
                    case -1:
                        Console.WriteLine("shouldn't happen");
                        break;
                    case int n when n >= 0 && n <= 9:
                        Console.WriteLine($"Level {numLevel} bullet");
                        break;
                }

                Console.WriteLine($"Num id {evaluator.GetNumId(para)}");
                Console.WriteLine($"{para.Value}");
            }
        }

        return 0;
    }
}
