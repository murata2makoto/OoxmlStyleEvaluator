open System
open System.IO
open System.IO.Compression
open System.Xml.Linq
open OoxmlStyleEvaluator

[<EntryPoint>]
let main argv =
    System.Console.OutputEncoding <- System.Text.Encoding.UTF8
    if argv.Length <> 1 then
        Console.Error.WriteLine("Usage: Example.FSharpConsole <document.docx>")
        1
    else
        let docxPath = argv.[0]

        if not (File.Exists(docxPath)) then
            Console.Error.WriteLine($"File not found: {docxPath}")
            1
        else
            use archive = ZipFile.OpenRead(docxPath)
            let documentXml =
                archive.GetEntry("word/document.xml")
                |> Option.ofObj
                |> Option.map (fun entry -> XDocument.Load(entry.Open()))
                |> Option.defaultWith (fun () -> failwith "document.xml not found in the DOCX file.")

            let evaluator = StyleEvaluator(archive, documentXml)

            let paras =
                match documentXml.Root with
                | null -> []
                | root -> root.Descendants(XmlHelpers.w + "p") |> Seq.toList

            for para in paras do
                if evaluator.IsHeadingParagraph(para) then
                    let levelOpt = evaluator.GetHeadingLevel(para)
                    let labelOpt = evaluator.GetHeadingNumberLabel(para)
                    match levelOpt, labelOpt with
                    | level, "" when level > 0 ->
                        Console.WriteLine($"Heading (Level {level}): (no label)")
                    |  level,  label when level >= 0 ->
                        Console.WriteLine($"Heading (Level {level}): {label}")
                    | _ -> ()
                elif evaluator.IsBulletParagraph(para) then
                    let levelOpt = evaluator.GetBulletLevel(para)
                    let labelOpt = evaluator.GetBulletLabel(para)
                    match levelOpt, labelOpt with
                    | level, "" when level > 0 ->
                        Console.WriteLine($"Bullet (Level {level}): (no label)")
                    | level, label  when level >= 0->
                        Console.WriteLine($"Bullet (Level {level}): {label}")
                    | _ -> ()

            0
