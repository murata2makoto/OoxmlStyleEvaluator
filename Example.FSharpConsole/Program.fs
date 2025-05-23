open System
open System.IO
open System.IO.Compression
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

            let evaluator = StyleEvaluator(archive)
            let documentRoot = evaluator.DocumentRoot

            let paras =
                documentRoot.Descendants(XmlHelpers.w + "p") 

            for para in paras do
                if evaluator.IsHeadingParagraph(para) then 
                    let level = evaluator.GetHeadingLevel(para)
                    match level with
                    | -1  ->
                        Console.WriteLine("shouldn't happen")
                    | 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 ->
                        Console.WriteLine($"Level {level} Heading")
                    | _ -> ()
                    evaluator.GetNumLevel(para) |> printfn "Num level %d"
                    evaluator.GetNumId(para) |> printfn "Num id %d"
                    printfn "%s" para.Value
                elif evaluator.IsBulletParagraph(para) then
                    let numLevel = evaluator.GetNumLevel(para)
                    match numLevel with
                    | -1  ->
                        Console.WriteLine("shouldn't happen")
                    | 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 ->
                        Console.WriteLine($"Level {numLevel} bullet")
                    | _ -> ()
                    evaluator.GetNumId(para) |> printfn "Num id %d"
                    printfn "%s" para.Value

            0
