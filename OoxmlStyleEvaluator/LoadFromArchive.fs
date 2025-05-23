module OoxmlStyleEvaluator.LoadFromArchive
open System.IO.Compression
open System.Xml.Linq

// Load document.xml root element
let loadDocumentXml (archive: ZipArchive): XElement =
    match archive.GetEntry("word/document.xml") |> Option.ofObj with
    | Some entry ->
        match XDocument.Load(entry.Open()).Root |> Option.ofObj with
        | Some root -> root
        | None -> failwith "word/document.xml does not have a root element."
    | None ->
        failwith "document.xml not found in the DOCX archive."

// Load styles.xml root element
let loadStyles  (archive: ZipArchive): XElement =
    match archive.GetEntry("word/styles.xml") |> Option.ofObj with
    | Some entry ->
        match XDocument.Load(entry.Open()).Root |> Option.ofObj with
        | Some root -> root
        | None -> failwith "word/styles.xml does not have a root element."
    | None ->
        failwith "styles.xml not found in the DOCX archive."

// Load numbering.xml root element (optional)
let loadNumberingOpt (archive: ZipArchive): XElement option =
    archive.GetEntry("word/numbering.xml")
    |> Option.ofObj
    |> Option.bind (fun entry ->
        XDocument.Load(entry.Open()).Root
        |> Option.ofObj
    )

