module OoxmlStyleEvaluator.ParagraphStyle.Heading

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.ParagraphStyle.Core

/// <summary>
/// Expands a list label template such as "%1.%2." using the given list of level numbers.
/// </summary>
let expandLabelTemplate (template: string) (levelNumbers: int list) : string =
    let regex = System.Text.RegularExpressions.Regex(@"%(\d+)")
    regex.Replace(template, fun (m: System.Text.RegularExpressions.Match) ->
        let idx = int m.Groups.[1].Value - 1
        if idx < levelNumbers.Length then
            string levelNumbers.[idx]
        else
            ""
    )

/// <summary>
/// Gets the auto-generated heading number template (e.g. "%1.", "%1.%2.") if the paragraph is a heading.
/// Returns None if numberingRoot is null.
/// </summary>
let     tryGetHeadingNumberTemplate (para: XElement) (stylesRoot: XElement) (numberingRoot: XElement) : string option =
    if isNull numberingRoot then None
    else
        match tryGetParagraphStyleId para with
        | Some styleId ->
            match findParagraphStyle styleId stylesRoot with
            | Some style ->
                let outlineLvl =
                    style
                    |> tryElement (w + "pPr")
                    |> Option.bind (tryElement (w + "outlineLvl"))
                    |> Option.bind (tryAttrValue (w + "val"))
                    |> Option.bind (fun v -> match System.Int32.TryParse(v) with true, i -> Some i | _ -> None)

                let numId =
                    style
                    |> tryElement (w + "pPr")
                    |> Option.bind (tryElement (w + "numPr"))
                    |> Option.bind (tryElement (w + "numId"))
                    |> Option.bind (tryAttrValue (w + "val"))

                match outlineLvl, numId with
                | Some ilvl, Some nid ->
                    let num =
                        numberingRoot.Elements(w + "num")
                        |> Seq.tryFind (fun n -> tryAttrValue (w + "numId") n = Some nid)

                    let abstractNumId =
                        num
                        |> Option.bind (tryElement (w + "abstractNumId"))
                        |> Option.bind (tryAttrValue (w + "val"))

                    let abstractNum =
                        abstractNumId
                        |> Option.bind (fun aid ->
                            numberingRoot.Elements(w + "abstractNum")
                            |> Seq.tryFind 
                                (fun an -> 
                                   tryAttrValue (w + "abstractNumId") an = Some aid))

                    let lvlText =
                        abstractNum
                        |> Option.bind (fun an ->
                            an.Elements(w + "lvl")
                            |> Seq.tryFind (fun lvl -> tryAttrValue (w + "ilvl") lvl = Some (string ilvl)))
                        |> Option.bind (tryElement (w + "lvlText"))
                        |> Option.bind (tryAttrValue (w + "val"))

                    lvlText
                | _ -> None
            | None -> None
        | None -> None

/// <summary>
/// For a list of paragraphs, returns (para, label option) pairs.
/// </summary>
let generateExpandedHeadingLabels (paras: XElement list) (stylesRoot: XElement) (numberingRoot: XElement)
    : (XElement * string option) list =
    let templatePairs =
        paras
        |> List.choose (fun p -> tryGetHeadingNumberTemplate p stylesRoot numberingRoot |> Option.map (fun t -> p, t))

    let mutable levelStack = System.Collections.Generic.Dictionary<int, int>()

    paras
    |> List.map (fun p ->
        match templatePairs |> List.tryFind (fun (x, _) -> obj.ReferenceEquals(x, p)) with
        | Some (_, template) ->
            let level =
                try
                    let styleIdOpt = tryGetParagraphStyleId p
                    let styleOpt = styleIdOpt |> Option.bind (fun id -> findParagraphStyle id stylesRoot)
                    let outlineLvl =
                        styleOpt
                        |> Option.bind (tryElement (w + "pPr"))
                        |> Option.bind (tryElement (w + "outlineLvl"))
                        |> Option.bind (tryAttrValue (w + "val"))
                        |> Option.bind (fun v -> match System.Int32.TryParse(v) with true, i -> Some i | _ -> None)
                    outlineLvl |> Option.defaultValue 0
                with _ -> 0

            for i in 0 .. level do
                if not (levelStack.ContainsKey i) then levelStack.[i] <- 0
            levelStack.[level] <- levelStack.[level] + 1
            for i in level + 1 .. 8 do
                levelStack.Remove(i) |> ignore
            let current = [ for i in 0 .. level do yield levelStack.[i] ]
            p, Some (expandLabelTemplate template current)
        | None -> p, None
    )

/// Determines whether a paragraph is a heading by checking numbering linkage
let isHeadingParagraph (para: XElement) (stylesRoot: XElement) : bool =
    match tryGetParagraphStyleId para with
    | Some styleId ->
        let style =
            stylesRoot.Elements(w + "style")
            |> Seq.tryFind (fun s ->
                let sid = s.Attribute(w + "styleId")
                let typ = s.Attribute(w + "type")
                sid <> null && typ <> null &&
                sid.Value = styleId && typ.Value = "paragraph")
        match style with
        | Some s -> s.Descendants(w + "outlineLvl") |> Seq.exists (fun _ -> true)
        | None -> false
    | None -> false