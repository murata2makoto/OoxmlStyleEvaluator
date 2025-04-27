module OoxmlStyleEvaluator.ParagraphStyle.List

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.ParagraphStyle.Core
open OoxmlStyleEvaluator.ParagraphStyle.Heading

/// <summary>
/// Determines whether a paragraph is part of a list (bulleted or numbered),
/// by checking both direct formatting (w:numPr) and style-based definitions.
/// </summary>
let isBulletParagraph (para: XElement) (stylesRoot: XElement) : bool =
    let direct =
        para
        |> tryElement (w + "pPr")
        |> Option.bind (tryElement (w + "numPr"))
        |> Option.isSome

    let fromStyle =
        tryGetParagraphStyleId para
        |> Option.bind (fun styleId -> findParagraphStyle styleId stylesRoot)
        |> Option.bind (tryElement (w + "pPr"))
        |> Option.bind (tryElement (w + "numPr"))
        |> Option.isSome

    direct || fromStyle

/// <summary>
/// Gets the list nesting level (ilvl) of a paragraph, if numbered or bulleted.
/// </summary>
let tryGetListLevel (para: XElement) : int option =
    tryElement (w + "pPr")  para
    |> Option.bind (tryElement (w + "numPr"))
    |> Option.bind (tryElement (w + "ilvl"))
    |> Option.bind (tryAttrValue (w + "val"))
    |> Option.bind (fun s ->
        match System.Int32.TryParse(s) with
        | true, i -> Some i
        | _ -> None)

/// <summary>
/// Gets detailed list information for a paragraph, including numId, ilvl, label text, and bullet type.
/// Returns None if numberingRoot is null.
/// </summary>
let tryGetListInfo (para: XElement) (stylesRoot: XElement) (numberingRoot: XElement) : (string * int * string * bool) option =
    if obj.ReferenceEquals(numberingRoot, null) then None
    else
        let getNumPrFromParaOrStyle para =
            let direct = 
                para
                |> tryElement (w + "pPr") 
                |> Option.bind (tryElement (w + "numPr"))

            let fromStyle =
                tryGetParagraphStyleId para
                |> Option.bind (fun styleId -> findParagraphStyle styleId stylesRoot)
                |> Option.bind (tryElement (w + "pPr"))
                |> Option.bind (tryElement (w + "numPr"))

            direct |> Option.orElse fromStyle

        match getNumPrFromParaOrStyle para with
        | Some numPr ->
            let numIdOpt =
                numPr
                |> tryElement (w + "numId")
                |> Option.bind (tryAttrValue (w + "val"))

            let ilvlOpt =
                numPr
                |> tryElement (w + "ilvl")
                |> Option.bind (tryAttrValue (w + "val"))
                |> Option.bind (fun v -> match System.Int32.TryParse(v) with true, i -> Some i | _ -> None)

            match numIdOpt, ilvlOpt with
            | Some numId, Some ilvl ->
                let num =
                    numberingRoot.Elements(w + "num")
                    |> Seq.tryFind (fun n -> tryAttrValue (w + "numId") n = Some numId)

                let abstractNumId =
                    num
                    |> Option.bind (tryElement (w + "abstractNumId"))
                    |> Option.bind (tryAttrValue (w + "val"))

                let abstractNum =
                    abstractNumId
                    |> Option.bind (fun aid ->
                        numberingRoot.Elements(w + "abstractNum")
                        |> Seq.tryFind 
                            (fun an -> tryAttrValue (w + "abstractNumId") an = Some aid))

                let lvl =
                    abstractNum
                    |> Option.bind (fun an ->
                        an.Elements(w + "lvl")
                        |> Seq.tryFind (fun l -> tryAttrValue (w + "ilvl") l = Some (string ilvl)))

                let lvlText =
                    lvl
                    |> Option.bind (tryElement (w + "lvlText"))
                    |> Option.bind (tryAttrValue (w + "val"))

                let numFmt =
                    lvl
                    |> Option.bind (tryElement (w + "numFmt"))
                    |> Option.bind (tryAttrValue (w + "val"))

                let isBullet = (numFmt |> Option.exists (fun v -> v = "bullet"))
                match lvlText with
                | Some label -> Some (numId, ilvl, label, isBullet)
                | None -> None
            | _ -> None
        | None -> None

/// <summary>
/// Given a paragraph and the list of all paragraphs, generates the resolved list label (e.g. "1.2." or "●").
/// Returns None if numberingRoot is null.
/// </summary>
let generateResolvedListLabel
    (para: XElement)
    (stylesRoot: XElement)
    (numberingRoot: XElement)
    (paragraphs: XElement list) : string option =
    if obj.ReferenceEquals(numberingRoot, null) then None
    else
        match tryGetListInfo para stylesRoot numberingRoot with
        | Some (numId, ilvl, template, isBullet) when not isBullet ->
            let numberedParas =
                paragraphs
                |> List.choose (fun p ->
                    match tryGetListInfo p stylesRoot numberingRoot with
                    | Some (nid, lvl, _, false) when nid = numId -> Some (p, lvl)
                    | _ -> None)

            let counters = System.Collections.Generic.Dictionary<int, int>()
            let levelMap = System.Collections.Generic.Dictionary<XElement, int list>()

            for (p, lvl) in numberedParas do
                for i in 0 .. lvl do
                    if not (counters.ContainsKey i) then counters.[i] <- 0
                counters.[lvl] <- counters.[lvl] + 1
                for i in lvl + 1 .. counters.Count - 1 do
                    counters.[i] <- 0
                let current = [ for i in 0 .. lvl do yield counters.[i] ]
                levelMap.[p] <- current

            match levelMap.TryGetValue para with
            | true, nums -> Some (expandLabelTemplate template nums)
            | _ -> None

        | Some (_, _, template, true) ->
            Some template

        | _ -> None
