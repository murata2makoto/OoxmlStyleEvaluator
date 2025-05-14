module OoxmlStyleEvaluator.ParagraphStyle.Detection

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.ParagraphStyleEvaluator

/// <summary>
/// Determines whether a given paragraph is a heading,
/// based on the presence of an effective outline level (outlineLvl).
/// Fully respects style inheritance and document defaults.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for paragraph properties.</param>
/// <param name="tableStylePPr">Optional paragraph properties derived from the table style, if applicable.</param>
/// <returns>True if the paragraph is a heading; otherwise, false.</returns>
let isHeadingParagraph
    (para: XElement)
    (stylesRoot: XElement)
    (docDefaults: PPr)
    (tableStylePPr: PPr option) : bool =

    evaluateEffectiveParagraphProperty "outlineLvl" para stylesRoot docDefaults tableStylePPr
    |> Option.isSome

/// <summary>
/// Retrieves the heading level (outlineLvl) of a given paragraph,
/// fully respecting style inheritance and document defaults.
/// </summary>
/// <param name="para">The paragraph (w:p) element to evaluate.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for paragraph properties.</param>
/// <param name="tableStylePPr">Optional paragraph properties derived from the table style, if applicable.</param>
/// <returns>The heading level as an integer (0-8) if specified; otherwise, None.</returns>
let getHeadingLevel
    (para: XElement)
    (stylesRoot: XElement)
    (docDefaults: PPr)
    (tableStylePPr: PPr option) : int option =

    evaluateEffectiveParagraphProperty "outlineLvl" para stylesRoot docDefaults tableStylePPr
    |> Option.bind (fun e ->
        tryAttrValue (w + "val") e
        |> Option.bind (fun s ->
            match System.Int32.TryParse(s) with
            | true, value -> Some value
            | false, _ -> None))

/// <summary>
/// Generates hierarchical heading labels (e.g., "1.", "1.1.", "1.1.1.")
/// for a given list of paragraphs, based on their heading levels.
/// Non-heading paragraphs are assigned None.
/// </summary>
/// <param name="paras">The list of paragraph (w:p) elements.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for paragraph properties.</param>
/// <param name="tableStylePPrs">A list of optional table style paragraph properties corresponding to the paragraphs.</param>
/// <returns>A list of option strings: Some(label) for headings, None for non-headings.</returns>
let generateExpandedHeadingLabels
    (paras: XElement list)
    (stylesRoot: XElement)
    (docDefaults: PPr)
    (tableStylePPrs: (PPr option) list)
    : string option list =

    let rec updateCounters (counters: int list) (level: int) : int list =
        // Extend counters list if needed
        let counters =
            if counters.Length <= level then
                counters @ List.replicate (level - counters.Length + 1) 0
            else counters
        // Increment the counter at the current level, reset lower levels
        counters
        |> List.mapi (fun idx v ->
            if idx < level then v
            elif idx = level then v + 1
            else 0)

    let buildLabel (counters: int list) (level: int) : string =
        counters
        |> List.take (level + 1)
        |> List.map string
        |> String.concat "."
        |> fun s -> s + "."

    // Main loop
    let rec loop (paras: XElement list) (tableStylePPrs: (PPr option) list) (counters: int list) (acc: string option list) =
        match paras, tableStylePPrs with
        | [], [] -> List.rev acc
        | para::restParas, tPPrOpt::restTableStylePPrs ->
            if isHeadingParagraph para stylesRoot docDefaults tPPrOpt then
                match getHeadingLevel para stylesRoot docDefaults tPPrOpt with
                | Some level ->
                    let counters' = updateCounters counters level
                    let label = buildLabel counters' level
                    loop restParas restTableStylePPrs counters' (Some label :: acc)
                | None ->
                    // Should not happen if isHeadingParagraph = true, but fallback
                    loop restParas restTableStylePPrs counters (None :: acc)
            else
                loop restParas restTableStylePPrs counters (None :: acc)
        | _ -> failwith "The lists of paragraphs and tableStylePPrs must have the same length."

    loop paras tableStylePPrs [] []


/// <summary>
/// Generates hierarchical list labels (e.g., "1.", "1.1.", "1.1.2.")
/// for a given list of paragraphs based on their list information.
/// Non-list paragraphs are assigned None.
/// </summary>
/// <param name="paras">The list of paragraph (w:p) elements.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for paragraph properties.</param>
/// <param name="tableStylePPrs">A list of optional tableStylePPrs corresponding to the paragraphs.</param>
/// <returns>A list of option strings: Some(label) for list items, None for non-list items.</returns>
let generateResolvedListLabel
    (paras: XElement list)
    (stylesRoot: XElement)
    (docDefaults: PPr)
    (tableStylePPrs: (PPr option) list)
    : string option list =

    let rec updateCounters (counters: int list) (level: int) : int list =
        let counters =
            if counters.Length <= level then
                counters @ List.replicate (level - counters.Length + 1) 0
            else counters
        counters
        |> List.mapi (fun idx v ->
            if idx < level then v
            elif idx = level then v + 1
            else 0)

    let buildLabel (counters: int list) (level: int) : string =
        counters
        |> List.take (level + 1)
        |> List.map string
        |> String.concat "."
        |> fun s -> s + "."

    let rec tryGetListInfo (para: XElement) (stylesRoot: XElement) (docDefaults: PPr) (tableStylePPr: PPr option) : (string * int) option =
        evaluateEffectiveParagraphProperty "numPr" para stylesRoot docDefaults tableStylePPr
        |> Option.bind (fun numPr ->
            let numIdOpt =
                tryElement (w + "numId") numPr
                |> Option.bind (tryAttrValue (w + "val"))

            let ilvlOpt =
                tryElement (w + "ilvl") numPr
                |> Option.bind (tryAttrValue (w + "val"))
                |> Option.bind (fun s ->
                    match System.Int32.TryParse(s) with
                    | true, v -> Some v
                    | false, _ -> None)

            match numIdOpt, ilvlOpt with
            | Some numId, Some ilvl -> Some (numId, ilvl)
            | _ -> None)

    let rec loop (paras: XElement list) (tableStylePPrs: (PPr option) list) (counters: int list) (currentNumId: string option) (acc: string option list) =
        match paras, tableStylePPrs with
        | [], [] -> List.rev acc
        | para::restParas, tPPrOpt::restTableStylePPrs ->
            if isBulletParagraph para stylesRoot docDefaults tPPrOpt then
                match tryGetListInfo para stylesRoot docDefaults tPPrOpt with
                | Some (numId, ilvl) ->
                    let counters, currentNumId =
                        match currentNumId with
                        | Some id when id = numId -> counters, currentNumId
                        | _ -> [], Some numId // リストIDが変わったらリセット

                    let counters' = updateCounters counters ilvl
                    let label = buildLabel counters' ilvl
                    loop restParas restTableStylePPrs counters' currentNumId (Some label :: acc)
                | None ->
                    loop restParas restTableStylePPrs counters currentNumId (None :: acc)
            else
                loop restParas restTableStylePPrs counters currentNumId (None :: acc)
        | _ -> failwith "The lists of paragraphs and tableStylePPrs must have the same length."

    loop paras tableStylePPrs [] None []
