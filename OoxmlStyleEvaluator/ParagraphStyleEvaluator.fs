module OoxmlStyleEvaluator.ParagraphStyleEvaluator

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.StyleEvaluatorUtilities
open OoxmlStyleEvaluator.StyleInheritance

/// <summary>
/// Represents a paragraph property map, mapping property names (as strings) to their corresponding XElement values.
/// </summary>
type PPr = Map<string, XElement>

/// <summary>
/// Evaluates the effective value of a paragraph property (e.g., outlineLvl, numPr)
/// based on direct formatting, paragraph style, table style, and document defaults.
/// Fully respects the style inheritance (basedOn) hierarchy as specified in ISO/IEC 29500-1.
/// </summary>
/// <param name="propName">The name of the paragraph property to evaluate.</param>
/// <param name="para">The paragraph (w:p) element to evaluate.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for paragraph properties.</param>
/// <param name="tableStylePPr">Optional paragraph properties derived from the table style, if applicable.</param>
/// <returns>The effective paragraph property as an XElement if found, otherwise None.</returns>
let evaluateEffectiveParagraphProperty
    (propName: string)
    (para: XElement)
    (stylesRoot: XElement)
    (docDefaults: PPr)
    (tableStylePPr: PPr option) : XElement option =

    let paraPPr =
        para
        |> tryElement (w + "pPr")
        |> Option.map (fun pprNode ->
            pprNode.Elements()
            |> Seq.map (fun e -> e.Name.ToString(), e)
            |> Map.ofSeq)
        |> Option.defaultValue Map.empty

    if paraPPr.ContainsKey propName then
        paraPPr.TryFind propName
    else
        let pStyleId = tryGetStyleId para "pPr" "pStyle"

        let pStylePPr =
            pStyleId
            |> Option.map (resolveStyleChain StyleType.Paragraph stylesRoot)
            |> Option.defaultValue Map.empty

        let tStylePPr = tableStylePPr |> Option.defaultValue Map.empty

        if isToggleProperty propName then
            let d = getToggleValue docDefaults propName
            let t = getToggleValue tStylePPr propName
            let p = getToggleValue pStylePPr propName

            match d with
            | Some true -> Some (new XElement(XName.Get propName))
            | _ ->
                xor3 t p None
                |> Option.map (fun b ->
                    new XElement(XName.Get propName, new XAttribute(w + "val", b.ToString().ToLower())))
        else
            [ docDefaults; tStylePPr; pStylePPr ]
            |> List.choose (fun ppr -> ppr.TryFind propName)
            |> List.tryLast

/// <summary>
/// Determines whether a given paragraph is a bulleted or numbered list item,
/// based on the presence of an effective numbering definition (numPr).
/// Fully respects style inheritance and document defaults.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for paragraph properties.</param>
/// <param name="tableStylePPr">Optional paragraph properties derived from the table style, if applicable.</param>
/// <returns>True if the paragraph is part of a list (bullet or numbered); otherwise, false.</returns>
let isBulletParagraph
    (para: XElement)
    (stylesRoot: XElement)
    (docDefaults: PPr)
    (tableStylePPr: PPr option) : bool =

    evaluateEffectiveParagraphProperty "numPr" para stylesRoot docDefaults tableStylePPr
    |> Option.isSome