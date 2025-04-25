module OoxmlStyleEvaluator.RunStyleEvaluator

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.StyleEvaluatorUtilities
open OoxmlStyleEvaluator.StyleInheritance

/// <summary>
/// Represents a run property map, mapping property names (as strings) to their corresponding XElement values.
/// </summary>
type RPr = Map<string, XElement>

/// <summary>
/// Evaluates the effective value of a run property (e.g., bold, color, font) based on direct formatting,
/// character style, paragraph style, table style, and document defaults.
/// Fully respects the style inheritance (basedOn) hierarchy as specified in ISO/IEC 29500-1.
/// </summary>
/// <param name="propName">The name of the run property to evaluate.</param>
/// <param name="run">The run (w:r) element to evaluate.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional run properties derived from the table style, if applicable.</param>
/// <returns>The effective run property as an XElement if found, otherwise None.</returns>
let evaluateEffectiveProperty
    (propName: string)
    (run: XElement)
    (stylesRoot: XElement)
    (docDefaults: RPr)
    (tableStyleRPr: RPr option) : XElement option =

    let runRPr =
        run
        |> tryElement (w + "rPr")
        |> Option.map (fun rprNode ->
            rprNode.Elements()
            |> Seq.map (fun e -> e.Name.ToString(), e)
            |> Map.ofSeq)
        |> Option.defaultValue Map.empty

    if runRPr.ContainsKey propName then
        runRPr.TryFind propName
    else
        let rStyleId = tryGetStyleId run "rPr" "rStyle"
        let para = run.Ancestors(w + "p") |> Seq.tryHead
        let pStyleId = para |> Option.bind (fun p -> tryGetStyleId p "pPr" "pStyle")

        let rStyleRPr =
            rStyleId
            |> Option.map (resolveStyleChain StyleType.Character stylesRoot)
            |> Option.defaultValue Map.empty

        let pStyleRPr =
            pStyleId
            |> Option.map (resolveStyleChain StyleType.Paragraph stylesRoot)
            |> Option.defaultValue Map.empty

        let tStyleRPr = tableStyleRPr |> Option.defaultValue Map.empty

        if isToggleProperty propName then
            let d = getToggleValue docDefaults propName
            let t = getToggleValue tStyleRPr propName
            let p = getToggleValue pStyleRPr propName
            let r = getToggleValue rStyleRPr propName
            match d with
            | Some true -> Some (new XElement(XName.Get propName))
            | _ ->
                xor3 t p r
                |> Option.map (fun b ->
                    new XElement(XName.Get propName, new XAttribute(w + "val", b.ToString().ToLower())))
        else
            [ docDefaults; tStyleRPr; pStyleRPr; rStyleRPr ]
            |> List.choose (fun rpr -> rpr.TryFind propName)
            |> List.tryLast
