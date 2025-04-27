module OoxmlStyleEvaluator.ParagraphPropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open PropertyTypes
open StyleUtilities
open DocumentDefaults
open TableStyleInheritance
open ParagraphStyleInheritance
open ParagraphDirectProperties

/// <summary>
/// Resolves the effective paragraph properties (`pPr`) for a given paragraph (`<w:p>`).
/// </summary>
/// <param name="paragraphElement">The paragraph (`<w:p>`) element to process.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>A map of effective paragraph properties (`pPr`).</returns>
let resolveEffectiveParagraphProperties (paragraphElement: XElement) (stylesRoot: XElement): PPr =

    // Step 1: Get document defaults for paragraph properties
    let docDefaults =
        let _, pPrDefaults = getDocDefaults stylesRoot
        pPrDefaults

    // Step 2: Resolve paragraph properties from the nearest ancestor table's style
    let tableStyleProperties =
        let tableElement = paragraphElement.Ancestors(w + "tbl") |> Seq.tryHead
        match tableElement with
        | Some tbl ->
            let styleIdOpt =
                tbl
                |> tryElement (w + "tblPr")
                |> Option.bind (tryElement (w + "tblStyle"))
                |> Option.bind (tryAttrValue (w + "val"))

            match styleIdOpt with
            | Some styleId ->
                let styleChain = resolveTableStyleChain stylesRoot styleId
                mergeProperties styleChain.TopLevel.PPr
                    (styleChain.ByType
                     |> Map.tryFind "wholeTable"
                     |> Option.map (fun group -> group.PPr)
                     |> Option.defaultValue Map.empty)
            | None -> Map.empty
        | None -> Map.empty

    // Step 3: Resolve paragraph properties from the associated paragraph style
    let paragraphStyleProperties =
        let styleIdOpt = tryGetStyleId paragraphElement "pPr" "pStyle"
        match styleIdOpt with
        | Some styleId ->
            let _, pPr = resolveParagraphStyleChain stylesRoot styleId
            pPr
        | None -> Map.empty

    // Step 4: Resolve directly specified paragraph properties
    let directProperties = resolveDirectParagraphProperties paragraphElement

    // Combine properties in order of priority
    docDefaults
    |> mergeProperties tableStyleProperties
    |> mergeProperties paragraphStyleProperties
    |> mergeProperties directProperties

