module OoxmlStyleEvaluator.RunPropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open PropertyTypes
open StyleUtilities
open DocumentDefaults
open TableStyleInheritance
open ParagraphStyleInheritance
open CharacterStyleInheritance
open RunDirectProperties

/// <summary>
/// Resolves the effective run properties (`rPr`) for a given run (`<w:r>`).
/// </summary>
/// <param name="runElement">The run (`<w:r>`) element to process.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>A map of effective run properties (`rPr`).</returns>
let resolveEffectiveRunProperties (runElement: XElement) (stylesRoot: XElement): PPr =

    // Step 1: Get document defaults for run properties
    let docDefaults =
        let rPrDefaults, _ = getDocDefaults stylesRoot
        rPrDefaults

    // Step 2: Resolve run properties from the nearest ancestor table's style
    let tableStyleProperties =
        let tableElement = runElement.Ancestors(w + "tbl") |> Seq.tryHead
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
                mergeProperties styleChain.TopLevel.RPr
                    (styleChain.ByType
                     |> Map.tryFind "wholeTable"
                     |> Option.map (fun group -> group.RPr)
                     |> Option.defaultValue Map.empty)
            | None -> Map.empty
        | None -> Map.empty

    // Step 3: Resolve run properties from the associated paragraph style
    let paragraphStyleProperties =
        let paragraphElement = runElement.Ancestors(w + "p") |> Seq.tryHead
        match paragraphElement with
        | Some paragraph ->
            let styleIdOpt = tryGetStyleId paragraph "pPr" "pStyle"
            match styleIdOpt with
            | Some styleId ->
                let rPr, _ = resolveParagraphStyleChain stylesRoot styleId
                rPr
            | None -> Map.empty
        | None -> Map.empty

    // Step 4: Resolve run properties from the associated character style
    let characterStyleProperties =
        let styleIdOpt = tryGetStyleId runElement "rPr" "rStyle"
        match styleIdOpt with
        | Some styleId -> resolveCharacterStyleChain stylesRoot styleId
        | None -> Map.empty

    // Step 5: Resolve directly specified run properties
    let directProperties = resolveDirectRunProperties runElement

    // Step 6: Combine properties with special handling for toggle properties
    let allProperties =
        docDefaults
        |> mergeProperties tableStyleProperties
        |> mergeProperties paragraphStyleProperties
        |> mergeProperties characterStyleProperties
        |> mergeProperties directProperties

    // Handle toggle properties
    allProperties
    |> Map.map (fun key value ->
        if isToggleProperty key then
            // Step 6.1: Check directly specified value
            match getToggleValue directProperties key with
            | Some toggleValue -> new XElement(XName.Get key, new XAttribute(w + "val", toggleValue.ToString().ToLower()))
            | None ->
                // Step 6.2: Check document defaults
                match getToggleValue docDefaults key with
                | Some true -> new XElement(XName.Get key, new XAttribute(w + "val", "true"))
                | _ ->
                    // Step 6.3: Use xor3 for table, paragraph, and character styles
                    let tableToggle = getToggleValue tableStyleProperties key
                    let paragraphToggle = getToggleValue paragraphStyleProperties key
                    let characterToggle = getToggleValue characterStyleProperties key
                    match xor3 tableToggle paragraphToggle characterToggle with
                    | Some result -> new XElement(XName.Get key, new XAttribute(w + "val", result.ToString().ToLower()))
                    | None -> value // Fallback to the original value (which is false)
        else
            value)

