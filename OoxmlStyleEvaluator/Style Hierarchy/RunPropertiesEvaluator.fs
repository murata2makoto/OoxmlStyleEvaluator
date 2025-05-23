module OoxmlStyleEvaluator.RunPropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open PropertyTypes
open StyleUtilities

/// <summary>
/// Resolves the effective value of a single run property (from <c>&lt;rPr&gt;</c>)
/// for a given run element, using document defaults, styles, and direct formatting.
/// </summary>
/// <param name="runElement">The <c>&lt;w:r&gt;</c> element to process.</param>
/// <param name="docDefaults">The document default run properties.</param>
/// <param name="tableStyleResolver">A function that resolves table styleId into properties.</param>
/// <param name="paragraphStyleResolver">A function that resolves paragraph styleId into (RPr, PPr).</param>
/// <param name="characterStyleResolver">A function that resolves character styleId into RPr.</param>
/// <param name="getPropertyByKey">A memoized function to get a property from an element by key.</param>
/// <param name="key">The property key to retrieve (e.g., <c>"b/@val"</c>).</param>
/// <returns>The effective property value as <c>Some string</c>, or <c>None</c>.</returns>
let resolveEffectiveRunProperty
    (runElement: XElement)
    (docDefaults: RPr)
    (tableStyleResolver: string -> TableStyleProperties)
    (paragraphStyleResolver: string -> RPr * PPr)
    (characterStyleResolver: string -> RPr)
    (getPropertyByKey: XElement -> string -> string option)
    (key: string)
    : string option =

    // Step 1: Check direct formatting
    let directValue =
        runElement
        |> tryElement (w + "rPr")
        |> Option.bind (fun rPr -> getPropertyByKey rPr key)

    match directValue with
    | Some v -> Some v

    | None when isToggleProperty key ->
        // Special logic for toggle properties

        // Step 2: Check document default = true
        match getToggleValue docDefaults key with
        | Some true -> Some "true"
        | _ ->
            // Step 3: xor3(tableStyle, paragraphStyle, characterStyle)
            let fromTableStyle =
                runElement.Ancestors(w + "tbl") |> Seq.tryHead
                |> Option.bind (tryElement (w + "tblPr"))
                |> Option.bind (tryElement (w + "tblStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
                |> Option.bind (fun styleId ->
                    let styleChain = tableStyleResolver styleId
                    let top = getToggleValue styleChain.TopLevel.RPr key
                    let whole =
                        styleChain.ByType
                        |> Map.tryFind "wholeTable"
                        |> Option.bind (fun group -> getToggleValue group.RPr key)
                    top |> Option.orElse whole)

            let fromParagraphStyle =
                runElement.Ancestors(w + "p") |> Seq.tryHead
                |> Option.bind (fun p ->
                    tryGetStyleId p "pPr" "pStyle"
                    |> Option.map paragraphStyleResolver
                    |> Option.bind (fun (rPr, _) -> getToggleValue rPr key))

            let fromCharacterStyle =
                tryGetStyleId runElement "rPr" "rStyle"
                |> Option.map characterStyleResolver
                |> Option.bind (fun rPr -> getToggleValue rPr key)

            Some (xor3 fromTableStyle fromParagraphStyle fromCharacterStyle |> string)

    | None ->
        // Non-toggle property fallback (priority order)

        let fromCharacterStyle =
            tryGetStyleId runElement "rPr" "rStyle"
            |> Option.map characterStyleResolver
            |> Option.bind (fun rPr -> Map.tryFind key rPr)
        let fromParagraphStyle =
            runElement.Ancestors(w + "p") |> Seq.tryHead
            |> Option.bind (fun p ->
                tryGetStyleId p "pPr" "pStyle"
                |> Option.map paragraphStyleResolver
                |> Option.bind (fun (rPr, _) -> Map.tryFind key rPr))

        let fromTableStyle =
            runElement.Ancestors(w + "tbl") |> Seq.tryHead
            |> Option.bind (tryElement (w + "tblPr"))
            |> Option.bind (tryElement (w + "tblStyle"))
            |> Option.bind (tryAttrValue (w + "val"))
            |> Option.bind (fun styleId ->
                let styleChain = tableStyleResolver styleId
                let top = Map.tryFind key styleChain.TopLevel.RPr
                let whole =
                    styleChain.ByType
                    |> Map.tryFind "wholeTable"
                    |> Option.bind (fun group -> Map.tryFind key group.RPr)
                top |> Option.orElse whole)

        docDefaults
        |> Map.tryFind key
        |> Option.orElse fromCharacterStyle
        |> Option.orElse fromParagraphStyle
        |> Option.orElse fromTableStyle
