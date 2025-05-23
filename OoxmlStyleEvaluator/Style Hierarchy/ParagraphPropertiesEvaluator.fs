module OoxmlStyleEvaluator.ParagraphPropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open PropertyTypes
open StyleUtilities


    /// <summary>
    /// Resolves the effective value of a single paragraph property (from <c>&lt;pPr&gt;</c>)
    /// for a given paragraph, using document defaults, styles, and direct formatting.
    /// </summary>
    /// <param name="paragraphElement">The paragraph (<c>&lt;w:p&gt;</c>) element to process.</param>
    /// <param name="stylesRoot">The root <c>styles.xml</c> element.</param>
    /// <param name="docPPrDefaults">The paragraph property defaults defined in the document.</param>
    /// <param name="tableStyleResolver">A function to resolve a styleId into table style properties.</param>
    /// <param name="paragraphStyleResolver">A function to resolve a paragraph styleId into its RPr and PPr properties.</param>
    /// <param name="key">The property key to look up (e.g., <c>"spacing/@before"</c>).</param>
    /// <returns>The effective value as <c>Some string</c>, or <c>None</c> if not found.</returns>
    let resolveEffectiveParagraphProperty
        (paragraphElement: XElement)
        (docPPrDefaults: PPr)
        (tableStyleResolver: string -> TableStyleProperties)
        (paragraphStyleResolver: string -> RPr * PPr)
        (getPropertyByKey: XElement -> string -> string option)
        (key: string)
        : string option =


        // Step 1: Try direct value from <pPr>
        let directValue =
            paragraphElement
            |> tryElement (w + "pPr")
            |> Option.bind (fun pPr -> getPropertyByKey pPr key)

        match directValue with
        | Some v -> Some v
        | None ->
            // Step 2: Try value from paragraph style
            let fromStyle =
                let styleIdOpt = tryGetStyleId paragraphElement "pPr" "pStyle"
                styleIdOpt
                |> Option.bind (fun styleId ->
                    let _, pPr = paragraphStyleResolver styleId
                    Map.tryFind key pPr
                )

            match fromStyle with
            | Some v -> Some v
            | None ->
                // Step 3: Try value from table style (if inside <tbl>)
                let fromTableStyle =
                    let tableElement = paragraphElement.Ancestors(w + "tbl") |> Seq.tryHead
                    tableElement
                    |> Option.bind (tryElement (w + "tblPr"))
                    |> Option.bind (tryElement (w + "tblStyle"))
                    |> Option.bind (tryAttrValue (w + "val"))
                    |> Option.bind (fun styleId ->
                        let styleChain = tableStyleResolver styleId
                        let fromTop = Map.tryFind key styleChain.TopLevel.PPr
                        let fromWhole =
                            styleChain.ByType
                            |> Map.tryFind "wholeTable"
                            |> Option.bind (fun group -> Map.tryFind key group.PPr)
                        fromTop |> Option.orElse fromWhole
                    )

                match fromTableStyle with
                | Some v -> Some v
                | None ->
                    // Step 4: Try document default
                    Map.tryFind key docPPrDefaults
