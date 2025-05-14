module OoxmlStyleEvaluator.TableRowPropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open TableStyleInheritance
open TableDirectProperties
open PropertyTypes

/// <summary>
/// Resolves the effective table row properties (`trPr`) for a given table row (`<w:tr>`).
/// </summary>
/// <param name="rowElement">The table row (`<w:tr>`) element to process.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>A map of effective table row properties (`trPr`).</returns>
let resolveEffectiveRowProperties (rowElement: XElement) (stylesRoot: XElement): TrPr =

    // Step 1: Find the nearest ancestor table (`<w:tbl>`) element
    let tableElement = rowElement.Ancestors(w + "tbl") |> Seq.tryHead

    // Step 2: Resolve trPr properties from the table style
    let tableStyleProperties =
        match tableElement with
        | Some tbl ->
            // Get the styleId of the table
            let styleIdOpt =
                tbl
                |> tryElement (w + "tblPr")
                |> Option.bind (tryElement (w + "tblStyle"))
                |> Option.bind (tryAttrValue (w + "val"))

            match styleIdOpt with
            | Some styleId ->
                // Resolve the table style chain
                let styleChain = resolveTableStyleChain stylesRoot styleId
                // Combine trPr properties from the top-level and tblStylePr/trPr
                mergeProperties styleChain.TopLevel.TrPr
                    (styleChain.ByType
                     |> Map.tryFind "wholeTable"
                     |> Option.map (fun group -> group.TrPr)
                     |> Option.defaultValue Map.empty)
            | None -> Map.empty
        | None -> Map.empty

    // Step 3: Resolve trPr properties directly specified in the row
    let rowProperties = resolveDirectRowProperties rowElement

    // Step 4: Combine properties in order of priority
    mergeProperties tableStyleProperties rowProperties

