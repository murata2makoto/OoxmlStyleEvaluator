module OoxmlStyleEvaluator.TableCellPropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open TableStyleInheritance
open TableDirectProperties
open PropertyTypes

/// <summary>
/// Resolves the effective table cell properties (`tcPr`) for a given table cell (`<w:tc>`).
/// </summary>
/// <param name="cellElement">The table cell (`<w:tc>`) element to process.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>A map of effective table cell properties (`tcPr`).</returns>
let resolveEffectiveCellProperties (cellElement: XElement) (stylesRoot: XElement): TcPr =

    // Step 1: Find the nearest ancestor table (`<w:tbl>`) element
    let tableElement = cellElement.Ancestors(w + "tbl") |> Seq.tryHead

    // Step 2: Resolve tcPr properties from the table style
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
                // Combine tcPr properties from the top-level and tblStylePr/tcPr
                mergeProperties styleChain.TopLevel.TcPr
                    (styleChain.ByType
                     |> Map.tryFind "wholeTable"
                     |> Option.map (fun group -> group.TcPr)
                     |> Option.defaultValue Map.empty)
            | None -> Map.empty
        | None -> Map.empty

    // Step 3: Resolve tcPr properties directly specified in the cell
    let cellProperties = resolveDirectCellProperties cellElement

    // Step 4: Combine properties in order of priority
    mergeProperties tableStyleProperties cellProperties

