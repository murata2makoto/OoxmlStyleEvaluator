module OoxmlStyleEvaluator.TablePropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open TableStyleInheritance
open TableDirectProperties
open PropertyTypes

/// <summary>
/// Resolves the effective table properties (`tblPr`) for a given table (`<w:tbl>`).
/// </summary>
/// <param name="tableElement">The table (`<w:tbl>`) element to process.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>A map of effective table properties (`tblPr`).</returns>
let resolveEffectiveTableProperties (tableElement: XElement) (stylesRoot: XElement): TblPr =

    // Step 1: Resolve tblPr properties from the table style
    let tableStyleProperties =
        // Get the styleId of the table
        let styleIdOpt =
            tableElement
            |> tryElement (w + "tblPr")
            |> Option.bind (tryElement (w + "tblStyle"))
            |> Option.bind (tryAttrValue (w + "val"))

        match styleIdOpt with
        | Some styleId ->
            // Resolve the table style chain
            let styleChain = resolveTableStyleChain stylesRoot styleId
            // Combine tblPr properties from the top-level and tblStylePr/tblPr
            mergeProperties styleChain.TopLevel.TblPr
                (styleChain.ByType
                 |> Map.tryFind "wholeTable"
                 |> Option.map (fun group -> group.TblPr)
                 |> Option.defaultValue Map.empty)
        | None -> Map.empty

    // Step 2: Resolve tblPr properties directly specified in the table
    let directTableProperties = resolveTableProperties tableElement

    // Step 3: Combine properties in order of priority
    mergeProperties tableStyleProperties directTableProperties

