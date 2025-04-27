module OoxmlStyleEvaluator.NumberingStyleInheritance

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes

/// <summary>
/// Resolves the numbering style for a given style ID, including all properties in `pPr`.
/// Note: According to CD 29500-1:2025, numbering styles do not support inheritance via the `basedOn` attribute.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="styleId">The style ID to resolve (value of `w:styleId`).</param>
/// <returns>
/// A map representing the effective paragraph properties (`pPr`) for the specified numbering style.
/// </returns>
let resolveNumberingStyle (stylesRoot: XElement) (styleId: string): PPr =

    match findStyle stylesRoot "numbering" styleId with
    | None -> Map.empty
    | Some style ->
        let thisPPr =
            style
            |> tryElement (w + "pPr")
            |> extractProperties

        // Since numbering styles do not support inheritance, return only the properties in thisPPr
        thisPPr
