module OoxmlStyleEvaluator.DocumentDefaults

open System.Xml.Linq
open XmlHelpers
open PropertyTypes
open StyleUtilities

/// <summary>
/// Retrieves the default run properties (`rPr`) defined in the styles.xml's docDefaults section.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml</param>
/// <returns>
/// A map representing the default run properties (`rPr`), where keys are XPath-like strings and values are their corresponding run properties.
/// If no docDefaults, rPrDefault, or pPrDefault is found, the corresponding map will be empty.
/// </returns>
let getRPrDocDefaults (stylesRoot: XElement) : RPr =
    stylesRoot
    |> tryElement (w + "docDefaults")
    |> Option.bind (tryElement (w + "rPrDefault"))
    |> Option.bind (tryElement (w + "rPr"))
    |> extractProperties 

/// <summary>
/// Retrieves the paragraph properties (`pPr`) defined in the styles.xml's docDefaults section.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml</param>
/// <returns>
/// A map representing the default paragraph properties (`pPr`), where keys are XPath-like strings and values are their corresponding paragraph properties.
/// If no docDefaults, rPrDefault, or pPrDefault is found, the corresponding map will be empty.
/// </returns>
let getPPrDocDefaults (stylesRoot: XElement) : PPr =
    stylesRoot
    |> tryElement (w + "docDefaults")
    |> Option.bind (tryElement (w + "pPrDefault"))
    |> Option.bind (tryElement (w + "pPr"))
    |> extractProperties 