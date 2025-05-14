module OoxmlStyleEvaluator.DocumentDefaults

open System.Xml.Linq
open XmlHelpers
open PropertyTypes

/// <summary>
/// Retrieves the default run properties (`rPr`) and paragraph properties (`pPr`) defined in the styles.xml's docDefaults section.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml</param>
/// <returns>
/// A tuple containing two maps:
/// - The first map represents the default run properties (`rPr`), where keys are fully qualified strings and values are their corresponding XElement definitions.
/// - The second map represents the default paragraph properties (`pPr`), where keys are fully qualified strings and values are their corresponding XElement definitions.
/// If no docDefaults, rPrDefault, or pPrDefault is found, the corresponding map will be empty.
/// </returns>
let getDocDefaults (stylesRoot: XElement) : RPr * PPr =
    // Extract rPrDefault properties
    let rPrDefaults =
        stylesRoot
        |> tryElement (w + "docDefaults")
        |> Option.bind (tryElement (w + "rPrDefault"))
        |> Option.bind (tryElement (w + "rPr"))
        |> Option.map (fun rPr ->
            rPr.Elements()
            |> Seq.map (fun e -> e.Name.ToString(), e)
            |> Map.ofSeq)
        |> Option.defaultValue Map.empty

    // Extract pPrDefault properties
    let pPrDefaults =
        stylesRoot
        |> tryElement (w + "docDefaults")
        |> Option.bind (tryElement (w + "pPrDefault"))
        |> Option.bind (tryElement (w + "pPr"))
        |> Option.map (fun pPr ->
            pPr.Elements()
            |> Seq.map (fun e -> e.Name.ToString(), e)
            |> Map.ofSeq)
        |> Option.defaultValue Map.empty

    (rPrDefaults, pPrDefaults)
