module OoxmlStyleEvaluator.DocumentDefaults

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// <summary>
/// Retrieves the default run properties (rPr) defined in the styles.xml's docDefaults section.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml</param>
/// <returns>
/// A map of run property names (as fully qualified strings) to their corresponding XElement definitions.
/// If no docDefaults or rPrDefault is found, returns an empty map.
/// </returns>
let getDocDefaultsRPr (stylesRoot: XElement) : Map<string, XElement> =
    stylesRoot
    |> tryElement (w + "docDefaults")
    |> Option.bind (tryElement (w + "rPrDefault"))
    |> Option.bind (tryElement (w + "rPr"))
    |> Option.map (fun rPr ->
        rPr.Elements()
        |> Seq.map (fun e -> e.Name.ToString(), e)
        |> Map.ofSeq)
    |> Option.defaultValue Map.empty
