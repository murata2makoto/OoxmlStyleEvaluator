module OoxmlStyleEvaluator.CharacterStyleInheritance

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes

/// <summary>
/// Resolves the character style inheritance chain for a given style ID.
/// Combines the run properties (`rPr`) from the entire chain, following the `basedOn` attribute recursively.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="styleId">The style ID to resolve (value of `w:styleId`).</param>
/// <returns>
/// A map representing the effective run properties (`rPr`) by merging from all applicable styles in the inheritance chain.
/// The most specific style (closest to the target) takes precedence.
/// </returns>
let rec resolveCharacterStyleChain (stylesRoot: XElement) (styleId: string): RPr =

    match findStyle stylesRoot "character" styleId with
    | None -> Map.empty
    | Some style ->
        // Resolve the `basedOn` attribute to find the parent style
        let basedOn = resolveBasedOn style

        // Recursively resolve the parent style's run properties
        let parentRPr =
            basedOn
            |> Option.map (resolveCharacterStyleChain stylesRoot)
            |> Option.defaultValue Map.empty

        // Extract the current style's run properties
        let thisRPr =
            style
            |> tryElement (w + "rPr")
            |> extractProperties

        // Merge parent properties, giving priority to properties in thisRPr
        Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisRPr parentRPr
