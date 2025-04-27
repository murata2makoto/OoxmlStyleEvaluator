module OoxmlStyleEvaluator.ParagraphStyleInheritance

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes

/// <summary>
/// Resolves the paragraph style inheritance chain for a given style ID.
/// Combines the run properties (`rPr`) and paragraph properties (`pPr`) from the entire chain,
/// following the `basedOn` attribute recursively.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="styleId">The style ID to resolve (value of `w:styleId`).</param>
/// <returns>
/// A tuple containing two maps:
/// - The effective run properties (`rPr`) by merging from all applicable styles in the inheritance chain.
/// - The effective paragraph properties (`pPr`) by merging from all applicable styles in the inheritance chain.
/// The most specific style (closest to the target) takes precedence.
/// </returns>
let rec resolveParagraphStyleChain (stylesRoot: XElement) (styleId: string): RPr * PPr =

    match findStyle stylesRoot "paragraph" styleId with
    | None -> (Map.empty, Map.empty)
    | Some style ->
        // Resolve the `basedOn` attribute to find the parent style
        let basedOn = resolveBasedOn style

        // Recursively resolve the parent style's properties
        let (parentRPr, parentPPr) =
            basedOn
            |> Option.map (resolveParagraphStyleChain stylesRoot)
            |> Option.defaultValue (Map.empty, Map.empty)

        // Extract the current style's run properties
        let thisRPr =
            style
            |> tryElement (w + "rPr")
            |> extractProperties

        // Extract the current style's paragraph properties
        let thisPPr =
            style
            |> tryElement (w + "pPr")
            |> extractProperties

        // Merge parent properties, giving priority to properties in thisRPr and thisPPr
        let mergedRPr: RPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisRPr parentRPr

        let mergedPPr: PPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisPPr parentPPr

        (mergedRPr, mergedPPr)
