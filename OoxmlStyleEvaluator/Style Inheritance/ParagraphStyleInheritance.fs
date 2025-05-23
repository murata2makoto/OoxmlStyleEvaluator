module OoxmlStyleEvaluator.ParagraphStyleInheritance

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes


/// <summary>
/// Creates a paragraph style resolver with memoization.
/// Resolves both run properties (`rPr`) and paragraph properties (`pPr`)
/// by following the `basedOn` inheritance chain.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// A function that takes a style ID and returns a tuple:
/// - The effective run properties (`rPr`)
/// - The effective paragraph properties (`pPr`)
/// </returns>
let createParagraphStyleResolver (stylesRoot: XElement) : (string -> RPr * PPr) =
    let memo = System.Collections.Generic.Dictionary<string, RPr * PPr>()

    // Recursive resolver function with memoization
    let rec resolve (styleId: string) : RPr * PPr =
        match memo.TryGetValue styleId with
        | true, cached -> cached
        | false, _ ->
            // Look up the paragraph style element by styleId
            match findStyle stylesRoot "paragraph" styleId with
            | None ->
                // No style found; cache and return empty properties
                memo.[styleId] <- (Map.empty, Map.empty)
                (Map.empty, Map.empty)
            | Some style ->
                // Resolve the parent style (via basedOn)
                let basedOnId = resolveBasedOn style
                let (parentRPr, parentPPr) =
                    basedOnId
                    |> Option.map resolve
                    |> Option.defaultValue (Map.empty, Map.empty)

                // Extract run properties (`rPr`) from the current style
                let thisRPr =
                    style
                    |> tryElement (w + "rPr")
                    |> extractProperties

                // Extract paragraph properties (`pPr`) from the current style
                let thisPPr =
                    style
                    |> tryElement (w + "pPr")
                    |> extractProperties

                // Merge parent and current properties (current takes precedence)
                let mergedRPr =
                    Map.fold (fun (acc: Map<string,string>) k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisRPr parentRPr
                let mergedPPr =
                    Map.fold (fun (acc: Map<string,string>) k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisPPr parentPPr

                // Cache and return the merged result
                let result = (mergedRPr, mergedPPr)
                memo.[styleId] <- result
                result

    resolve  // Return the memoized resolver function as a closure
