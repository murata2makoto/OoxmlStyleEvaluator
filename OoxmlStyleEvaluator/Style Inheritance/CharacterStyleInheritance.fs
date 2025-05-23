module OoxmlStyleEvaluator.CharacterStyleInheritance

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes


/// <summary>
/// Creates a character style resolver with memoization.
/// Resolves run properties (`rPr`) by following the `basedOn` 
/// inheritance chain.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// A function that takes a style ID and returns a map representing the 
/// effective run properties (`rPr`) by merging from all applicable styles 
/// in the inheritance chain.
/// The most specific style (closest to the target) takes precedence.
/// </returns>
let createCharacterStyleResolver (stylesRoot: XElement) : (string -> RPr) =
    let memo = System.Collections.Generic.Dictionary<string, RPr>()

    // Recursive resolver function. Checks memo first.
    let rec resolve (styleId: string) : RPr =
        match memo.TryGetValue styleId with
        | true, cached -> cached
        | false, _ ->
            // Look up the character style element by styleId
            match findStyle stylesRoot "character" styleId with
            | None ->
                // No such style found; cache and return empty properties
                memo.[styleId] <- Map.empty
                Map.empty
            | Some style ->
                // Resolve the parent style (via basedOn)
                let basedOnId = resolveBasedOn style
                let parentRPr =
                    basedOnId
                    |> Option.map resolve
                    |> Option.defaultValue Map.empty

                // Extract the current style's run properties
                let thisRPr =
                    style
                    |> tryElement (w + "rPr")
                    |> extractProperties

                // Merge parent properties, giving priority to thisRPr
                let merged =
                    Map.fold (fun (acc: RPr) k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisRPr parentRPr

                // Cache and return the result
                memo.[styleId] <- merged
                merged

    resolve  // Return the resolver function as a closure

