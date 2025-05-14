module OoxmlStyleEvaluator.StyleInheritance

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// <summary>
/// Represents the type of style as defined in styles.xml (`w:type` attribute).
/// </summary>
type StyleType =
    | Paragraph
    | Character
    | Table
    | Numbering

    /// <summary>
    /// Converts the discriminated union to its corresponding string representation.
    /// </summary>
    member this.AsString =
        match this with
        | Paragraph -> "paragraph"
        | Character -> "character"
        | Table -> "table"
        | Numbering -> "numbering"

/// <summary>
/// A type alias for run properties, represented as a map from fully qualified element names to their XElement definitions.
/// </summary>
type RPr = Map<string, XElement>

/// <summary>
/// Extracts the run properties (`rPr`) from an optional XElement and converts them into a map.
/// </summary>
/// <param name="rPrOpt">An optional XElement representing a &lt;w:rPr&gt; element.</param>
/// <returns>
/// A map of property names to their corresponding XElement definitions. If rPr is None, returns an empty map.
/// </returns>
let extractRPr (rPrOpt: XElement option) : RPr =
    match rPrOpt with
    | Some rPr -> 
        rPr.Elements() 
        |> Seq.map (fun e -> e.Name.ToString(), e) 
        |> Map.ofSeq
    | None -> Map.empty

/// <summary>
/// Resolves the style inheritance chain for a given style, following `basedOn` recursively,
/// and combines the run properties (`rPr`) from the entire chain.
/// </summary>
/// <param name="styleId">The style ID to resolve (value of `w:styleId`).</param>
/// <param name="styleType">The type of style: paragraph, character, etc.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// A map representing the effective run properties by merging from all applicable styles in the inheritance chain.
/// The most specific style (closest to the target) takes precedence.
/// </returns>
let rec resolveStyleChain (styleType: StyleType) (stylesRoot: XElement)  (styleId: string): RPr =
    let findStyle (sid: string) =
        stylesRoot.Elements(w + "style")
        |> Seq.tryFind (fun s ->
            tryAttrValue (w + "styleId") s = Some sid &&
            tryAttrValue (w + "type") s = Some styleType.AsString)

    match findStyle styleId with
    | None -> Map.empty
    | Some style ->
        let basedOn =
            style
            |> tryElement (w + "basedOn")
            |> Option.bind (tryAttrValue (w + "val"))

        let parentRPr =
            basedOn
            |> Option.map (resolveStyleChain styleType stylesRoot)
            |> Option.defaultValue Map.empty

        let thisRPr =
            style
            |> tryElement (w + "rPr")
            |> extractRPr

        // Merge parent properties, giving priority to properties in thisRPr
        Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisRPr parentRPr
