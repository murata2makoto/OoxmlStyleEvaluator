module OoxmlStyleEvaluator.StyleUtilities

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open PropertyExtractor

/// Extracts properties from an optional XElement and converts them into a map.
let extractProperties (elementOpt: XElement option) : Map<string, string> =
    match elementOpt with
    | Some element -> 
        extractPropertiesAsMap element
    | None -> Map.empty

/// Resolves the `basedOn` attribute for a given style.
let resolveBasedOn (style: XElement) : string option =
    style
    |> tryElement (w + "basedOn")
    |> Option.bind (tryAttrValue (w + "val"))

/// <summary>
/// Attempts to retrieve the styleId of a character or paragraph style from a node.
/// </summary>
/// <param name="node">The XML node (XElement) to search within.</param>
/// <param name="prTag">The name of the property tag (e.g., "rPr" for run, "pPr" for paragraph).</param>
/// <param name="styleTag">The style tag name (e.g., "rStyle" or "pStyle").</param>
/// <returns>The styleId string if found; otherwise, None.</returns>
let tryGetStyleId (node: XElement) (prTag: string) (styleTag: string) : string option =
    node
    |> tryElement (w + prTag)
    |> Option.bind (tryElement (w + styleTag))
    |> Option.bind (tryAttrValue (w + "val"))

/// <summary>
/// Determines whether a property name corresponds to a toggle property
/// as defined by ISO/IEC 29500-1.
/// </summary>
/// <param name="name">The property name to check.</param>
/// <returns>True if the property is a toggle property; otherwise, false.</returns>
let isToggleProperty (name: string) : bool =

    let toggleProperties = 
        [ "b"; "bcs"; "caps"; "emboss"; "i"; "iCs"; "imprint"; 
          "outline"; "shadow"; "smallCaps"; "strike"; "vanish" ]
    
    toggleProperties 
    |> List.exists (fun prop ->
            name.EndsWith($"w:{prop}/@w:val" ))

/// <summary>
/// Computes a 3-way XOR for toggle property combination across multiple style levels.
/// </summary>
/// <param name="a">The first toggle value.</param>
/// <param name="b">The second toggle value.</param>
/// <param name="c">The third toggle value.</param>
/// <returns>
/// The effective toggle result: Some true if an odd number of true values exist;
/// Some false if an even number; otherwise, None if no value exists.
/// </returns>
let xor3 (a: bool option) (b: bool option) (c: bool option) : bool  =
    match a, b, c with
    | Some(x), Some(y), Some(z) -> x <> y <> z
    | Some(x), Some(y), None | Some(x), None, Some(y) 
    | None, Some(x), Some(y) -> x<>y
    | Some(x), None, None | None, Some(x), None 
    | None, None, Some(x) -> x
    | None, None, None -> false

/// <summary>
/// Retrieves the effective boolean value of a toggle property (e.g., bold, italic) from a property map.
/// </summary>
/// <param name="propMap">The property map (RPr or PPr).</param>
/// <param name="name">The property name to retrieve.</param>
/// <returns>
/// Some true, Some false depending on the property value;
/// or None if the property does not exist.
/// </returns>
let getToggleValue (propMap: Map<string, string>) (name: string) : bool option  =
    match propMap.TryFind name |> Option.map (fun x -> x.Trim()) with
    | None -> None
    | Some "true" | Some "1" ->  Some(true)
    | Some "false" | Some "0" ->  Some(false)
    | Some(_) -> failwith "Shouldn't happen"

/// <summary>
/// Finds a style element by its ID and ensures it is of the specified type".
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="sid">The style type to find.</param>
/// <param name="sid">The style ID to find.</param>
/// <returns>The corresponding style element if found; otherwise, None.</returns>
let findStyle (stylesRoot: XElement) (styleType: string) (sid: string) =
    stylesRoot.Elements(w + "style")
    |> Seq.tryFind (fun s ->
        tryAttrValue (w + "styleId") s = Some sid &&
        tryAttrValue (w + "type") s = Some styleType)