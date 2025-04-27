module OoxmlStyleEvaluator.StyleUtilities

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// Extracts properties from an optional XElement and converts them into a map.
let extractProperties (elementOpt: XElement option) : Map<string, XElement> =
    match elementOpt with
    | Some element -> 
        element.Elements() 
        |> Seq.map (fun e -> e.Name.LocalName.ToString(), e) 
        |> Map.ofSeq
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
    let local =
        if name.Contains "}" then name.Substring(name.IndexOf('}') + 1)
        else name
    [ "b"; "bcs"; "caps"; "emboss"; "i"; "iCs"; "imprint"; 
      "outline"; "shadow"; "smallCaps"; "strike"; "vanish" ]
    |> List.contains local

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
let xor3 (a: bool option) (b: bool option) (c: bool option) : bool option =
    let values = [a; b; c] |> List.choose id
    match values.Length with
    | 0 -> None
    | 1 | 3 -> Some true
    | 2 -> Some false
    | _ -> failwith "Shouldn't happen"

/// <summary>
/// Retrieves the effective boolean value of a toggle property (e.g., bold, italic) from a property map.
/// </summary>
/// <param name="propMap">The property map (RPr or PPr).</param>
/// <param name="name">The property name to retrieve.</param>
/// <returns>
/// Some true, Some false depending on the property value;
/// or None if the property does not exist.
/// </returns>
let getToggleValue (propMap: Map<string, XElement>) (name: string) : bool option =
    match propMap.TryFind name with
    | Some e ->
        match tryAttrValue (w + "val") e with
        | None -> Some true  // <w:b/> etc. without val means true
        | Some "true" | Some "1" -> Some true
        | Some "false" | Some "0" -> Some false
        | _ -> None
    | None -> None

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