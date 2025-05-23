module OoxmlStyleEvaluator.PropertyExtractor

open System.Xml.Linq
open XmlHelpers

/// <summary>
/// Local names of elements having the type `CT_OnOff`.
/// </summary>
let localNamesOfCTOnOff = [
        "keepNext"; "keepLines"; "pageBreakBefore"; "widowControl"; 
        "suppressLineNumbers"; "suppressAutoHyphens"; "kinsoku"; 
        "wordWrap"; "overflowPunct"; "topLinePunct"; "autoSpaceDE"; 
        "autoSpaceDN"; "bidi"; "adjustRightInd"; "snapToGrid"; 
        "contextualSpacing"; "mirrorIndents"; "suppressOverlap"; 
    
        "b"; "bCs"; "i"; "iCs"; "caps"; "smallCaps"; "strike"; "dstrike"; 
        "outline"; "shadow"; "emboss"; "imprint"; "noProof"; "vanish"; 
        "webHidden"; "rtl"; "cs"; "specVanish"; "oMath"; 
        
        "noWrap"; "tcFitText"; "hideMark"; 
    
        "hidden"; "cantSplit"; "tblHeader"; 
    
        "bidiVisual"]

/// <summary>
/// Determines if an element is of the type `CT_OnOff`.
/// </summary>
let isCTOnOff (element: XElement): bool =
    let localName = element.Name.LocalName
    List.contains localName localNamesOfCTOnOff

/// <summary>
/// Extracts all attributes from the descendants of the given XElement,
/// returning a map where keys are formed as "w:child1/w:child2/.../w:attribute"
/// and values are the corresponding attribute values.
/// The root element itself (rPr, pPr, tblPr, trPr, or tcPr) is not included in the path.
/// </summary>
/// <param name="element">The root XElement to process. Must have local name rPr, pPr, tblPr, trPr, or tcPr.</param>
/// <returns>
/// A map from string keys of the form "w:child1/w:child2/.../w:attribute" to string values.
/// </returns>
/// <exception cref="System.Exception">
/// Thrown if the root element's local name is not one of the permitted names.
/// </exception>
let extractPropertiesAsMap (element: XElement) : Map<string, string> =
    let validNames = set ["rPr"; "pPr"; "tblPr"; "trPr"; "tcPr"]

    let localName = element.Name.LocalName
    if not (validNames.Contains localName) then
        failwithf "Unexpected element: %s. Expected one of: %A" localName validNames

    /// <summary>
    /// Recursively extracts attributes from descendants,
    /// skipping the root element itself from the path.
    /// </summary>
    let rec extract (parent: XElement) (pathPrefix: string) (acc: Map<string, string>) : Map<string, string> =
        let currentElemName = parent.Name.LocalName
        let currentPath =
            if pathPrefix = "" then currentElemName
            else pathPrefix + "/" + currentElemName

        // Treat <w:i/> as <w:i w:val="true"/>, for example.
        if isCTOnOff parent then 
            let v = 
                match tryAttrValue (w + "val") parent 
                      |> Option.map (fun x -> x.Trim()) with 
                | None -> "true"
                | Some ("1") | Some ("true")  -> "true"
                | Some ("0") | Some ("false") -> "false"
                | _ -> failwith "shouldn't happen"
            acc |> Map.add (currentPath + "/@val") v

        else
            // Ensure the element is registered if it has no attributes.
            let accWithElement =
                if parent.HasAttributes then 
                    acc
                else 
                    acc |> Map.add currentPath ""

            // Add attributes at current level
            let updatedAcc =
                parent.Attributes()
                |> Seq.fold (fun m attr ->
                    let attrName = attr.Name.LocalName
                    let key = sprintf "%s/@%s" currentPath attrName
                    m |> Map.add key attr.Value
                ) accWithElement

            // Recurse into children
            parent.Elements()
            |> Seq.fold (fun m child -> extract child currentPath m) updatedAcc

    // Start recursion from children of root, skipping root itself
    element.Elements()
    |> Seq.fold (fun acc child -> extract child "" acc) Map.empty

/// <summary>
/// Creates a function that retrieves a single property value from a WordprocessingML style element
/// (such as <c>&lt;rPr&gt;</c>, <c>&lt;pPr&gt;</c>, <c>&lt;tblPr&gt;</c>, etc.)
/// given a property key of the form <c>"color/@val"</c> or <c>"b/@val"</c>.
/// </summary>
/// <returns>
/// A function of type <c>XElement -&gt; string -&gt; string option</c> that takes:
/// <list type="bullet">
///   <item><description><c>element</c>: the root style element (e.g., &lt;rPr&gt;)</description></item>
///   <item><description><c>key</c>: a string representing a property path, such as <c>"color/@val"</c></description></item>
/// </list>
/// and returns <c>Some value</c> if found, or <c>None</c> if not found.
/// </returns>
/// <remarks>
/// <para>The key path should be written using local names only (without namespace prefixes).</para>
/// <para>Keys ending in <c>"@val"</c> and targeting <c>CT_OnOff</c> elements like <c>&lt;b/&gt;</c> will return "true" if no value is explicitly set.</para>
/// <para>Repeated calls with the same key are optimized via internal memoization of the split path.</para>
/// </remarks>
let makeGetPropertyByKey (): (XElement -> string -> string option) =
    let keyPartsCache = System.Collections.Generic.Dictionary<string, string list>()

    let getKeyParts (key: string) =
        match keyPartsCache.TryGetValue(key) with
        | true, parts -> parts
        | false, _ ->
            let parts = key.Split('/') |> Array.toList
            keyPartsCache.Add(key, parts)
            parts

    fun (element: XElement) (key: string) ->
        let parts = getKeyParts key

        let rec follow (elem: XElement) (path: string list) : string option =
            match path with
            | [] -> None

            | [ "@val" ] when isCTOnOff elem ->
                match tryAttrValue (w + "val") elem with
                | Some v ->
                    let v = v.Trim()
                    match v with
                    | "1" | "true" -> Some "true"
                    | "0" | "false" -> Some "false"
                    | _ -> failwith "Invalid CT_OnOff value"
                | None -> Some "true"

            | [ attr ] when attr.StartsWith("@") ->
                let attrName = attr.Substring(1) 
                elem.Attribute(w + attrName) |> Option.ofObj |> Option.map (fun a -> a.Value)

            | name :: rest ->
                let local = name
                let child = elem.Elements() |> Seq.tryFind (fun e -> e.Name.LocalName = local)
                match child with
                | Some e -> follow e rest
                | None -> None

        match parts with
        | [] -> None
        | name :: rest ->
            let local = name
            let start = element.Elements() |> Seq.tryFind (fun e -> e.Name.LocalName = local)
            match start with
            | Some e -> follow e rest
            | None -> None


