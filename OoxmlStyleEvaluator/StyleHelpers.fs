module OoxmlStyleEvaluator.StyleHelpers

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// Extracts properties from an optional XElement and converts them into a map.
let extractProperties (elementOpt: XElement option) : Map<string, XElement> =
    match elementOpt with
    | Some element -> 
        element.Elements() 
        |> Seq.map (fun e -> e.Name.ToString(), e) 
        |> Map.ofSeq
    | None -> Map.empty

/// Resolves the `basedOn` attribute for a given style.
let resolveBasedOn (style: XElement) : string option =
    style
    |> tryElement (w + "basedOn")
    |> Option.bind (tryAttrValue (w + "val"))
