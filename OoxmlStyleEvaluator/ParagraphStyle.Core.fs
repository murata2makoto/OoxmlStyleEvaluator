module OoxmlStyleEvaluator.ParagraphStyle.Core 

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// <summary>
/// Attempts to retrieve the styleId from a paragraph element's pPr/pStyle.
/// </summary>
let tryGetParagraphStyleId (para: XElement) : string option =
    para
    |> tryElement (w + "pPr")
    |> Option.bind (tryElement (w + "pStyle"))
    |> Option.bind (tryAttrValue (w + "val"))

/// <summary>
/// Finds the style definition element (&lt;w:style&gt;) for the given styleId and type "paragraph".
/// </summary>
let findParagraphStyle (styleId: string) (stylesRoot: XElement) : XElement option =
    stylesRoot.Elements(w + "style")
    |> Seq.tryFind (fun style ->
        tryAttrValue (w + "styleId") style = Some styleId &&
        tryAttrValue (w + "type") style = Some "paragraph")
