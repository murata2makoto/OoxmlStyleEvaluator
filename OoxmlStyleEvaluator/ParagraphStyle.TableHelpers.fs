module OoxmlStyleEvaluator.ParagraphStyle.TableHelpers

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// <summary>
/// Finds the effective paragraph style for a given paragraph,
/// considering both direct pStyle and inherited table style if present.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// Some (styleId, styleElement) if a paragraph style is found; otherwise, None.
/// </returns>
let findEffectiveParagraphStyle (para: XElement) (stylesRoot: XElement) : (string * XElement) option =

    // Try direct pStyle on the paragraph
    let directStyleIdOpt =
        tryElement (w + "pPr") para
        |> Option.bind (tryElement (w + "pStyle"))
        |> Option.bind (tryAttrValue (w + "val"))

    match directStyleIdOpt with
    | Some styleId ->
        stylesRoot.Elements(w + "style")
        |> Seq.tryFind (fun s ->
            (tryAttrValue (w + "styleId") s = Some styleId) &&
            (tryAttrValue (w + "type") s = Some "paragraph"))
        |> Option.map (fun style -> (styleId, style))
    | None ->
        // No direct pStyle, fallback to table's tblStyle
        para.Ancestors(w + "tbl")
        |> Seq.tryHead
        |> Option.bind (fun tbl ->
            tryElement (w + "tblPr") tbl
            |> Option.bind (tryElement (w + "tblStyle"))
            |> Option.bind (tryAttrValue (w + "val")))
        |> Option.bind (fun tblStyleId ->
            stylesRoot.Elements(w + "style")
            |> Seq.tryFind (fun s ->
                (tryAttrValue (w + "styleId") s = Some tblStyleId) &&
                (tryAttrValue (w + "type") s = Some "table"))
            |> Option.map (fun style -> (tblStyleId, style)))
