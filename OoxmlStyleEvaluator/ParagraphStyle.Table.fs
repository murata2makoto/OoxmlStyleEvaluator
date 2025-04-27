module OoxmlStyleEvaluator.ParagraphStyle.Table

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.TableStyleHelpers
open Core

/// <summary>
/// Finds the effective paragraph style for a paragraph element,
/// considering the table style if no direct paragraph style is specified.
/// </summary>
/// <param name="para">The paragraph (w:p) element.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// The paragraph style XElement if found, either directly from the paragraph
/// or indirectly from the containing table's style; otherwise None.
/// </returns>
let findEffectiveParagraphStyle (para: XElement) (stylesRoot: XElement) : XElement option =
    match tryGetParagraphStyleId para with
    | Some styleId ->
        // The paragraph has an explicit style; search for it
        stylesRoot.Elements(w + "style")
        |> Seq.tryFind (fun s ->
            match Option.ofObj (s.Attribute(w + "styleId")), Option.ofObj (s.Attribute(w + "type")) with
            | Some sid, Some typ -> sid.Value = styleId && typ.Value = "paragraph"
            | _ -> false)
    | None ->
        // No paragraph style; check if the paragraph is inside a table with a table style
        match findAncestorTable para with
        | Some tbl ->
            match tryGetTableStyleId tbl with
            | Some tableStyleId ->
                findTableStyle tableStyleId stylesRoot
            | None -> None
        | None -> None

