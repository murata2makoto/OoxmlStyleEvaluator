module OoxmlStyleEvaluator.TableStyleHelpers

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OptionMonad

/// <summary>
/// Finds the closest ancestor table (w:tbl) element for a given paragraph.
/// </summary>
/// <param name="para">The paragraph (w:p) element.</param>
/// <returns>The closest ancestor table element if found, otherwise None.</returns>
let findAncestorTable (para: XElement) : XElement option =
    para.Ancestors(w + "tbl") |> Seq.tryHead

/// <summary>
/// Attempts to retrieve the tableStyleId (w:tblStyle/@w:val) from a table element.
/// </summary>
/// <param name="tbl">The table (w:tbl) element.</param>
/// <returns>The style ID string if found, otherwise None.</returns>
let tryGetTableStyleId (tbl: XElement) : string option =
    tryElement (w + "tblPr") tbl
    |> Option.bind (tryElement (w + "tblStyle"))
    |> Option.bind (tryAttrValue (w + "val"))

/// <summary>
/// Finds the corresponding table style definition from styles.xml given a styleId.
/// </summary>
/// <param name="styleId">The style ID to search for.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>The style definition XElement if found, otherwise None.</returns>
let findTableStyle (styleId: string) (stylesRoot: XElement) : XElement option =
    stylesRoot.Elements(w + "style")
    |> Seq.tryFind (fun s ->
        option {
            let! sid = s.Attribute(w + "styleId") |> Option.ofObj
            let! typ = s.Attribute(w + "type") |> Option.ofObj
            return sid.Value = styleId && typ.Value = "table"
        } |> function Some(x) -> x | None -> false
        )
/// <summary>
/// Extracts run properties (rPr) from a table style definition.
/// </summary>
/// <param name="tableStyle">The table style XElement.</param>
/// <returns>A map of run properties (property name to XElement).</returns>
let extractTableStyleRPr (tableStyle: XElement) : Map<string, XElement> =
    tableStyle
    |> tryElement (w + "rPr")
    |> Option.map (fun rpr ->
        rpr.Elements()
        |> Seq.map (fun e -> e.Name.ToString(), e)
        |> Map.ofSeq)
    |> Option.defaultValue Map.empty
