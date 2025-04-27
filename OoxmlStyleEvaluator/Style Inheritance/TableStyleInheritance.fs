module OoxmlStyleEvaluator.TableStyleInheritance

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes

/// <summary>
/// Resolves the table style inheritance chain for a given style ID.
/// Combines the top-level properties (`rPr`, `pPr`, `tblPr`, `trPr`, `tcPr`) and `tblStylePr[@type]` groups.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="styleId">The style ID to resolve (value of `w:styleId`).</param>
/// <returns>
/// A `TableStyleProperties` object containing:
/// - Top-level properties (`rPr`, `pPr`, `tblPr`, `trPr`, `tcPr`).
/// - A map of `tblStylePr[@type]` groups, where the key is the `@type` value.
/// </returns>
let rec resolveTableStyleChain (stylesRoot: XElement) (styleId: string): TableStyleProperties =

    /// <summary>
    /// Extracts properties from an element, prioritizing later occurrences of elements with the same QName.
    /// </summary>
    /// <param name="elementOpt">The optional XElement to process.</param>
    /// <returns>A map of properties where later elements override earlier ones.</returns>
    let extractPropertiesWithPriority (elementOpt: XElement option) : Map<string, XElement> =
        match elementOpt with
        | Some element ->
            element.Elements()
            |> Seq.rev // Reverse the order to prioritize later elements
            |> Seq.fold (fun acc elem -> Map.add (elem.Name.ToString()) elem acc) Map.empty
        | None -> Map.empty

    match findStyle stylesRoot "table" styleId with
    | None -> 
        { TopLevel = { RPr = Map.empty; PPr = Map.empty; TblPr = Map.empty; TrPr = Map.empty; TcPr = Map.empty }
          ByType = Map.empty }
    | Some style ->
        // Resolve the `basedOn` attribute to find the parent style
        let basedOn = resolveBasedOn style

        // Recursively resolve the parent style's properties
        let parentProperties =
            basedOn
            |> Option.map (resolveTableStyleChain stylesRoot)
            |> Option.defaultValue 
                { TopLevel = { RPr = Map.empty; PPr = Map.empty; TblPr = Map.empty; TrPr = Map.empty; TcPr = Map.empty }
                  ByType = Map.empty }

        // Extract the current style's top-level properties
        let thisTopLevel = {
            RPr = style |> tryElement (w + "rPr") |> extractProperties
            PPr = style |> tryElement (w + "pPr") |> extractProperties
            TblPr = style |> tryElement (w + "tblPr") |> extractProperties
            TrPr = style |> tryElement (w + "trPr") |> extractPropertiesWithPriority // Use the new function
            TcPr = style |> tryElement (w + "tcPr") |> extractProperties
        }

        // Process tblStylePr[@type] elements
        let thisByType =
            style.Elements(w + "tblStylePr")
            |> Seq.fold (fun acc elem ->
                match tryAttrValue (w + "type") elem with
                | Some typeValue ->
                    let group = {
                        RPr = elem |> tryElement (w + "rPr") |> extractProperties
                        PPr = elem |> tryElement (w + "pPr") |> extractProperties
                        TblPr = elem |> tryElement (w + "tblPr") |> extractProperties
                        TrPr = elem |> tryElement (w + "trPr") |> extractPropertiesWithPriority // Use the new function
                        TcPr = elem |> tryElement (w + "tcPr") |> extractProperties
                    }
                    Map.add typeValue group acc
                | None -> acc) Map.empty

        // Merge parent properties with current properties
        let mergedTopLevel = {
            RPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.RPr parentProperties.TopLevel.RPr
            PPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.PPr parentProperties.TopLevel.PPr
            TblPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.TblPr parentProperties.TopLevel.TblPr
            TrPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.TrPr parentProperties.TopLevel.TrPr
            TcPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.TcPr parentProperties.TopLevel.TcPr
        }

        let mergedByType =
            Map.fold (fun (acc: Map<string, TblStylePrGroup>) key value ->
                if acc.ContainsKey key then acc
                else Map.add key value acc) thisByType parentProperties.ByType

        { TopLevel = mergedTopLevel; ByType = mergedByType }