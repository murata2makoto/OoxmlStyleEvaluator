module OoxmlStyleEvaluator.TableStyleInheritance

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes

/// <summary>
/// Creates a resolver function that returns the fully resolved TableStyleProperties for a given table style ID,
/// following the inheritance chain via the `basedOn` attribute. Results are memoized for performance.
/// </summary>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>A function that takes a style ID and returns a TableStyleProperties object.</returns>
let createTableStyleResolver (stylesRoot: XElement) : (string -> TableStyleProperties) =
    let memo = System.Collections.Generic.Dictionary<string, TableStyleProperties>()

    /// Extracts properties from an element, prioritizing later occurrences of elements with the same QName.
    let extractPropertiesWithPriority (elementOpt: XElement option) : Map<string, XElement> =
        match elementOpt with
        | Some element ->
            element.Elements()
            |> Seq.rev // Later elements override earlier ones
            |> Seq.fold (fun acc elem -> Map.add (elem.Name.ToString()) elem acc) Map.empty
        | None -> Map.empty

    let rec resolve (styleId: string) : TableStyleProperties =
        match memo.TryGetValue styleId with
        | true, cached -> cached
        | false, _ ->
            match findStyle stylesRoot "table" styleId with
            | None ->
                let empty = {
                    TopLevel = { RPr = Map.empty; PPr = Map.empty; TblPr = Map.empty; TrPr = Map.empty; TcPr = Map.empty }
                    ByType = Map.empty
                }
                memo.[styleId] <- empty
                empty
            | Some style ->
                // Resolve parent style via basedOn
                let basedOnId = resolveBasedOn style
                let parentProperties =
                    basedOnId
                    |> Option.map resolve
                    |> Option.defaultValue {
                        TopLevel = { RPr = Map.empty; PPr = Map.empty; TblPr = Map.empty; TrPr = Map.empty; TcPr = Map.empty }
                        ByType = Map.empty
                    }

                // Extract current style's top-level properties
                let thisTopLevel = {
                    RPr = style |> tryElement (w + "rPr") |> extractProperties
                    PPr = style |> tryElement (w + "pPr") |> extractProperties
                    TblPr = style |> tryElement (w + "tblPr") |> extractProperties
                    TrPr = style |> tryElement (w + "trPr") |> extractProperties
                    TcPr = style |> tryElement (w + "tcPr") |> extractProperties
                }

                // Extract tblStylePr[@type] elements
                let thisByType =
                    style.Elements(w + "tblStylePr")
                    |> Seq.fold (fun acc elem ->
                        match tryAttrValue (w + "type") elem with
                        | Some typeValue ->
                            let group = {
                                RPr = elem |> tryElement (w + "rPr") |> extractProperties
                                PPr = elem |> tryElement (w + "pPr") |> extractProperties
                                TblPr = elem |> tryElement (w + "tblPr") |> extractProperties
                                TrPr = elem |> tryElement (w + "trPr") |> extractProperties
                                TcPr = elem |> tryElement (w + "tcPr") |> extractProperties
                            }
                            Map.add typeValue group acc
                        | None -> acc
                    ) Map.empty

                // Merge parent and current properties (current takes precedence)
                let mergedTopLevel = {
                    RPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.RPr parentProperties.TopLevel.RPr
                    PPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.PPr parentProperties.TopLevel.PPr
                    TblPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.TblPr parentProperties.TopLevel.TblPr
                    TrPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.TrPr parentProperties.TopLevel.TrPr
                    TcPr = Map.fold (fun acc k v -> if acc.ContainsKey k then acc else Map.add k v acc) thisTopLevel.TcPr parentProperties.TopLevel.TcPr
                }

                let mergedByType =
                    Map.fold (fun (acc: Map<string,TblStylePrGroup>) key value ->
                        if acc.ContainsKey key then acc
                        else Map.add key value acc
                    ) thisByType parentProperties.ByType

                let result = { TopLevel = mergedTopLevel; ByType = mergedByType }
                memo.[styleId] <- result
                result

    resolve
