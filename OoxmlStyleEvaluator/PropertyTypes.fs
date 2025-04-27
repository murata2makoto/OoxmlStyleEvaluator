module OoxmlStyleEvaluator.PropertyTypes

open System.Xml.Linq

/// <summary>
/// A type alias for paragraph properties, represented as a map from fully qualified element names to their XElement definitions.
/// </summary>
type PPr = Map<string, XElement>

/// <summary>
/// A type alias for run properties, represented as a map from fully qualified element names to their XElement definitions.
/// </summary>
type RPr = Map<string, XElement>

/// <summary>
/// A type alias for table properties, represented as a map from fully qualified element names to their XElement definitions.
/// </summary>
type TblPr = Map<string, XElement>

/// <summary>
/// A type alias for table cell properties, represented as a map from fully qualified element names to their XElement definitions.
/// </summary>
type TcPr = Map<string, XElement>

/// <summary>
/// A type alias for table row properties, represented as a map from fully qualified element names to their XElement definitions.
/// </summary>
type TrPr = Map<string, XElement>

/// <summary>
/// Represents the properties for a specific `tblStylePr[@type]`.
/// </summary>
type TblStylePrGroup = {
    RPr: RPr
    PPr: PPr
    TblPr: TblPr
    TrPr: TrPr
    TcPr: TcPr
}

/// <summary>
/// Represents the full set of table style properties, including top-level properties and `tblStylePr[@type]` groups.
/// </summary>
type TableStyleProperties = {
    TopLevel: TblStylePrGroup
    ByType: Map<string, TblStylePrGroup>
}

/// <summary>
/// Merges two property maps, giving priority to the second map.
/// </summary>
/// <param name="baseMap">The base map of properties.</param>
/// <param name="overrideMap">The map of properties that overrides the base map.</param>
/// <returns>A merged map where properties in `overrideMap` take precedence.</returns>
let mergeProperties (baseMap: Map<string, XElement>) (overrideMap: Map<string, XElement>) : Map<string, XElement> =
    Map.fold (fun acc k v -> Map.add k v acc) baseMap overrideMap

let getProperty (propertyName: string) (map: Map<string, XElement>) : XElement option =
    map.TryFind propertyName