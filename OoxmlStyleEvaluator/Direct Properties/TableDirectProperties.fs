module OoxmlStyleEvaluator.TableDirectProperties

open System.Xml.Linq
open XmlHelpers
open StyleUtilities

/// <summary>
/// Resolves the table properties (`tblPr`) specified for a table (`<w:tbl>`).
/// </summary>
/// <param name="tableElement">The table (`<w:tbl>`) element to process.</param>
/// <returns>A map of table properties (`tblPr`).</returns>
let resolveDirectTableProperties (tableElement: XElement): Map<string, XElement> =
    // Extract tblPr (table properties for the table)
    tableElement
    |> tryElement (w + "tblPr")
    |> extractProperties

/// <summary>
/// Resolves the row properties (`trPr`) specified for a table row (`<w:tr>`).
/// </summary>
/// <param name="tableRowElement">The table row (`<w:tr>`) element to process.</param>
/// <returns>A map of row properties (`trPr`).</returns>
let resolveDirectRowProperties (tableRowElement: XElement): Map<string, XElement> =
    // Extract trPr (row properties for the row)
    tableRowElement
    |> tryElement (w + "trPr")
    |> extractProperties

/// <summary>
/// Resolves the cell properties (`tcPr`) specified for a table cell (`<w:tc>`).
/// </summary>
/// <param name="cellElement">The table cell (`<w:tc>`) element to process.</param>
/// <returns>A map of cell properties (`tcPr`).</returns>
let resolveDirectCellProperties (cellElement: XElement): Map<string, XElement> =
    // Extract tcPr (cell properties for the table cell)
    cellElement
    |> tryElement (w + "tcPr")
    |> extractProperties

/// <summary>
/// Resolves the table properties (`tblPrEx`) specified for a table row (`<w:tr>`).
/// </summary>
/// <param name="tableRowElement">The table row (`<w:tr>`) element to process.</param>
/// <returns>A map of table properties (`tblPrEx`).</returns>
let resolveTablePropertiesForRow (tableRowElement: XElement): Map<string, XElement> =
    // Extract tblPrEx (table properties for the row)
    tableRowElement
    |> tryElement (w + "tblPrEx")
    |> extractProperties
