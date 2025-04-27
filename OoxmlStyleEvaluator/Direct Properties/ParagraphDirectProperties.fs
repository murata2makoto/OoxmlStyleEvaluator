module OoxmlStyleEvaluator.ParagraphDirectProperties

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes

/// <summary>
/// Resolves the effective paragraph properties by combining directly specified `pPr`.
/// </summary>
/// <param name="paragraphElement">The paragraph (`<w:p>`) element to process.</param>
/// <returns>A map of effective paragraph properties.</returns>
let resolveDirectParagraphProperties (paragraphElement: XElement): PPr =

    // Extract pPr (directly specified paragraph properties)
    paragraphElement |> tryElement (w + "pPr") |> extractProperties


