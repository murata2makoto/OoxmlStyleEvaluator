module OoxmlStyleEvaluator.RunDirectProperties

open System.Xml.Linq
open XmlHelpers
open StyleUtilities
open PropertyTypes

/// <summary>
/// Resolves the effective run properties by combining directly specified `rPr`.
/// </summary>
/// <param name="runElement">The run (`<w:r>`) element to process.</param>
/// <returns>A map of effective run properties.</returns>
let resolveDirectRunProperties (runElement: XElement) : RPr =

    // Extract rPr (directly specified run properties)
    runElement |> tryElement (w + "rPr") |> extractProperties


