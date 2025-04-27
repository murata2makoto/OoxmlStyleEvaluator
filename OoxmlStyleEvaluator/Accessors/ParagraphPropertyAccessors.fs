///<summary>This module defines accessors for paragraph properties.  
///Ideally, we need at least one accessor for every paragraph property.
///</summary>
module OoxmlStyleEvaluator.ParagraphPropertyAccessors

open System.Xml.Linq
open XmlHelpers
open PropertyTypes
open ParagraphPropertiesEvaluator

/// <summary>
/// Determines whether a given paragraph is a heading in a TOC,
/// based on the presence of an effective outline level (outlineLvl).
/// Fully respects style inheritance and document defaults.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>True if the paragraph is a heading; otherwise, false.</returns>
let isHeadingInTOC (para: XElement) (stylesRoot: XElement) : bool =
    resolveEffectiveParagraphProperties para stylesRoot
    |> getProperty "outlineLvl"
    |> Option.isSome
    
/// <summary>
/// Retrieves the TOC heading level (outlineLvl) of a given paragraph,
/// fully respecting style inheritance and document defaults.
/// </summary>
/// <param name="para">The paragraph (w:p) element to evaluate.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>The heading level as an integer (0-8) if specified; otherwise, None.</returns>
let getHeadingLevelInTOC (para: XElement) (stylesRoot: XElement) : int option =
    resolveEffectiveParagraphProperties para stylesRoot
    |> getProperty "outlineLvl"
    |> Option.bind (tryAttrValue (w + "val"))
    |> Option.map int

/// <summary>
/// Determines whether the given paragraph is justified.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// True if the paragraph is justified (e.g., "distribute", "end", or "both" values for the `jc` property); 
/// otherwise, false.
/// </returns>
let ifJustified (para: XElement) (stylesRoot: XElement) : bool  = 
    resolveEffectiveParagraphProperties para stylesRoot
    |> getProperty "jc"
    |> Option.bind (tryAttrValue (w + "val"))
    |> Option.defaultValue "none" //Does the presese of the jc element always imply justification? 
    |> (fun x -> x = "distribute" || x = "end" || x = "both")

/// <summary>
/// Determines whether the given paragraph is numbered.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// True if the paragraph has an associated numPr property having w:ilvl; 
/// otherwise, false.
/// </returns>    
let isNumbered (para: XElement) (stylesRoot: XElement) : bool  = 
    resolveEffectiveParagraphProperties para stylesRoot
    |> getProperty "numPr"
    |> Option.bind (tryElement (w+"ilvl"))
    |> Option.isSome

/// <summary>
/// Get the numbering level of the paragraph.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// Get the numbering level of the paragraph; if it is not numbered, None.
/// </returns>    
let getNumbereingLevel (para: XElement) (stylesRoot: XElement) : int option  = 
    resolveEffectiveParagraphProperties para stylesRoot
    |> getProperty "numPr"
    |> Option.bind (tryElement (w+"ilvl"))
    |> Option.bind (tryAttrValue (w+"val"))
    |> Option.map int

/// <summary>
/// Get a numbering definition instance reference.
/// </summary>
/// <param name="para">The paragraph (w:p) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>
/// Get a numbering definition instance reference of the paragraph; if it not numbered or has "0" as a reference, None.
/// </returns>    
let getNumbereingDefinitionInstance (para: XElement) (stylesRoot: XElement) : int option  = 
    resolveEffectiveParagraphProperties para stylesRoot
    |> getProperty "numPr"
    |> Option.bind (tryElement (w+"numId"))
    |> Option.bind (tryAttrValue (w+"val"))
    |> function
        | Some(x) -> 
            let i = int x 
            if i = 0 then None else Some(i)
        | None -> None
