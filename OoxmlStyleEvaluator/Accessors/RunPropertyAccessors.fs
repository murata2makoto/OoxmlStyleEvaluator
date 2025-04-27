///<summary>This module defines accessors for run properties.  
///Ideally, we need at least one accessor for every run property.
///</summary>
module OoxmlStyleEvaluator.RunPropertyAccessors

open System.Xml.Linq
open XmlHelpers
open PropertyTypes
open RunPropertiesEvaluator

/// <summary>
/// Gets the effective emphasis mark (e.g., "dot", "comma") for the run.
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>The emphasis mark string if specified; otherwise, "none".</returns>
let getEmphasisMark (run: XElement) (stylesRoot: XElement) : string =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "em"
    |> Option.bind (tryAttrValue (w + "val"))
    |> Option.defaultValue "none"

/// <summary>
/// Gets the effective ASCII font name for the run (w:ascii).
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <returns>The ASCII font name if specified; otherwise, None.</returns>
let getAsciiFont (run: XElement) (stylesRoot: XElement) : string option =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "rFonts"
    |> Option.bind (tryAttrValue (w + "ascii"))

/// <summary>
/// Gets the effective ASCII theme font name for the run (w:asciiTheme).
/// </summary>
let getAsciiThemeFont (run: XElement) (stylesRoot: XElement): string option =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "rFonts"
    |> Option.bind (tryAttrValue (w + "asciiTheme"))

/// <summary>
/// Gets the effective high ANSI font name for the run (w:hAnsi).
/// </summary>
let getHAnsiFont (run: XElement) (stylesRoot: XElement): string option =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "rFonts"
    |> Option.bind (tryAttrValue (w + "hAnsi"))

/// <summary>
/// Gets the effective high ANSI theme font name for the run (w:hAnsiTheme).
/// </summary>
let getHAnsiThemeFont (run: XElement) (stylesRoot: XElement): string option =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "rFonts"
    |> Option.bind (tryAttrValue (w + "hAnsiTheme"))

/// <summary>
/// Gets the effective East Asian font name for the run (w:eastAsia).
/// </summary>
let getEastAsiaFont (run: XElement) (stylesRoot: XElement) : string option =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "rFonts"
    |> Option.bind (tryAttrValue (w + "eastAsia"))

/// <summary>
/// Gets the effective East Asian theme font name for the run (w:eastAsiaTheme).
/// </summary>
let getEastAsiaThemeFont (run: XElement) (stylesRoot: XElement) : string option =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "rFonts"
    |> Option.bind (tryAttrValue (w + "eastAsiaTheme"))


/// <summary>
/// Gets the value of the `val` attribute from the `color` property of the run.
/// </summary>
/// <param name="run">The run (`<w:r>`) element to check.</param>
/// <param name="stylesRoot">The root element of `styles.xml`.</param>
/// <returns>The value of the `val` attribute if specified; otherwise, None.</returns>
let getColorValue (run: XElement) (stylesRoot: XElement) : string  =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "color"
    |> Option.bind (tryAttrValue (w + "val"))
    |> Option.defaultValue "auto" //This is my guess

/// <summary>
/// Gets the value of the `theme` attribute from the `color` property of the run.
/// </summary>
/// <param name="run">The run (`<w:r>`) element to check.</param>
/// <param name="stylesRoot">The root element of `styles.xml`.</param>
/// <returns>The value of the `theme` attribute if specified; otherwise, None.</returns>
let getThemeColor (run: XElement) (stylesRoot: XElement) : string  =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "color"
    |> Option.bind (tryAttrValue (w + "theme"))
    |> Option.defaultValue "none" //This is my guess

/// <summary>
/// Gets the value of the `val` attribute from the `sz` property of the run.
/// </summary>
/// <param name="run">The run (`<w:r>`) element to check.</param>
/// <param name="stylesRoot">The root element of `styles.xml`.</param>
/// <returns>The value of the `val` attribute if specified; otherwise, None.</returns>
let getSz (run: XElement) (stylesRoot: XElement) : string option =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "sz"
    |> Option.bind (tryAttrValue (w + "val"))


/// <summary>
/// Gets the value of the `val` attribute from the `spacing` property of the run.
/// </summary>
/// <param name="run">The run (`<w:r>`) element to check.</param>
/// <param name="stylesRoot">The root element of `styles.xml`.</param>
/// <returns>The value of the `val` attribute if specified; otherwise, 0.</returns>
let getSpacingValue (run: XElement) (stylesRoot: XElement) : int =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "spacing"
    |> Option.bind (tryAttrValue (w + "val"))
    |> Option.map int
    |> Option.defaultValue 0

/// <summary>
/// Determines whether the `i` property is present in the run.
/// </summary>
/// <param name="run">The run (`<w:r>`) element to check.</param>
/// <param name="stylesRoot">The root element of `styles.xml`.</param>
/// <returns>True if the `i` property is present; otherwise, false.</returns>
let isItalic (run: XElement) (stylesRoot: XElement) : bool =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "i"
    |> Option.isSome

/// <summary>
/// Determines whether the `b` property is present in the run.
/// </summary>
/// <param name="run">The run (`<w:r>`) element to check.</param>
/// <param name="stylesRoot">The root element of `styles.xml`.</param>
/// <returns>True if the `b` property is present; otherwise, false.</returns>
let isBold (run: XElement) (stylesRoot: XElement) : bool =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "b"
    |> Option.isSome


/// <summary>
/// Gets the underline (e.g., "double", "single") for the run.
/// </summary>
/// <param name="run">The run (`<w:r>`) element to check.</param>
/// <param name="stylesRoot">The root element of `styles.xml`.</param>
/// <returns>The underline type if specified; otherwise, "none".</returns>
let getUnderlineValue (run: XElement) (stylesRoot: XElement) : string =
    resolveEffectiveRunProperties run stylesRoot
    |> getProperty "u"
    |> Option.bind (tryAttrValue (w + "val"))
    |> Option.defaultValue "none" //I don't believe that "auto" is the default.

    
let isUnderlined (run: XElement) (stylesRoot: XElement) : bool =
    getUnderlineValue run stylesRoot <> "none"

