module OoxmlStyleEvaluator.CharacterDetection

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.RunStyleEvaluator

type RPr = Map<string, XElement>

/// <summary>
/// Determines whether a specific toggle run property is enabled for a given run.
/// </summary>
/// <param name="propName">The name of the property (e.g., "b" for bold).</param>
/// <param name="run">The run (w:r) element.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>True if the property is enabled; otherwise, false.</returns>
let isRunPropertyEnabled
    (propName: string)
    (run: XElement)
    (stylesRoot: XElement)
    (docDefaults: RPr)
    (tableStyleRPr: RPr option) : bool =

    match evaluateEffectiveProperty propName run stylesRoot docDefaults tableStyleRPr with
    | Some e ->
        match tryAttrValue (w + "val") e with
        | Some "false" | Some "0" -> false
        | _ -> true  // Default is true if val missing or "true"/"1"
    | None -> false
    /// <summary>
/// Determines if the run is bold.
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>True if bold is enabled; otherwise, false.</returns>
let isRunBold (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : bool =
    isRunPropertyEnabled "b" run stylesRoot docDefaults tableStyleRPr

/// <summary>
/// Determines if the run is italic.
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>True if italic is enabled; otherwise, false.</returns>
let isRunItalic (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : bool =
    isRunPropertyEnabled "i" run stylesRoot docDefaults tableStyleRPr

/// <summary>
/// Determines if the run is struck through (strikethrough).
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>True if strikethrough is enabled; otherwise, false.</returns>
let isRunStrike (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : bool =
    isRunPropertyEnabled "strike" run stylesRoot docDefaults tableStyleRPr

/// <summary>
/// Determines if the run uses capitalization (caps).
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>True if caps is enabled; otherwise, false.</returns>
let isRunCaps (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : bool =
    isRunPropertyEnabled "caps" run stylesRoot docDefaults tableStyleRPr

/// <summary>
/// Gets the effective underline type (e.g., "single", "double") for the run.
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>The underline type string if specified; otherwise, None.</returns>
let getEffectiveUnderlineType
    (run: XElement)
    (stylesRoot: XElement)
    (docDefaults: RPr)
    (tableStyleRPr: RPr option) : string option =

    evaluateEffectiveProperty "u" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "val"))

/// <summary>
/// Gets the effective emphasis mark (e.g., "dot", "comma") for the run.
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>The emphasis mark string if specified; otherwise, None.</returns>
let getEffectiveEmphasisMark
    (run: XElement)
    (stylesRoot: XElement)
    (docDefaults: RPr)
    (tableStyleRPr: RPr option) : string option =

    evaluateEffectiveProperty "em" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "val"))

/// <summary>
/// Gets the effective text color (hex RGB string) for the run.
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>The text color as a hex string (e.g., "FF0000") if specified; otherwise, None.</returns>
let getEffectiveColor
    (run: XElement)
    (stylesRoot: XElement)
    (docDefaults: RPr)
    (tableStyleRPr: RPr option) : string option =

    evaluateEffectiveProperty "color" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "val"))

/// <summary>
/// Gets the effective ASCII font name for the run (w:ascii).
/// </summary>
/// <param name="run">The run (w:r) element to check.</param>
/// <param name="stylesRoot">The root element of styles.xml.</param>
/// <param name="docDefaults">The document defaults for run properties.</param>
/// <param name="tableStyleRPr">Optional table style run properties.</param>
/// <returns>The ASCII font name if specified; otherwise, None.</returns>
let getEffectiveAsciiFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "ascii"))

/// <summary>
/// Gets the effective ASCII theme font name for the run (w:asciiTheme).
/// </summary>
let getEffectiveAsciiThemeFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "asciiTheme"))

/// <summary>
/// Gets the effective high ANSI font name for the run (w:hAnsi).
/// </summary>
let getEffectiveHAnsiFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "hAnsi"))

/// <summary>
/// Gets the effective high ANSI theme font name for the run (w:hAnsiTheme).
/// </summary>
let getEffectiveHAnsiThemeFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "hAnsiTheme"))

/// <summary>
/// Gets the effective East Asian font name for the run (w:eastAsia).
/// </summary>
let getEffectiveEastAsiaFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "eastAsia"))

/// <summary>
/// Gets the effective East Asian theme font name for the run (w:eastAsiaTheme).
/// </summary>
let getEffectiveEastAsiaThemeFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "eastAsiaTheme"))

/// <summary>
/// Gets the effective complex script font name for the run (w:cs).
/// </summary>
let getEffectiveCsFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "cs"))

/// <summary>
/// Gets the effective complex script theme font name for the run (w:cstheme).
/// </summary>
let getEffectiveCsThemeFont (run: XElement) (stylesRoot: XElement) (docDefaults: RPr) (tableStyleRPr: RPr option) : string option =
    evaluateEffectiveProperty "rFonts" run stylesRoot docDefaults tableStyleRPr
    |> Option.bind (tryAttrValue (w + "cstheme"))
