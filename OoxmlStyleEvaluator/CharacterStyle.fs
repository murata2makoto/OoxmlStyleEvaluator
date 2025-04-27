module OoxmlStyleEvaluator.CharacterStyle

open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers
open OoxmlStyleEvaluator.RunStyleEvaluator
open OoxmlStyleEvaluator.DocumentDefaults

/// <summary>
/// Checks if a boolean run property (e.g., bold, italic) is effectively enabled on the given run element.
/// </summary>
let isRunPropertyEnabled (propName: string) (run: XElement) (stylesRoot: XElement) : bool =
    let docDefaults = getDocDefaultsRPr stylesRoot
    evaluateEffectiveProperty ((w + propName).ToString()) run stylesRoot docDefaults None
    |> Option.isSome

/// <summary>
/// Determines whether the run is effectively bold.
/// </summary>
let isRunBold run stylesRoot = isRunPropertyEnabled "b" run stylesRoot

/// <summary>
/// Determines whether the run is effectively italic.
/// </summary>
let isRunItalic run stylesRoot = isRunPropertyEnabled "i" run stylesRoot

/// <summary>
/// Determines whether the run has effective strikethrough applied.
/// </summary>
let isRunStrike run stylesRoot = isRunPropertyEnabled "strike" run stylesRoot

/// <summary>
/// Determines whether the run is displayed in all capital letters.
/// </summary>
let isRunCaps run stylesRoot = isRunPropertyEnabled "caps" run stylesRoot

/// <summary>
/// Determines whether the run is displayed in small caps.
/// </summary>
let isRunSmallCaps run stylesRoot = isRunPropertyEnabled "smallCaps" run stylesRoot

/// <summary>
/// Gets the underline style of the run (e.g., "single", "double", "none").
/// </summary>
let getEffectiveUnderlineType (run: XElement) (stylesRoot: XElement) : string =
    let docDefaults = getDocDefaultsRPr stylesRoot
    match evaluateEffectiveProperty ((w + "u").ToString()) run stylesRoot docDefaults None with
    | Some el ->
        tryAttrValue (w + "val") el |> Option.defaultValue "single"
    | None -> "none"

/// <summary>
/// Gets the emphasis mark style applied to the run (e.g., "dot", "comma", "none").
/// </summary>
let getEffectiveEmphasisMark (run: XElement) (stylesRoot: XElement) : string =
    let docDefaults = getDocDefaultsRPr stylesRoot
    match evaluateEffectiveProperty ((w + "em").ToString()) run stylesRoot docDefaults None with
    | Some el ->
        tryAttrValue (w + "val") el |> Option.defaultValue "dot"
    | None -> "none"

/// <summary>
/// Gets the effective text color of the run (e.g., "FF0000", "auto").
/// </summary>
let getEffectiveColor (run: XElement) (stylesRoot: XElement) : string =
    let docDefaults = getDocDefaultsRPr stylesRoot
    match evaluateEffectiveProperty ((w + "color").ToString()) run stylesRoot docDefaults None with
    | Some el ->
        tryAttrValue (w + "val") el |> Option.defaultValue "auto"
    | None -> "auto"

/// <summary>
/// Gets the effective font name for ASCII characters (e.g., "Times New Roman").
/// </summary>
let getEffectiveFont (run: XElement) (stylesRoot: XElement) : string =
    let docDefaults = getDocDefaultsRPr stylesRoot
    match evaluateEffectiveProperty ((w + "rFonts").ToString()) run stylesRoot docDefaults None with
    | Some el ->
        tryAttrValue (w + "ascii") el |> Option.defaultValue "Times New Roman"
    | None -> "Times New Roman"
