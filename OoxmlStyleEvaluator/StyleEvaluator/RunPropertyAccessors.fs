///<summary>This module defines accessors for run properties.  
///Ideally, we need at least one accessor for every run property.
///</summary>
module OoxmlStyleEvaluator.RunPropertyAccessors

open System.Xml.Linq
open RunPropertiesEvaluator
open OoxmlStyleEvaluator
open StyleUtilities

let resolveRunProperty (state: StyleEvaluatorState) (run: XElement) (key: string) =
    resolveEffectiveRunProperty
        run
        state.DocDefaultsRPr
        state.TableStyleResolver
        state.ParagraphStyleResolver
        state.CharacterStyleResolver
        state.GetPropertyByKey
        key

let getRunProperty state run key defaultValue =
    resolveRunProperty state run key |> Option.defaultValue defaultValue

let getRunPropertyBool state run key defaultValue =
    resolveRunProperty state run key |> Option.map System.Boolean.Parse |> Option.defaultValue defaultValue

let getRunPropertyInt state run key defaultValue =
    resolveRunProperty state run key |> Option.map int |> Option.defaultValue defaultValue

let getEmphasisMark state run =
    getRunProperty state run "em/@val" "none"

let getEffectiveFont state run script =
    match script with
    | "ascii" | "asciiTheme" | "eastAsia" | "eastAsiaTheme"
    | "hAnsi" | "hAnsiTheme" | "cs" | "csTheme" ->
        getRunProperty state run ($"rFonts/@{script}") ""
    | _ -> failwith ($"Script {script} is not supported")

let getAsciiFont state run = getEffectiveFont state run "ascii"
let getAsciiThemeFont state run = getEffectiveFont state run "asciiTheme"
let getHAnsiFont state run = getEffectiveFont state run "hAnsi"
let getHAnsiThemeFont state run = getEffectiveFont state run "hAnsiTheme"
let getEastAsiaFont state run = getEffectiveFont state run "eastAsia"
let getEastAsiaThemeFont state run = getEffectiveFont state run "eastAsiaTheme"
let getColorValue state run = getRunProperty state run "color/@val" "auto"
let getThemeColor state run = getRunProperty state run "color/@themeColor" "none"
let getSz state run = getRunProperty state run "sz/@val" ""
let getSpacingValue state run = getRunPropertyInt state run "spacing/@val" 0
let isItalic state run = getRunPropertyBool state run "i/@val" false
let isBold state run = getRunPropertyBool state run "b/@val" false
let getUnderlineType state run = getRunProperty state run "u/@val" "none"
let isRunStrike state run = getRunPropertyBool state run "strike/@val" false
let isRunCaps state run = getRunPropertyBool state run "caps/@val" false
let tryGetRStyleId state run = tryGetStyleId run "rPr" "rStyle" |> Option.defaultValue ""
let getRunShadingPattern state run = getRunProperty state run  "shd/@val" "nil"
let getVertAlign state run = getRunProperty state run "vertAlign/@val" "baseline"
