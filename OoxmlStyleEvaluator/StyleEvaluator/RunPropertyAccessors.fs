///<summary>This module defines accessors for run properties.  
///Ideally, we need at least one accessor for every run property.
///</summary>
namespace OoxmlStyleEvaluator
module RunPropertyAccessors =

    open System.Xml.Linq
    open RunPropertiesEvaluator
    open OoxmlStyleEvaluator

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

    let getRunPropertyBool state run key =
        resolveRunProperty state run key |> Option.map System.Boolean.Parse |> Option.defaultValue false

    let getRunPropertyInt state run key =
        resolveRunProperty state run key |> Option.map int |> Option.defaultValue 0

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
    let getSpacingValue state run = getRunPropertyInt state run "spacing/@val"
    let isItalic state run = getRunPropertyBool state run "i/@val"
    let isBold state run = getRunPropertyBool state run "b/@val"
    let getUnderlineType state run = getRunProperty state run "u/@val" "none"
    let isRunStrike state run = getRunPropertyBool state run "strike/@val"
    let isRunCaps state run = getRunPropertyBool state run "caps/@val"
    let getRunShadingPattern state run = getRunProperty state run "shd/@fill" "none"
