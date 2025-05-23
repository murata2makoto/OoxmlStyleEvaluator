///<summary>This module defines accessors for run properties.  
///Ideally, we need at least one accessor for every run property.
///</summary>
module OoxmlStyleEvaluator.RunPropertyAccessors

open OoxmlStyleEvaluator
open System.Xml.Linq
open RunPropertiesEvaluator

type StyleEvaluator with

    /// <summary>
    /// Resolves the effective value of a run property for a given run element,
    /// considering direct formatting, character/paragraph/table styles, and document defaults.
    /// </summary>
    /// <param name="run">The run (<c>&lt;w:r&gt;</c>) element to evaluate.</param>
    /// <param name="key">The property key to retrieve (e.g., <c>"b/@val"</c>).</param>
    /// <returns>The resolved property value as <c>Some string</c> or <c>None</c>.</returns>
    member this.ResolveRunProperty(run: XElement, key: string) : string option =
        resolveEffectiveRunProperty
            run
            this.DocDefaultsRPr
            this.TableStyleResolver
            this.ParagraphStyleResolver
            this.CharacterStyleResolver
            this.GetPropertyByKey
            key

    /// <summary>Gets a string run property value with a default.</summary>
    member this.GetRunProperty(run: XElement, key: string, defaultValue: string) : string =
        this.ResolveRunProperty(run, key) |> Option.defaultValue defaultValue

    /// <summary>Gets a boolean run property value (e.g., bold or italic).</summary>
    member this.GetRunPropertyBool(run: XElement, key: string) : bool =
        this.ResolveRunProperty(run, key) |> Option.map System.Boolean.Parse |> Option.defaultValue false

    /// <summary>Gets an integer run property value (e.g., spacing).</summary>
    member this.GetRunPropertyInt(run: XElement, key: string) : int =
        this.ResolveRunProperty(run, key) |> Option.map int |> Option.defaultValue 0

    /// <summary>Gets the effective emphasis mark (e.g., "dot", "comma") for the run.</summary>
    member this.getEmphasisMark(run: XElement) : string =
        this.GetRunProperty(run, "em/@val", "none")

    /// <summary>Gets the effective font name for a given script range (e.g., ascii, cs).</summary>
    /// <param name="run">The run (<c>&lt;w:r&gt;</c>) element.</param>
    /// <param name="script">The script attribute (e.g., "ascii", "cs").</param>
    member this.GetEffectiveFont(run: XElement, script: string) : string =
        match script with
        | "ascii" | "asciiTheme" | "eastAsia" | "eastAsiaTheme"
        | "hAnsi" | "hAnsiTheme" | "cs" | "csTheme" ->
            this.GetRunProperty(run, $"rFonts/@{script}", "")
        | _ -> failwith $"Script {script} is not supported"

    /// <summary>Gets the effective ASCII font name.</summary>
    member this.getAsciiFont(run: XElement) = this.GetEffectiveFont(run, "ascii")

    /// <summary>Gets the effective ASCII theme font name.</summary>
    member this.getAsciiThemeFont(run: XElement) = this.GetEffectiveFont(run, "asciiTheme")

    /// <summary>Gets the effective HAnsi font name.</summary>
    member this.getHAnsiFont(run: XElement) = this.GetEffectiveFont(run, "hAnsi")

    /// <summary>Gets the effective HAnsi theme font name.</summary>
    member this.getHAnsiThemeFont(run: XElement) = this.GetEffectiveFont(run, "hAnsiTheme")

    /// <summary>Gets the effective East Asian font name.</summary>
    member this.getEastAsiaFont(run: XElement) = this.GetEffectiveFont(run, "eastAsia")

    /// <summary>Gets the effective East Asian theme font name.</summary>
    member this.getEastAsiaThemeFont(run: XElement) = this.GetEffectiveFont(run, "eastAsiaTheme")

    /// <summary>Gets the color value (e.g., "auto", "FF0000").</summary>
    member this.getColorValue(run: XElement) = this.GetRunProperty(run, "color/@val", "auto")

    /// <summary>Gets the theme color (e.g., "accent1").</summary>
    member this.getThemeColor(run: XElement) = this.GetRunProperty(run, "color/@themeColor", "none")

    /// <summary>Gets the font size (in half-points).</summary>
    member this.getSz(run: XElement) = this.GetRunProperty(run, "sz/@val", "")

    /// <summary>Gets the character spacing value.</summary>
    member this.getSpacingValue(run: XElement) = this.GetRunPropertyInt(run, "spacing/@val")

    /// <summary>Returns whether the run is italic.</summary>
    member this.isItalic(run: XElement) = this.GetRunPropertyBool(run, "i/@val")

    /// <summary>Returns whether the run is bold.</summary>
    member this.isBold(run: XElement) = this.GetRunPropertyBool(run, "b/@val")

    /// <summary>Gets the underline type (e.g., "single", "double").</summary>
    member this.getUnderlineType(run: XElement) = this.GetRunProperty(run, "u/@val", "none")

    /// <summary>Returns whether the run is struck through.</summary>
    member this.IsRunStrike(run: XElement) = this.GetRunPropertyBool(run, "strike/@val")

    /// <summary>Returns whether the run is in all caps or small caps.</summary>
    member this.IsRunCaps(run: XElement) = this.GetRunPropertyBool(run, "caps/@val")

    /// <summary>Gets the shading pattern of the run .</summary>
    member this.GetRunShadingPattern(run: XElement) = this.GetRunProperty(run, "shd/@fill", "none")
