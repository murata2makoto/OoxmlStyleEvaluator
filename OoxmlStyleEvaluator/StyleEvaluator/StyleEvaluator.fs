namespace OoxmlStyleEvaluator

open System.IO.Compression
open System.Xml.Linq
open DocumentDefaults
open CharacterStyleInheritance
open ParagraphStyleInheritance
open TableStyleInheritance
open LoadFromArchive
open PropertyExtractor
open ParagraphPropertyAccessors
open RunPropertyAccessors
open TableCellPropertyAccessors
open TableRowPropertyAccessors
open TablePropertyAccessors



/// <summary>
/// Provides evaluation utilities for WordprocessingML styles and numbering inside a DOCX/OOXML document.
/// </summary>
type public StyleEvaluator(archive: ZipArchive) =

    let documentRoot = loadDocumentXml archive
    let stylesRoot : XElement = loadStyles archive

    let state =
        {
            StylesRoot = stylesRoot
            DocDefaultsRPr = getRPrDocDefaults stylesRoot
            DocDefaultsPPr = getPPrDocDefaults stylesRoot
            CharacterStyleResolver = createCharacterStyleResolver stylesRoot
            ParagraphStyleResolver = createParagraphStyleResolver stylesRoot
            TableStyleResolver = createTableStyleResolver stylesRoot
            GetPropertyByKey = makeGetPropertyByKey()
        }

    /// <summary>Gets the loaded document.xml root element.</summary>
    member public this.DocumentRoot = documentRoot

    // -------- Paragraph Properties --------

    /// <summary>
    /// Resolves the effective value of a paragraph property for a given paragraph element,
    /// taking into account direct formatting, paragraph style, table style, and document defaults.
    /// </summary>
    member public this.ResolveParagraphProperty(para, key) = resolveParagraphProperty state para key

    /// <summary>Gets a string paragraph property value with a default.</summary>
    member public this.GetParagraphProperty(para, key, defaultValue) = getParagraphProperty state para key defaultValue

    /// <summary>Gets a boolean paragraph property value.</summary>
    member public this.GetParagraphPropertyBool(para, key, defaultValue) = getParagraphPropertyBool state para key defaultValue

    /// <summary>Gets an integer paragraph property value.</summary>
    member public this.GetParagraphPropertyInt(para, key, defaultValue) = getParagraphPropertyInt state para key defaultValue

    /// <summary>Gets the outline level (heading level) of the paragraph.</summary>
    member public this.GetHeadingLevel(para) = getHeadingLevel state para

    /// <summary>Gets the numbering ID for the paragraph (if any).</summary>
    member public this.GetNumId(para) = getNumId state para

    /// <summary>Gets the numbering level for the paragraph (if any).</summary>
    member public this.GetNumLevel(para) = getNumLevel state para

    /// <summary>Determines whether the paragraph is a heading (not a list bullet).</summary>
    member public this.IsHeadingParagraph(para) = isHeadingParagraph state para

    /// <summary>Determines whether the paragraph is a list bullet (not a heading).</summary>
    member public this.IsBulletParagraph(para) = isBulletParagraph state para

    /// <summary>Gets the justification type (e.g., "start", "center"); if not, an empty string.</summary>
    member public this.GetJustificationType(para) = getJustificationType state para

    /// <summary>Gets the character spacing value (in twentieths of a point).</summary>
    member public this.GetCharacterSpacing(para) = getCharacterSpacing state para

    // -------- Run Properties --------

    /// <summary>
    /// Resolves the effective value of a run property for a given run element,
    /// considering direct formatting, character/paragraph/table styles, and document defaults.
    /// </summary>
    member public this.ResolveRunProperty(run, key) = resolveRunProperty state run key

    /// <summary>Gets a string run property value with a default.</summary>
    member public this.GetRunProperty(run, key, defaultValue) = getRunProperty state run key defaultValue

    /// <summary>Gets a boolean run property value (e.g., bold or italic).</summary>
    member public this.GetRunPropertyBool(run, key, defaultValue) = getRunPropertyBool state run key defaultValue

    /// <summary>Gets an integer run property value (e.g., spacing).</summary>
    member public this.GetRunPropertyInt(run, key, defaultValue) = getRunPropertyInt state run key defaultValue

    /// <summary>Gets the effective emphasis mark (e.g., "dot", "comma") for the run.</summary>
    member public this.GetEmphasisMark(run) = getEmphasisMark state run

    /// <summary>Gets the effective font name for a given script range (e.g., ascii, cs).</summary>
    member public this.GetEffectiveFont(run, script) = getEffectiveFont state run script

    /// <summary>Gets the effective ASCII font name.</summary>
    member public this.GetAsciiFont(run) = getAsciiFont state run

    /// <summary>Gets the effective ASCII theme font name.</summary>
    member public this.GetAsciiThemeFont(run) = getAsciiThemeFont state run

    /// <summary>Gets the effective HAnsi font name.</summary>
    member public this.GetHAnsiFont(run) = getHAnsiFont state run

    /// <summary>Gets the effective HAnsi theme font name.</summary>
    member public this.GetHAnsiThemeFont(run) = getHAnsiThemeFont state run

    /// <summary>Gets the effective East Asian font name.</summary>
    member public this.GetEastAsiaFont(run) = getEastAsiaFont state run

    /// <summary>Gets the effective East Asian theme font name.</summary>
    member public this.GetEastAsiaThemeFont(run) = getEastAsiaThemeFont state run

    /// <summary>Gets the color value (e.g., "auto", "FF0000").</summary>
    member public this.GetColorValue(run) = getColorValue state run

    /// <summary>Gets the theme color (e.g., "accent1").</summary>
    member public this.GetThemeColor(run) = getThemeColor state run

    /// <summary>Gets the font size (in half-points).</summary>
    member public this.GetSz(run) = getSz state run

    /// <summary>Gets the character spacing value.</summary>
    member public this.GetSpacingValue(run) = getSpacingValue state run

    /// <summary>Returns whether the run is italic.</summary>
    member public this.IsItalic(run) = isItalic state run

    /// <summary>Returns whether the run is bold.</summary>
    member public this.IsBold(run) = isBold state run

    /// <summary>Gets the underline type (e.g., "single", "double").</summary>
    member public this.GetUnderlineType(run) = getUnderlineType state run

    /// <summary>Returns whether the run is struck through.</summary>
    member public this.IsRunStrike(run) = isRunStrike state run

    /// <summary>Returns whether the run is in all caps or small caps.</summary>
    member public this.IsRunCaps(run) = isRunCaps state run

    /// <summary>Gets the shading pattern of the run.</summary>
    member public this.GetRunShadingPattern(run) = getRunShadingPattern state run
    
    // -------- Cell Properties --------

    /// <summary>
    /// Resolves the effective value of a run property for a given cell element,
    /// considering direct formatting, character/paragraph/table styles, and document defaults.
    /// </summary>
    member public this.ResolveCellProperty(cell, key) = resolveCellProperty state cell key

    // -------- Row Properties --------

    /// <summary>
    /// Resolves the effective value of a run property for a given cell element,
    /// considering direct formatting, character/paragraph/table styles, and document defaults.
    /// </summary>
    member public this.ResolvRowProperty(row, key) = resolveRowProperty state row key

    // -------- Table Properties --------

    /// <summary>
    /// Resolves the effective value of a run property for a given cell element,
    /// considering direct formatting, character/paragraph/table styles, and document defaults.
    /// </summary>
    member public this.ResolveTableProperty(tbl, key) = resolveTableProperty state tbl key