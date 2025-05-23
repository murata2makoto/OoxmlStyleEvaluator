///<summary>This module defines accessors for paragraph properties.
///Ideally, we need at least one accessor for every paragraph property.
///</summary>
namespace OoxmlStyleEvaluator

module ParagraphPropertyAccessors =

    open OoxmlStyleEvaluator
    open System.Xml.Linq
    open ParagraphPropertiesEvaluator

    type StyleEvaluator with

        /// <summary>
        /// Resolves the effective value of a paragraph property for a given paragraph element,
        /// taking into account direct formatting, paragraph style, table style, and document defaults.
        /// </summary>
        /// <param name="para">The paragraph (<c>&lt;w:p&gt;</c>) element.</param>
        /// <param name="key">The property key to retrieve (e.g., <c>"spacing/@before"</c>).</param>
        /// <returns>The resolved property value as <c>Some string</c> or <c>None</c>.</returns>
        member public this.ResolveParagraphProperty(para: XElement, key: string) : string option =
            resolveEffectiveParagraphProperty
                para
                this.DocDefaultsPPr
                this.TableStyleResolver
                this.ParagraphStyleResolver
                this.GetPropertyByKey
                key

        /// <summary>Gets a string paragraph property value with a default.</summary>
        member this.GetParagraphProperty(para: XElement, key: string, defaultValue: string) : string =
            this.ResolveParagraphProperty(para, key) 
            |> Option.defaultValue defaultValue

        /// <summary>Gets a boolean paragraph property value.</summary>
        member this.GetParagraphPropertyBool(para: XElement, key: string, defaultValue: bool) : bool =
            this.ResolveParagraphProperty(para, key) |> Option.map System.Boolean.Parse 
            |> Option.defaultValue defaultValue

        /// <summary>Gets an integer paragraph property value.</summary>
        member this.GetParagraphPropertyInt(para: XElement, key: string, defaultValue: int) : int =
            this.ResolveParagraphProperty(para, key) |> Option.map int 
            |> Option.defaultValue defaultValue

        /// <summary>Gets the outline level (heading level) of the paragraph.</summary>
        member this.GetHeadingLevel(para: XElement) : int =
            this.GetParagraphPropertyInt(para, "outlineLvl/@val", -1)

        /// <summary>Gets the numbering ID for the paragraph (if any).</summary>
        member this.GetNumId(para: XElement) : int =
            this.GetParagraphPropertyInt(para, "numPr/numId/@val", -1)

        /// <summary>Gets the numbering level for the paragraph (if any).</summary>
        member this.GetNumLevel(para: XElement) : int =
            this.GetParagraphPropertyInt(para, "numPr/ilvl/@val", -1)

        /// <summary>Determines whether the paragraph is a heading (not a list bullet).</summary>
        member public this.IsHeadingParagraph(para: XElement) : bool =
            this.GetHeadingLevel(para) >= 0

        /// <summary>Determines whether the paragraph is a list bullet (not a heading).</summary>
        member this.IsBulletParagraph(para: XElement) : bool =
            this.GetHeadingLevel(para) = -1 &&
            this.GetNumId(para) <> -1

        /// <summary>Gets the justification type (e.g., "start", "center"); if not, an empty string.</summary>
        member this.GetJustificationType(para: XElement) : string =
            this.GetParagraphProperty(para, "jc/@val", "")

        /// <summary>Gets the character spacing value (in twentieths of     a point).</summary>
        member this.GetCharacterSpacing(para: XElement) : int =
            this.GetParagraphPropertyInt(para, "spacing/@val", 0)

