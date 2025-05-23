namespace OoxmlStyleEvaluator

open System.Xml.Linq
open PropertyTypes

type StyleEvaluatorState =
    {
        StylesRoot: XElement
        DocDefaultsRPr: RPr
        DocDefaultsPPr: PPr
        CharacterStyleResolver: string -> RPr
        ParagraphStyleResolver: string -> RPr * PPr
        TableStyleResolver: string -> TableStyleProperties
        GetPropertyByKey: XElement -> string -> string option
    }