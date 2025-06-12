module OoxmlStyleEvaluator.ParagraphPropertyAccessors 

open System.Xml.Linq
open ParagraphPropertiesEvaluator
open StyleUtilities

let resolveParagraphProperty (state: StyleEvaluatorState) (para: XElement) (key: string) =
    resolveEffectiveParagraphProperty
        para
        state.DocDefaultsPPr
        state.TableStyleResolver
        state.ParagraphStyleResolver
        state.GetPropertyByKey
        key

let getParagraphProperty state para key defaultValue =
    resolveParagraphProperty state para key 
    |> Option.defaultValue defaultValue

let getParagraphPropertyBool state para key defaultValue =
    resolveParagraphProperty state para key 
    |> Option.map System.Boolean.Parse 
    |> Option.defaultValue defaultValue

let getParagraphPropertyInt state para key defaultValue =
    resolveParagraphProperty state para key 
    |> Option.map int 
    |> Option.defaultValue defaultValue

let getHeadingLevel state para =
    getParagraphPropertyInt state para "outlineLvl/@val" -1

let getNumId state para =
    getParagraphPropertyInt state para "numPr/numId/@val" -1

let getNumLevel state para =
    getParagraphPropertyInt state para "numPr/ilvl/@val" 0 //0 appears to be the default!

let isHeadingParagraph state para =
    getHeadingLevel state para >= 0

let isBulletParagraph state para =
    getHeadingLevel state para = -1 && getNumId state para <> -1

let getJustificationType state para =
    getParagraphProperty state para "jc/@val" ""

let getCharacterSpacing state para =
    getParagraphPropertyInt state para "spacing/@val" 0

let tryGetPStyleId state para =
     tryGetStyleId para "pPr" "pStyle"
     |> Option.defaultValue ""