///<summary>This module defines accessors for table row properties.  
///Ideally, we need at least one accessor for every table row property.
///</summary>
module OoxmlStyleEvaluator.TableRowPropertyAccessors

open System.Xml.Linq
open TableRowPropertiesEvaluator
open OoxmlStyleEvaluator
        
let resolveRowProperty (state: StyleEvaluatorState) (row: XElement) (key: string) : string option =
    resolveEffectiveRowProperty
        row
        state.TableStyleResolver
        state.GetPropertyByKey
        key
            
let getRowProperty state row key defaultValue =
    resolveRowProperty state row key 
    |> Option.defaultValue defaultValue

let getRowPropertyBool state row key defaultValue =
    resolveRowProperty state row key 
    |> Option.map System.Boolean.Parse 
    |> Option.defaultValue defaultValue

let getRowPropertyInt state row key defaultValue =
    resolveRowProperty state row key 
    |> Option.map int 
    |> Option.defaultValue defaultValue
