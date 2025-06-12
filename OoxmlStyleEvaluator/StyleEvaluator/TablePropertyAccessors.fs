///<summary>This module defines accessors for table properties.  
///Ideally, we need at least one accessor for every table property.
///</summary>
module OoxmlStyleEvaluator.TablePropertyAccessors

open System.Xml.Linq
open TablePropertiesEvaluator
open OoxmlStyleEvaluator
open StyleUtilities
        
let resolveTableProperty (state: StyleEvaluatorState) (table: XElement) (key: string) : string option =
    resolveEffectiveTableProperty
        table
        state.TableStyleResolver
        state.GetPropertyByKey
        key

            
let getTableProperty state table key defaultValue =
    resolveTableProperty state table key 
    |> Option.defaultValue defaultValue

let getTablePropertyBool state table key defaultValue =
    resolveTableProperty state table key 
    |> Option.map System.Boolean.Parse 
    |> Option.defaultValue defaultValue

let getTablePropertyInt state table key defaultValue =
    resolveTableProperty state table key 
    |> Option.map int 
    |> Option.defaultValue defaultValue

let tryGetTblStyleId state table = 
    tryGetStyleId table "tblPr" "tblStyle" 
    |> Option.defaultValue ""