///<summary>This module defines accessors for table properties.  
///Ideally, we need at least one accessor for every table property.
///</summary>
namespace OoxmlStyleEvaluator
module TablePropertyAccessors =

    open System.Xml.Linq
    open TablePropertiesEvaluator
    open OoxmlStyleEvaluator
        
    let resolveTableProperty (state: StyleEvaluatorState) (table: XElement) (key: string) : string option =
        resolveEffectiveTableProperty
            table
            state.TableStyleResolver
            state.GetPropertyByKey
            key

            
    let getTableProperty state row key defaultValue =
        resolveTableProperty state row key |> Option.defaultValue defaultValue

    let getTablePropertyBool state row key defaultValue =
        resolveTableProperty state row key |> Option.map System.Boolean.Parse |> Option.defaultValue defaultValue

    let getTablePropertyInt state row key defaultValue =
        resolveTableProperty state row key |> Option.map int |> Option.defaultValue defaultValue