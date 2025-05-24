///<summary>This module defines accessors for table cell properties.  
///Ideally, we need at least one accessor for every table cell property.
///</summary>
namespace OoxmlStyleEvaluator
module TableCellPropertyAccessors =

    open System.Xml.Linq
    open TableCellPropertiesEvaluator
    open OoxmlStyleEvaluator
        
    let resolveCellProperty (state: StyleEvaluatorState) (cell: XElement) (key: string) : string option =
        resolveEffectiveCellProperty
            cell
            state.TableStyleResolver
            state.GetPropertyByKey
            key
            
    let getCellProperty state run key defaultValue =
        resolveCellProperty state run key |> Option.defaultValue defaultValue

    let getCellPropertyBool state run key defaultValue =
        resolveCellProperty state run key |> Option.map System.Boolean.Parse |> Option.defaultValue  defaultValue

    let getCellPropertyInt state run key defaultValue =
        resolveCellProperty state run key |> Option.map int |> Option.defaultValue defaultValue
