///<summary>This module defines accessors for table row properties.  
///Ideally, we need at least one accessor for every table row property.
///</summary>
namespace OoxmlStyleEvaluator
module TableRowPropertyAccessors =

    open System.Xml.Linq
    open TableRowPropertiesEvaluator
    open OoxmlStyleEvaluator
        
    let resolveRowProperty (state: StyleEvaluatorState) (row: XElement) (key: string) : string option =
        resolveEffectiveRowProperty
            row
            state.TableStyleResolver
            state.GetPropertyByKey
            key
