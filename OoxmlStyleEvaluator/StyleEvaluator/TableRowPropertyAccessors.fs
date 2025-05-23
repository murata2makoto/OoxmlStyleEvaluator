///<summary>This module defines accessors for table row properties.  
///Ideally, we need at least one accessor for every table row property.
///</summary>
module OoxmlStyleEvaluator.TableRowPropertyAccessors

open System.Xml.Linq
open TableRowPropertiesEvaluator

type StyleEvaluator with
    member this.ResolveRowProperty(rowElement: XElement, key: string) : string option =
        resolveEffectiveRowProperty
            rowElement
            this.TableStyleResolver
            this.GetPropertyByKey
            key
