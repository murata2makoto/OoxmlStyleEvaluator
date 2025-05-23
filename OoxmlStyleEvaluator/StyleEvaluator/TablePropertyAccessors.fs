///<summary>This module defines accessors for table properties.  
///Ideally, we need at least one accessor for every table property.
///</summary>
module OoxmlStyleEvaluator.TablePropertyAccessors

open System.Xml.Linq
open TablePropertiesEvaluator

type StyleEvaluator with
    member this.ResolveTableProperty(tableElement: XElement, key: string) : string option =
        resolveEffectiveTableProperty
            tableElement
            this.TableStyleResolver
            this.GetPropertyByKey
            key

