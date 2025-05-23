///<summary>This module defines accessors for table cell properties.  
///Ideally, we need at least one accessor for every table cell property.
///</summary>
module OoxmlStyleEvaluator.TableCellPropertyAccessors

open System.Xml.Linq
open TableCellPropertiesEvaluator

type StyleEvaluator with
    member this.ResolveCellProperty(cellElement: XElement, key: string) : string option =
        resolveEffectiveCellProperty
            cellElement
            this.TableStyleResolver
            this.GetPropertyByKey
            key
