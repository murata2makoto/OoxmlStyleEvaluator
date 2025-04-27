module OoxmlStyleEvaluator.CreateOutlines

open System.Xml.Linq
open System.Collections.Generic
open OutlineAlgorithm.OutlineIndex
open OutlineAlgorithm.TreeAndHedge

let generateOutlineIndices (hedge: Hedge<XElement>) =
    let dict = new Dictionary<XElement, int list>()
    let addToDictionary (label: XElement) (outlineIndex: int list): unit =
        dict.[label] <- outlineIndex
    visitHedge hedge [] addToDictionary
    dict