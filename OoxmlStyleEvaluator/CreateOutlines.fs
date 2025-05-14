module OoxmlStyleEvaluator.CreateOutlines

open System.Xml.Linq
open System.Collections.Generic
open OutlineAlgorithm.Interop.InteropFSharp
open OutlineAlgorithm.Interop
open OutlineAlgorithm

let generateOutlineIndices (tree: InteropTree<XElement>) =
    let dict = new Dictionary<XElement, int list>()
    let addToDictionary (label: XElement) (outlineIndex: int list): unit =
        dict.[label] <- outlineIndex
    traverseWithOutlineIndex tree addToDictionary
    dict