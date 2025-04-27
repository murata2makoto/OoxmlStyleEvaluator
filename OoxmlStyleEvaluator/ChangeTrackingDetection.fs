module OoxmlStyleEvaluator.ChangeTrackingDetection

open XmlHelpers
open System.Xml.Linq

let containsChangeTrackingOrRelatedElements (documentRoot: XElement) : bool =
    let changeTrackingElements = 
        ["pPrChange";"rPrChange";
        "tcPrChange";
        "moveFrom";"moveTo";
        "tcPrChange";"trPrChange";"tblPrChange";"tblGridChange"]
        |> List.map (fun localName -> w + localName)
    changeTrackingElements
    |> List.exists (fun name -> 
        documentRoot.Descendants(name) |> Seq.isEmpty |> not)

