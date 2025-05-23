module OoxmlStyleEvaluator.TableRowPropertiesEvaluator

open System.Xml.Linq
open XmlHelpers
open PropertyTypes


/// <summary>
/// Resolves the effective value of a single table row property (from <c>&lt;trPr&gt;</c>)
/// for a given table row, using both style-based and directly specified properties.
/// </summary>
/// <param name="rowElement">The table row (<c>&lt;w:tr&gt;</c>) element to process.</param>
/// <param name="tableStyleResolver">A function to resolve a styleId into table style properties.</param>
/// <param name="getPropertyByKey">A memoized function to get a property from an element by key.</param>
/// <param name="key">The property key to look up (e.g., <c>"tblHeader/@val"</c>).</param>
/// <returns>The effective value as <c>Some string</c>, or <c>None</c> if not found.</returns>
let resolveEffectiveRowProperty
    (rowElement: XElement)
    (tableStyleResolver: string -> TableStyleProperties)
    (getPropertyByKey: XElement -> string -> string option)
    (key: string)
    : string option =

    // Step 1: Try direct value from <trPr>
    let directValue =
        rowElement
        |> tryElement (w + "trPr")
        |> Option.bind (fun trPr -> getPropertyByKey trPr key)

    match directValue with
    | Some v -> Some v
    | None ->
        // Step 2: Try value from table style
        let fromTableStyle =
            rowElement.Ancestors(w + "tbl") |> Seq.tryHead
            |> Option.bind (tryElement (w + "tblPr"))
            |> Option.bind (tryElement (w + "tblStyle"))
            |> Option.bind (tryAttrValue (w + "val"))
            |> Option.bind (fun styleId ->
                let styleChain = tableStyleResolver styleId
                let top = Map.tryFind key styleChain.TopLevel.TrPr
                let whole =
                    styleChain.ByType
                    |> Map.tryFind "wholeTable"
                    |> Option.bind (fun group -> Map.tryFind key group.TrPr)
                top |> Option.orElse whole)

        fromTableStyle

