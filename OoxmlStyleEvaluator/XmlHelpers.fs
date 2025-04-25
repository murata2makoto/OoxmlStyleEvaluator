module OoxmlStyleEvaluator.XmlHelpers

open System.Xml.Linq

/// <summary>
/// The WordprocessingML namespace (http://schemas.openxmlformats.org/wordprocessingml/2006/main).
/// Used to construct fully qualified element and attribute names.
/// </summary>
let w = XNamespace.Get "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

/// <summary>
/// Attempts to retrieve a child element with the specified name from the given parent element.
/// Returns None if the element is missing.
/// </summary>
/// <param name="name">The qualified name of the child element to find</param>
/// <param name="parent">The parent XElement</param>
/// <returns>An option containing the child element, or None</returns>
let tryElement (name: XName) (parent: XElement) : XElement option =
    parent.Element(name) |> Option.ofObj

/// <summary>
/// Attempts to retrieve the string value of an attribute with the specified name from the given element.
/// Returns None if the attribute is missing.
/// </summary>
/// <param name="name">The qualified name of the attribute</param>
/// <param name="elem">The element to search</param>
/// <returns>An option containing the attribute value, or None</returns>
let tryAttrValue  (name: XName) (elem: XElement) : string option =
    elem.Attribute(name) |> Option.ofObj |> Option.map (fun a -> a.Value)
