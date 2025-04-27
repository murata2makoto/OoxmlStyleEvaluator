module OoxmlStyleEvaluator.FindNearestPreceding

open System.Xml.Linq
  
/// <summary>
/// Traverses backwards from the given node to find the nearest preceding element that satisfies the specified condition.
/// The search includes the current node, its preceding siblings, and its ancestors in reverse order.
/// </summary>
/// <param name="node">The starting node (`XNode`) from which the search begins.</param>
/// <param name="predicate">A function that defines the condition the target element must satisfy.</param>
/// <returns>
/// The nearest preceding element (`XElement`) that satisfies the specified condition, 
/// or None if no such element is found.
/// </returns>
let findNearestPreceding
    (node: XNode)
    (predicate: XElement -> bool)
    : XElement option =

    // Start from the current node and traverse backwards
    let rec traverseBackwards (node: XNode) =
        if (node :? XElement) 
            && (let n = node :?> XElement in predicate n) then 
            // If the node is an element satisfying the given predicate, return it
              Some(node :?> XElement)
        else
            // Otherwise, move to the previous or parent node
            match node.PreviousNode |> Option.ofObj, node.Parent |> Option.ofObj with
            | Some(prev), _ -> traverseBackwards prev
            | None, Some(parent) -> traverseBackwards parent
            | None, None -> None // Stop if there are no more previous nodes and no ancestor nodes.

    // Start traversal from the given element

    traverseBackwards node



