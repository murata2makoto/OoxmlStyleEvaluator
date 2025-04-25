namespace OoxmlStyleEvaluator

open System
open System.IO.Compression
open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// <summary>
/// Provides evaluation utilities for WordprocessingML styles and numbering inside a DOCX/OOXML document.
/// </summary>
/// <param name="archive">The ZIP archive of the DOCX file.</param>
/// <param name="documentXml">The loaded XDocument of word/document.xml.</param>
type StyleEvaluator(archive: ZipArchive, documentXml: XDocument) =

    // Load styles.xml root element
    let stylesRoot : XElement =
        match archive.GetEntry("word/styles.xml") |> Option.ofObj with
        | Some entry -> XDocument.Load(entry.Open()).Root
        | None -> failwith "styles.xml not found in the DOCX archive."

    // Load numbering.xml root element (optional)
    let numberingRootOpt : XElement option =
        archive.GetEntry("word/numbering.xml")
        |> Option.ofObj
        |> Option.map (fun entry -> XDocument.Load(entry.Open()).Root)

    // Extract document default run properties from styles.xml
    let docDefaultsRPr : Map<string, XElement> =
        stylesRoot
        |> tryElement (w + "docDefaults")
        |> Option.bind (tryElement (w + "rPrDefault"))
        |> Option.bind (tryElement (w + "rPr"))
        |> Option.map (fun rPr ->
            rPr.Elements()
            |> Seq.map (fun e -> e.Name.ToString(), e)
            |> Map.ofSeq)
        |> Option.defaultValue Map.empty

    // Extract all paragraph (w:p) elements from document.xml
    let paras : XElement list =
        match documentXml.Root with
        | null -> []
        | root -> root.Descendants(w + "p") |> Seq.toList

    // Initialize counters for heading numbering (level 0 to 8)
    let mutable counters = Array.zeroCreate 9

    // Precompute heading number labels for all paragraphs
    let headingNumberLabels : Map<int, string> =
        paras
        |> List.indexed
        |> List.fold (fun acc (idx, para) ->
            let pPrOpt = tryElement (w + "pPr") para
            let pStyleIdOpt =
                pPrOpt
                |> Option.bind (tryElement (w + "pStyle"))
                |> Option.bind (tryAttrValue (w + "val"))

            match pStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "paragraph")
                |> Option.bind (fun style ->
                    tryElement (w + "pPr") style
                    |> Option.bind (tryElement (w + "outlineLvl"))
                    |> Option.bind (tryAttrValue (w + "val"))
                    |> Option.bind (fun v ->
                        match System.Int32.TryParse(v) with
                        | true, n -> Some n
                        | false, _ -> None))
                |> function
                    | Some level when level >= 0 && level <= 8 ->
                        counters.[level] <- counters.[level] + 1
                        for i in level + 1 .. 8 do
                            counters.[i] <- 0
                        let label =
                            [ for i in 0 .. level do
                                if counters.[i] > 0 then yield string counters.[i] ]
                            |> String.concat "."
                            |> fun s -> s + "."
                        Map.add idx label acc
                    | _ -> acc
            | None -> acc
        ) Map.empty

    // Dictionary to track bullet/numbered list counters per list (numId)
    let bulletCounters : System.Collections.Generic.Dictionary<string, int[]> =
        System.Collections.Generic.Dictionary()

    // Precompute bullet/numbered list labels for all paragraphs
    let bulletNumberLabels : Map<int, string option> =
        paras
        |> List.indexed
        |> List.fold (fun acc (idx, para) ->
            let pPrOpt = tryElement (w + "pPr") para
            let numPrOpt = pPrOpt |> Option.bind (tryElement (w + "numPr"))

            match numPrOpt with
            | Some numPr ->
                let numIdOpt =
                    tryElement (w + "numId") numPr
                    |> Option.bind (tryAttrValue (w + "val"))

                let ilvlOpt =
                    tryElement (w + "ilvl") numPr
                    |> Option.bind (tryAttrValue (w + "val"))

                match numIdOpt, ilvlOpt with
                | Some numId, Some ilvlStr ->
                    match System.Int32.TryParse(ilvlStr) with
                    | true, ilvl ->
                        let counters =
                            if bulletCounters.ContainsKey(numId) then
                                bulletCounters.[numId]
                            else
                                let arr = Array.zeroCreate 9
                                bulletCounters.Add(numId, arr)
                                arr
                        counters.[ilvl] <- counters.[ilvl] + 1
                        for i in ilvl + 1 .. 8 do
                            counters.[i] <- 0

                        let labelTemplateOpt =
                            numberingRootOpt
                            |> Option.bind (fun numberingRoot ->
                                numberingRoot.Elements(w + "num")
                                |> Seq.tryFind (fun num -> tryAttrValue (w + "numId") num = Some numId)
                                |> Option.bind (fun numElem ->
                                    tryElement (w + "abstractNumId") numElem
                                    |> Option.bind (tryAttrValue (w + "val"))
                                    |> Option.bind (fun abstractNumId ->
                                        numberingRoot.Elements(w + "abstractNum")
                                        |> Seq.tryFind (fun abs -> tryAttrValue (w + "abstractNumId") abs = Some abstractNumId)
                                        |> Option.bind (fun absElem ->
                                            absElem.Elements(w + "lvl")
                                            |> Seq.tryFind (fun lvl -> tryAttrValue (w + "ilvl") lvl = Some (string ilvl))
                                            |> Option.bind (fun lvlElem ->
                                                tryElement (w + "lvlText") lvlElem
                                                |> Option.bind (tryAttrValue (w + "val"))
                                            )
                                        )
                                    )
                                )
                            )

                        match labelTemplateOpt with
                        | Some template when not (System.String.IsNullOrWhiteSpace(template)) && template <> "?" ->
                            let mutable result = template
                            for i = 0 to ilvl do
                                result <- result.Replace("%" + string (i + 1), string counters.[i])
                            Map.add idx (Some result) acc
                        | _ ->
                            Map.add idx None acc
                    | false, _ -> acc
                | _ -> acc
            | None -> acc
        ) Map.empty

    /// <summary>
    /// Determines if the given paragraph is a heading paragraph based on its style definition.
    /// </summary>
    /// <param name="para">The paragraph (w:p element) to check.</param>
    /// <returns>True if it is a heading paragraph; otherwise, false.</returns>
    member _.IsHeadingParagraph(para: XElement) : bool =
        let pPrOpt = tryElement (w + "pPr") para
        let pStyleIdOpt =
            pPrOpt
            |> Option.bind (tryElement (w + "pStyle"))
            |> Option.bind (tryAttrValue (w + "val"))

        match pStyleIdOpt with
        | Some styleId ->
            stylesRoot.Elements(w + "style")
            |> Seq.tryFind (fun s ->
                tryAttrValue (w + "styleId") s = Some styleId &&
                tryAttrValue (w + "type") s = Some "paragraph")
            |> Option.bind (fun style ->
                tryElement (w + "pPr") style
                |> Option.bind (tryElement (w + "outlineLvl")))
            |> Option.isSome
        | None -> false

    /// <summary>
    /// Gets the heading level (outline level) of the given paragraph, if defined.
    /// </summary>
    /// <param name="para">The paragraph (w:p element).</param>
    /// <returns>The heading level (0-based) if available; otherwise, None.</returns>
    member _.GetHeadingLevel(para: XElement) : int option =
        let pPrOpt = tryElement (w + "pPr") para
        let pStyleIdOpt =
            pPrOpt
            |> Option.bind (tryElement (w + "pStyle"))
            |> Option.bind (tryAttrValue (w + "val"))

        match pStyleIdOpt with
        | Some styleId ->
            stylesRoot.Elements(w + "style")
            |> Seq.tryFind (fun s ->
                tryAttrValue (w + "styleId") s = Some styleId &&
                tryAttrValue (w + "type") s = Some "paragraph")
            |> Option.bind (fun style ->
                tryElement (w + "pPr") style
                |> Option.bind (tryElement (w + "outlineLvl"))
                |> Option.bind (tryAttrValue (w + "val"))
                |> Option.bind (fun v ->
                    match System.Int32.TryParse(v) with
                    | true, n -> Some n
                    | false, _ -> None))
        | None -> None

    /// <summary>
    /// Gets the heading number label (e.g., "1.2.3.") for the given paragraph, if any.
    /// </summary>
    /// <param name="para">The paragraph (w:p element).</param>
    /// <returns>The heading number label if available; otherwise, None.</returns>
    member _.GetHeadingNumberLabel(para: XElement) : string option =
        paras
        |> List.tryFindIndex (fun p -> System.Object.ReferenceEquals(p, para))
        |> Option.bind (fun idx -> headingNumberLabels |> Map.tryFind idx)

    /// <summary>
    /// Determines if the given paragraph is part of a bullet or numbered list.
    /// </summary>
    /// <param name="para">The paragraph (w:p element) to check.</param>
    /// <returns>True if it is a list item; otherwise, false.</returns>
    member _.IsBulletParagraph(para: XElement) : bool =
        tryElement (w + "pPr") para
        |> Option.bind (tryElement (w + "numPr"))
        |> Option.isSome

    /// <summary>
    /// Gets the bullet or numbering level (nesting depth) of the given paragraph, if any.
    /// </summary>
    /// <param name="para">The paragraph (w:p element).</param>
    /// <returns>The bullet/numbering level if available; otherwise, None.</returns>
    member _.GetBulletLevel(para: XElement) : int option =
        tryElement (w + "pPr") para
        |> Option.bind (tryElement (w + "numPr"))
        |> Option.bind (tryElement (w + "ilvl"))
        |> Option.bind (tryAttrValue (w + "val"))
        |> Option.bind (fun v ->
            match System.Int32.TryParse(v) with
            | true, n -> Some n
            | false, _ -> None)


    /// <summary>
    /// Gets the heading level (0-based) as Nullable&lt;int&gt; (for C# friendliness).
    /// </summary>
    /// <param name="para">The paragraph element.</param>
    /// <returns>Nullable heading level.</returns>
    member this.GetHeadingLevelNullable(para: XElement) : Nullable<int> =
        match this.GetHeadingLevel(para) with
        | Some v -> Nullable v
        | None -> Nullable()

    /// <summary>
    /// Gets the heading number label as string (null if none) for C# friendliness.
    /// </summary>
    /// <param name="para">The paragraph element.</param>
    /// <returns>Heading number label or null.</returns>
    member this.GetHeadingNumberLabelNullable(para: XElement) : string =
        match this.GetHeadingNumberLabel(para) with
        | Some s -> s
        | None -> null

    /// <summary>
    /// Gets the bullet level (nesting depth) as Nullable&lt;int&gt; (for C# friendliness).
    /// </summary>
    /// <param name="para">The paragraph element.</param>
    /// <returns>Nullable bullet level.</returns>
    member this.GetBulletLevelNullable(para: XElement) : Nullable<int> =
        match this.GetBulletLevel(para) with
        | Some v -> Nullable v
        | None -> Nullable()

    /// <summary>
    /// Gets the expanded bullet or numbering label for the given paragraph, if any.
    /// </summary>
    /// <param name="para">The paragraph (w:p element).</param>
    /// <returns>The expanded bullet or numbering label if available; otherwise, None.</returns>
    member this.GetBulletLabel(para: XElement) : string option =
        paras
        |> List.tryFindIndex (fun p -> System.Object.ReferenceEquals(p, para))
        |> Option.bind (fun idx ->
            bulletNumberLabels
            |> Map.tryFind idx
            |> Option.bind id
        )

        /// <summary>
        /// Gets the expanded bullet or numbering label for the given paragraph as nullable string (for C# friendliness).
        /// </summary>
        /// <param name="para">The paragraph (w:p element).</param>
        /// <returns>The expanded bullet or numbering label as string or null.</returns>
        member this.GetBulletLabelNullable(para: XElement) : string =
            match this.GetBulletLabel(para) with
            | Some label -> label
            | None -> null