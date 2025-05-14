namespace  OoxmlStyleEvaluator

open System
open System.IO.Compression
open System.Xml.Linq
open OoxmlStyleEvaluator.XmlHelpers

/// <summary>
/// Provides evaluation utilities for WordprocessingML styles and numbering inside a DOCX/OOXML document.
/// </summary>
/// <param name="archive">The ZIP archive of the DOCX file.</param>
/// <param name="documentXml">The loaded XDocument of word/document.xml.</param>
type public StyleEvaluator(archive: ZipArchive, documentXml: XDocument) =

    // Load styles.xml root element
    let stylesRoot : XElement =
        match archive.GetEntry("word/styles.xml") |> Option.ofObj with
        | Some entry ->
            match XDocument.Load(entry.Open()).Root |> Option.ofObj with
            | Some root -> root
            | None -> failwith "word/styles.xml does not have a root element."
        | None ->
            failwith "styles.xml not found in the DOCX archive."

    // Load numbering.xml root element (optional)
    let numberingRootOpt : XElement option =
        archive.GetEntry("word/numbering.xml")
        |> Option.ofObj
        |> Option.bind (fun entry ->
            XDocument.Load(entry.Open()).Root
            |> Option.ofObj
        )

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

    member _.Archive = archive
    member _.DocumentXml = documentXml

    /// <summary>
    /// Determines if the given paragraph is a heading paragraph based on its style definition.
    /// </summary>
    /// <param name="para">The paragraph (w:p element) to check.</param>
    /// <returns>True if it is a heading paragraph; otherwise, false.</returns>
    member public  _.IsHeadingParagraph(para: XElement) : bool =
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
    /// <returns>The heading level (0-based) if available; otherwise, -1.</returns>
    member public  _.GetHeadingLevel(para: XElement) : int  =
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
                |> Option.defaultValue -1
        | None -> -1

    /// <summary>
    /// Gets the heading number label (e.g., "1.2.3.") for the given paragraph, if any.
    /// </summary>
    /// <param name="para">The paragraph (w:p element).</param>
    /// <returns>The heading number label if available; otherwise, the empty string.</returns>
    member public  _.GetHeadingNumberLabel(para: XElement) : string  =
        paras
        |> List.tryFindIndex (fun p -> System.Object.ReferenceEquals(p, para))
        |> Option.bind (fun idx -> headingNumberLabels |> Map.tryFind idx)
        |> Option.defaultValue ""

    /// <summary>
    /// Determines if the given paragraph is part of a bullet or numbered list.
    /// </summary>
    /// <param name="para">The paragraph (w:p element) to check.</param>
    /// <returns>True if it is a list item; otherwise, false.</returns>
    member public  _.IsBulletParagraph(para: XElement) : bool =
        tryElement (w + "pPr") para
        |> Option.bind (tryElement (w + "numPr"))
        |> Option.isSome

    /// <summary>
    /// Gets the bullet or numbering level (nesting depth) of the given paragraph, if any.
    /// </summary>
    /// <param name="para">The paragraph (w:p element).</param>
    /// <returns>The bullet/numbering level if available; otherwise, -1.</returns>
    member public  _.GetBulletLevel(para: XElement): int  =
        tryElement (w + "pPr") para
        |> Option.bind (tryElement (w + "numPr"))
        |> Option.bind (tryElement (w + "ilvl"))
        |> Option.bind (tryAttrValue (w + "val"))
        |> Option.map (fun v ->
            match System.Int32.TryParse(v) with
            | true, n -> n
            | false, _ -> -1)
        |> Option.defaultValue -1


    /// <summary>
    /// Gets the expanded bullet or numbering label for the given paragraph, if any.
    /// </summary>
    /// <param name="para">The paragraph (w:p element).</param>
    /// <returns>The expanded bullet or numbering label if available; otherwise, the empty string.</returns>
    member public  this.GetBulletLabel(para: XElement) : string  =
        paras
        |> List.tryFindIndex (fun p -> System.Object.ReferenceEquals(p, para))
        |> Option.bind (fun idx -> bulletNumberLabels |> Map.tryFind idx |> Option.bind id)
        |> Option.defaultValue ""

    /// <summary>
    /// Evaluates the effective value of a run property (e.g., bold, italic) for a given run,
    /// based on direct formatting, character style, paragraph style, table style (if applicable), and document defaults.
    /// </summary>
    /// <param name="run">The run (w:r element) to evaluate.</param>
    /// <param name="propName">The local name of the run property to check (e.g., "b" for bold).</param>
    /// <returns>True if the property is effectively enabled; otherwise, false.</returns>
    member _.IsRunPropertyEnabled(run: XElement, propName: string) : bool =
        let tryGetDirectFormatting (run: XElement) =
            tryElement (w + "rPr") run
            |> Option.bind (tryElement (w + propName))

        let tryGetFromCharacterStyle (run: XElement) =
            let rStyleIdOpt =
                tryElement (w + "rPr") run
                |> Option.bind (tryElement (w + "rStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match rStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "character")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind (tryElement (w + propName))
            | None -> None

        let tryGetFromParagraphStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let pStyleIdOpt =
                paraOpt
                |> Option.bind (tryElement (w + "pPr"))
                |> Option.bind (tryElement (w + "pStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match pStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "paragraph")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind (tryElement (w + propName))
            | None -> None

        let tryGetFromTableStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let tblOpt = paraOpt |> Option.bind (fun p -> p.Ancestors(w + "tbl") |> Seq.tryHead)
            match tblOpt with
            | Some tbl ->
                let tblPrOpt = tryElement (w + "tblPr") tbl
                let tblStyleIdOpt =
                    tblPrOpt
                    |> Option.bind (tryElement (w + "tblStyle"))
                    |> Option.bind (tryAttrValue (w + "val"))
                match tblStyleIdOpt with
                | Some styleId ->
                    stylesRoot.Elements(w + "style")
                    |> Seq.tryFind (fun s ->
                        tryAttrValue (w + "styleId") s = Some styleId &&
                        tryAttrValue (w + "type") s = Some "table")
                    |> Option.bind (tryElement (w + "rPr"))
                    |> Option.bind (tryElement (w + propName))
                | None -> None
            | None -> None

        let tryGetFromDocDefaults () =
            docDefaultsRPr
            |> Map.tryFind (w + propName |> string)

        // Evaluation priority
        tryGetDirectFormatting run
        |> Option.orElseWith (fun () -> tryGetFromCharacterStyle run)
        |> Option.orElseWith (fun () -> tryGetFromParagraphStyle run)
        |> Option.orElseWith (fun () -> tryGetFromTableStyle run)
        |> Option.orElseWith (fun () -> tryGetFromDocDefaults ())
        |> Option.map (fun _ -> true)
        |> Option.defaultValue false

    /// <summary>
    /// Checks if the given run is bold.
    /// </summary>
    /// <param name="run">The run (w:r element) to check.</param>
    /// <returns>True if the run is bold; otherwise, false.</returns>
    member this.IsRunBold(run: XElement) : bool =
        this.IsRunPropertyEnabled(run, "b")

    /// <summary>
    /// Checks if the given run is italicized.
    /// </summary>
    /// <param name="run">The run (w:r element) to check.</param>
    /// <returns>True if the run is italicized; otherwise, false.</returns>
    member public  this.IsRunItalic(run: XElement) : bool =
        this.IsRunPropertyEnabled(run, "i")

    /// <summary>
    /// Checks if the given run is struck through.
    /// </summary>
    /// <param name="run">The run (w:r element) to check.</param>
    /// <returns>True if the run is struck through; otherwise, false.</returns>
    member public this.IsRunStrike(run: XElement) : bool =
        this.IsRunPropertyEnabled(run, "strike")

    /// <summary>
    /// Checks if the given run uses capital letters (small caps or all caps).
    /// </summary>
    /// <param name="run">The run (w:r element) to check.</param>
    /// <returns>True if the run uses capital letters; otherwise, false.</returns>
    member public this.IsRunCaps(run: XElement) : bool =
        this.IsRunPropertyEnabled(run, "caps")

    /// <summary>
    /// Gets the effective underline type (e.g., "single", "double") for the given run,
    /// based on direct formatting, character style, paragraph style, table style, and document defaults.
    /// </summary>
    /// <param name="run">The run (w:r element) to evaluate.</param>
    /// <returns>
    /// The underline type string if set (e.g., "single"), or an empty string if no underline is specified.
    /// </returns>
    member public  _.GetEffectiveUnderlineType(run: XElement) : string =
        let tryGetUnderlineValue (rPr: XElement) =
            tryElement (w + "u") rPr
            |> Option.bind (fun uElem -> tryAttrValue (w + "val") uElem)
    
        let tryFromDirectFormatting (run: XElement) =
            tryElement (w + "rPr") run
            |> Option.bind tryGetUnderlineValue

        let tryFromCharacterStyle (run: XElement) =
            let rStyleIdOpt =
                tryElement (w + "rPr") run
                |> Option.bind (tryElement (w + "rStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match rStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "character")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetUnderlineValue
            | None -> None

        let tryFromParagraphStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let pStyleIdOpt =
                paraOpt
                |> Option.bind (tryElement (w + "pPr"))
                |> Option.bind (tryElement (w + "pStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match pStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "paragraph")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetUnderlineValue
            | None -> None

        let tryFromTableStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let tblOpt = paraOpt |> Option.bind (fun p -> p.Ancestors(w + "tbl") |> Seq.tryHead)
            match tblOpt with
            | Some tbl ->
                let tblPrOpt = tryElement (w + "tblPr") tbl
                let tblStyleIdOpt =
                    tblPrOpt
                    |> Option.bind (tryElement (w + "tblStyle"))
                    |> Option.bind (tryAttrValue (w + "val"))
                match tblStyleIdOpt with
                | Some styleId ->
                    stylesRoot.Elements(w + "style")
                    |> Seq.tryFind (fun s ->
                        tryAttrValue (w + "styleId") s = Some styleId &&
                        tryAttrValue (w + "type") s = Some "table")
                    |> Option.bind (tryElement (w + "rPr"))
                    |> Option.bind tryGetUnderlineValue
                | None -> None
            | None -> None

        let tryFromDocDefaults () =
            docDefaultsRPr
            |> Map.tryFind (w + "u" |> string)
            |> Option.bind (fun uElem -> tryAttrValue (w + "val") uElem)

        // Priority order evaluation
        tryFromDirectFormatting run
        |> Option.orElseWith (fun () -> tryFromCharacterStyle run)
        |> Option.orElseWith (fun () -> tryFromParagraphStyle run)
        |> Option.orElseWith (fun () -> tryFromTableStyle run)
        |> Option.orElseWith (tryFromDocDefaults)
        |> Option.defaultValue ""

    /// <summary>
    /// Gets the effective text color for the given run,
    /// based on direct formatting, character style, paragraph style, table style, and document defaults.
    /// </summary>
    /// <param name="run">The run (w:r element) to evaluate.</param>
    /// <returns>
    /// The color value (hex RGB string, e.g., "FF0000") if specified; otherwise, an empty string.
    /// </returns>
    member public _.GetEffectiveColor(run: XElement) : string =
        let tryGetColorValue (rPr: XElement) =
            tryElement (w + "color") rPr
            |> Option.bind (fun colorElem -> tryAttrValue (w + "val") colorElem)

        let tryFromDirectFormatting (run: XElement) =
            tryElement (w + "rPr") run
            |> Option.bind tryGetColorValue

        let tryFromCharacterStyle (run: XElement) =
            let rStyleIdOpt =
                tryElement (w + "rPr") run
                |> Option.bind (tryElement (w + "rStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match rStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "character")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetColorValue
            | None -> None

        let tryFromParagraphStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let pStyleIdOpt =
                paraOpt
                |> Option.bind (tryElement (w + "pPr"))
                |> Option.bind (tryElement (w + "pStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match pStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "paragraph")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetColorValue
            | None -> None

        let tryFromTableStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let tblOpt = paraOpt |> Option.bind (fun p -> p.Ancestors(w + "tbl") |> Seq.tryHead)
            match tblOpt with
            | Some tbl ->
                let tblPrOpt = tryElement (w + "tblPr") tbl
                let tblStyleIdOpt =
                    tblPrOpt
                    |> Option.bind (tryElement (w + "tblStyle"))
                    |> Option.bind (tryAttrValue (w + "val"))
                match tblStyleIdOpt with
                | Some styleId ->
                    stylesRoot.Elements(w + "style")
                    |> Seq.tryFind (fun s ->
                        tryAttrValue (w + "styleId") s = Some styleId &&
                        tryAttrValue (w + "type") s = Some "table")
                    |> Option.bind (tryElement (w + "rPr"))
                    |> Option.bind tryGetColorValue
                | None -> None
            | None -> None

        let tryFromDocDefaults () =
            docDefaultsRPr
            |> Map.tryFind (w + "color" |> string)
            |> Option.bind (fun colorElem -> tryAttrValue (w + "val") colorElem)

        // Priority order evaluation
        tryFromDirectFormatting run
        |> Option.orElseWith (fun () -> tryFromCharacterStyle run)
        |> Option.orElseWith (fun () -> tryFromParagraphStyle run)
        |> Option.orElseWith (fun () -> tryFromTableStyle run)
        |> Option.orElseWith (tryFromDocDefaults)
        |> Option.defaultValue ""

    /// <summary>
    /// Gets the effective font name for the given run, for a specific script range (ascii, eastAsia, hAnsi, or cs),
    /// based on direct formatting, character style, paragraph style, table style, and document defaults.
    /// </summary>
    /// <param name="run">The run (w:r element) to evaluate.</param>
    /// <param name="script">The script attribute to target ("ascii", "eastAsia", "hAnsi", or "cs").</param>
    /// <returns>The font name if specified; otherwise, an empty string.</returns>
    member public _.GetEffectiveFont(run: XElement, script: string) : string =
        let tryGetFontValue (rPr: XElement) =
            tryElement (w + "rFonts") rPr
            |> Option.bind (fun rFontsElem ->
                tryAttrValue (w + script) rFontsElem
            )

        let tryFromDirectFormatting (run: XElement) =
            tryElement (w + "rPr") run
            |> Option.bind tryGetFontValue

        let tryFromCharacterStyle (run: XElement) =
            let rStyleIdOpt =
                tryElement (w + "rPr") run
                |> Option.bind (tryElement (w + "rStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match rStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "character")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetFontValue
            | None -> None

        let tryFromParagraphStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let pStyleIdOpt =
                paraOpt
                |> Option.bind (tryElement (w + "pPr"))
                |> Option.bind (tryElement (w + "pStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match pStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "paragraph")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetFontValue
            | None -> None

        let tryFromTableStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let tblOpt = paraOpt |> Option.bind (fun p -> p.Ancestors(w + "tbl") |> Seq.tryHead)
            match tblOpt with
            | Some tbl ->
                let tblPrOpt = tryElement (w + "tblPr") tbl
                let tblStyleIdOpt =
                    tblPrOpt
                    |> Option.bind (tryElement (w + "tblStyle"))
                    |> Option.bind (tryAttrValue (w + "val"))
                match tblStyleIdOpt with
                | Some styleId ->
                    stylesRoot.Elements(w + "style")
                    |> Seq.tryFind (fun s ->
                        tryAttrValue (w + "styleId") s = Some styleId &&
                        tryAttrValue (w + "type") s = Some "table")
                    |> Option.bind (tryElement (w + "rPr"))
                    |> Option.bind tryGetFontValue
                | None -> None
            | None -> None

        let tryFromDocDefaults () =
            docDefaultsRPr
            |> Map.tryFind (w + "rFonts" |> string)
            |> Option.bind (fun rFontsElem -> tryAttrValue (w + script) rFontsElem)

        // Priority order evaluation
        tryFromDirectFormatting run
        |> Option.orElseWith (fun () -> tryFromCharacterStyle run)
        |> Option.orElseWith (fun () -> tryFromParagraphStyle run)
        |> Option.orElseWith (fun () -> tryFromTableStyle run)
        |> Option.orElseWith (tryFromDocDefaults)
        |> Option.defaultValue ""


    /// <summary>
    /// Helper to retrieve a font-related attribute (e.g., ascii, asciiTheme, hAnsi) from rFonts.
    /// </summary>
    /// <param name="run">The run (w:r element) to evaluate.</param>
    /// <param name="attrName">The attribute name to look for.</param>
    /// <returns>The attribute value as string, or empty string if not set.</returns>
    member private this.GetEffectiveFontAttribute(run: XElement, attrName: string) : string =
        let tryGetFontAttr (rPr: XElement) =
            tryElement (w + "rFonts") rPr
            |> Option.bind (fun rFontsElem -> tryAttrValue (w + attrName) rFontsElem)

        let tryFromDirectFormatting (run: XElement) =
            tryElement (w + "rPr") run
            |> Option.bind tryGetFontAttr

        let tryFromCharacterStyle (run: XElement) =
            let rStyleIdOpt =
                tryElement (w + "rPr") run
                |> Option.bind (tryElement (w + "rStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match rStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "character")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetFontAttr
            | None -> None

        let tryFromParagraphStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let pStyleIdOpt =
                paraOpt
                |> Option.bind (tryElement (w + "pPr"))
                |> Option.bind (tryElement (w + "pStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match pStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "paragraph")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetFontAttr
            | None -> None

        let tryFromTableStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let tblOpt = paraOpt |> Option.bind (fun p -> p.Ancestors(w + "tbl") |> Seq.tryHead)
            match tblOpt with
            | Some tbl ->
                let tblPrOpt = tryElement (w + "tblPr") tbl
                let tblStyleIdOpt =
                    tblPrOpt
                    |> Option.bind (tryElement (w + "tblStyle"))
                    |> Option.bind (tryAttrValue (w + "val"))
                match tblStyleIdOpt with
                | Some styleId ->
                    stylesRoot.Elements(w + "style")
                    |> Seq.tryFind (fun s ->
                        tryAttrValue (w + "styleId") s = Some styleId &&
                        tryAttrValue (w + "type") s = Some "table")
                    |> Option.bind (tryElement (w + "rPr"))
                    |> Option.bind tryGetFontAttr
                | None -> None
            | None -> None

        let tryFromDocDefaults () =
            docDefaultsRPr
            |> Map.tryFind (w + "rFonts" |> string)
            |> Option.bind (fun rFontsElem -> tryAttrValue (w + attrName) rFontsElem)

        tryFromDirectFormatting run
        |> Option.orElseWith (fun () -> tryFromCharacterStyle run)
        |> Option.orElseWith (fun () -> tryFromParagraphStyle run)
        |> Option.orElseWith (fun () -> tryFromTableStyle run)
        |> Option.orElseWith (tryFromDocDefaults)
        |> Option.defaultValue ""

    /// <summary>
    /// Gets the effective ASCII font name for the given run.
    /// </summary>
    member public this.GetEffectiveAsciiFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "ascii")

    /// <summary>
    /// Gets the effective ASCII theme font name for the given run.
    /// </summary>
    member public this.GetEffectiveAsciiThemeFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "asciiTheme")

    /// <summary>
    /// Gets the effective High ANSI font name for the given run.
    /// </summary>
    member public this.GetEffectiveHAnsiFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "hAnsi")

    /// <summary>
    /// Gets the effective High ANSI theme font name for the given run.
    /// </summary>
    member public this.GetEffectiveHAnsiThemeFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "hAnsiTheme")

    /// <summary>
    /// Gets the effective East Asia font name for the given run.
    /// </summary>
    member public this.GetEffectiveEastAsiaFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "eastAsia")

    /// <summary>
    /// Gets the effective East Asia theme font name for the given run.
    /// </summary>
    member public this.GetEffectiveEastAsiaThemeFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "eastAsiaTheme")

    /// <summary>
    /// Gets the effective Complex Script font name for the given run.
    /// </summary>
    member public this.GetEffectiveCsFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "cs")

    /// <summary>
    /// Gets the effective Complex Script theme font name for the given run.
    /// </summary>
    member public this.GetEffectiveCsThemeFont(run: XElement) : string =
        this.GetEffectiveFontAttribute(run, "cstheme")

    /// <summary>
    /// Gets the effective emphasis mark (e.g., "dot", "comma") for the given run,
    /// based on direct formatting, character style, paragraph style, table style, and document defaults.
    /// </summary>
    /// <param name="run">The run (w:r element) to evaluate.</param>
    /// <returns>The emphasis mark value if specified; otherwise, an empty string.</returns>
    member public _.GetEffectiveEmphasisMark(run: XElement) : string =
        let tryGetEmphasisValue (rPr: XElement) =
            tryElement (w + "em") rPr
            |> Option.bind (fun emElem -> tryAttrValue (w + "val") emElem)

        let tryFromDirectFormatting (run: XElement) =
            tryElement (w + "rPr") run
            |> Option.bind tryGetEmphasisValue

        let tryFromCharacterStyle (run: XElement) =
            let rStyleIdOpt =
                tryElement (w + "rPr") run
                |> Option.bind (tryElement (w + "rStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match rStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "character")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetEmphasisValue
            | None -> None

        let tryFromParagraphStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let pStyleIdOpt =
                paraOpt
                |> Option.bind (tryElement (w + "pPr"))
                |> Option.bind (tryElement (w + "pStyle"))
                |> Option.bind (tryAttrValue (w + "val"))
            match pStyleIdOpt with
            | Some styleId ->
                stylesRoot.Elements(w + "style")
                |> Seq.tryFind (fun s ->
                    tryAttrValue (w + "styleId") s = Some styleId &&
                    tryAttrValue (w + "type") s = Some "paragraph")
                |> Option.bind (tryElement (w + "rPr"))
                |> Option.bind tryGetEmphasisValue
            | None -> None

        let tryFromTableStyle (run: XElement) =
            let paraOpt = run.Ancestors(w + "p") |> Seq.tryHead
            let tblOpt = paraOpt |> Option.bind (fun p -> p.Ancestors(w + "tbl") |> Seq.tryHead)
            match tblOpt with
            | Some tbl ->
                let tblPrOpt = tryElement (w + "tblPr") tbl
                let tblStyleIdOpt =
                    tblPrOpt
                    |> Option.bind (tryElement (w + "tblStyle"))
                    |> Option.bind (tryAttrValue (w + "val"))
                match tblStyleIdOpt with
                | Some styleId ->
                    stylesRoot.Elements(w + "style")
                    |> Seq.tryFind (fun s ->
                        tryAttrValue (w + "styleId") s = Some styleId &&
                        tryAttrValue (w + "type") s = Some "table")
                    |> Option.bind (tryElement (w + "rPr"))
                    |> Option.bind tryGetEmphasisValue
                | None -> None
            | None -> None

        let tryFromDocDefaults () =
            docDefaultsRPr
            |> Map.tryFind (w + "em" |> string)
            |> Option.bind (fun emElem -> tryAttrValue (w + "val") emElem)

        // Priority order evaluation
        tryFromDirectFormatting run
        |> Option.orElseWith (fun () -> tryFromCharacterStyle run)
        |> Option.orElseWith (fun () -> tryFromParagraphStyle run)
        |> Option.orElseWith (fun () -> tryFromTableStyle run)
        |> Option.orElseWith (tryFromDocDefaults)
        |> Option.defaultValue ""
