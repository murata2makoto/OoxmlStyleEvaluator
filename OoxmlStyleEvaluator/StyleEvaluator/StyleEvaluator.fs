namespace  OoxmlStyleEvaluator

open System.IO.Compression
open System.Xml.Linq
open DocumentDefaults
open CharacterStyleInheritance
open ParagraphStyleInheritance
open TableStyleInheritance
open LoadFromArchive
open PropertyExtractor

/// <summary>
/// Provides evaluation utilities for WordprocessingML styles and numbering inside a DOCX/OOXML document.
/// </summary>
/// <param name="archive">The ZIP archive of the DOCX file.</param>
/// <param name="documentXml">The loaded XDocument of word/document.xml.</param>
type public StyleEvaluator(archive: ZipArchive) =

    let documentRoot = loadDocumentXml archive
    let stylesRoot : XElement = loadStyles archive

    member public this.DocumentRoot = documentRoot
    member internal _.StylesRoot = stylesRoot
    member internal  _.DocDefaultsRPr  = getRPrDocDefaults stylesRoot
    member internal  _.DocDefaultsPPr  = getPPrDocDefaults stylesRoot
    member internal  _.CharacterStyleResolver = createCharacterStyleResolver stylesRoot
    member internal  _.ParagraphStyleResolver = createParagraphStyleResolver stylesRoot
    member internal  _.TableStyleResolver = createTableStyleResolver stylesRoot
    member internal  _.GetPropertyByKey = makeGetPropertyByKey()
    