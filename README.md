# OoxmlStyleEvaluator

**OoxmlStyleEvaluator** is a .NET library for accurate evaluation of styles (paragraph, character, table) in OOXML WordprocessingML documents.

It is usable from both **F# and C#** programs.

This library was developed under the technical supervision of the convenor of **ISO/IEC JTC 1/SC 34/WG 4**,  
the committee responsible for maintaining the OOXML standard.

## Overview

- Minimalistic and clean API: everything is based on `System.Xml.Linq` (`XElement`, `XDocument`).
- Full XPath compatibility: you can freely apply XPath queries on the loaded documents.
- Focused functionality: dedicated to the **complex style inheritance and style hierarchy** of OOXML.
- Faithful to the latest ISO/IEC standard drafts.

## Standards Compliance

OoxmlStyleEvaluator implements style evaluation strictly according to:

- **ISO/IEC CD 29500-1:2025** (Fundamentals and Markup Language Reference)
- **ISO/IEC CD 29500-4:2025** (Transitional Migration Features)

In particular, OoxmlStyleEvaluator faithfully implements the style resolution model described in  
**ISO/IEC CD 29500-1:2025**, covering:

- §17.7.1 Style inheritance
- §17.7.2 Style hierarchy
- §17.7.3 Toggle properties

Recent clarifications and improvements from the drafts are fully incorporated.

### Warning 

As of 2025 May, OoxmlStyleEvaluator is still non-conformant.  It is expected to be conformant when **ISO/IEC CD 29500-1:2025** and **ISO/IEC CD 29500-4:2025** are published as International Standards.

## Key Features

- **System.Xml.Linq-based**  
  No proprietary object models. Standard .NET types (`XElement`, `XDocument`) are used.

- **Full XPath Access**  
  Since it uses standard LINQ to XML types, XPath queries can be freely applied.

- **Precise Style Resolution**  
  Handles direct formatting, paragraph styles, character styles, table styles, and document defaults.

- **Internationalized Font Handling**  
  Supports retrieval of fonts for different code ranges:
  - ASCII (`w:ascii`)
  - East Asian (`w:eastAsia`)
  - Complex Scripts (`w:cs`)
  - Theme-based font mappings

- **Easy Integration with F# and C#**  
  Designed for smooth usage from both languages.

## Usage Examples

Example usage projects are included separately:

- `Example.FSharpConsole`
- `Example.CSharpConsole`

These examples demonstrate how to:

- Open a DOCX file
- Initialize `StyleEvaluator`
- Retrieve heading levels and bullet labels
- Check bold/italic/underline status of runs
- Handle multilingual font settings

## Summary

OoxmlStyleEvaluator is for developers who need:

- Trustworthy, spec-compliant evaluation of OOXML styles
- Simple integration with standard .NET types
- A small, focused, reliable tool without the overhead of large object models

If you want **accurate OOXML style resolution** easily usable from both F# and C#,  
**OoxmlStyleEvaluator** is the ideal solution.

---

# License

This project is licensed under the MIT License. See the [LICENSE](./LICENSE) file for details.
