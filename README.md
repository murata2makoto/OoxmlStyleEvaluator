# OoxmlStyleEvaluator-
OoxmlStyleEvaluator is a lightweight and standards-compliant library for evaluating styles in WordprocessingML documents (.docx), as defined in the updated ISO/IEC 29500-1 standard.

This library resolves both paragraph and character style inheritance chains, and provides logic to determine the effective formatting of text runs—such as bold, italic, underline types, emphasis marks, font names, and colors—as interpreted by Microsoft Word.

It also distinguishes headings and lists based on paragraph styles, making it ideal for developers building EPUB converters, accessibility validators, CBT players, and other Word-compatible applications.

This program was developed by the **convenor of ISO/IEC JTC1/SC34/WG4**, the maintenance committee for OOXML, with the assistance of **ChatGPT**.

Although implemented in F#, the public API is designed for easy consumption from C# and other .NET languages.

Licensed under the MIT License.

### Features

- Fully conforms to the revised ISO/IEC 29500-1 model for style evaluation
- Handles both toggle and non-toggle properties
- Resolves `basedOn` chains for both paragraph and character styles
- Evaluates:
  - Bold, italic, strike-through
  - Underline type (e.g., single, double)
  - Emphasis mark (e.g., dot, comma)
  - Font and color
  - Heading and list status of paragraphs
- F# implementation with a C#-friendly API

