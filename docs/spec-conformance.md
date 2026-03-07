# Spec Conformance Policy

This project follows the PresentationML part/relationship structure documented at:

- http://officeopenxml.com/prPresentation.php
- http://officeopenxml.com/anatomyofOOXML-pptx.php

## 1) Conformance Strategy

`pptx` roundtrip is handled in two modes:

1. `preserveSourceXml: true` (default)  
   Keep original slide XML and patch only model-managed nodes (`p:sp`, `p:cxnSp`, `p:pic`, `p:graphicFrame`).
2. `preserveSourceXml: false`  
   Rebuild slide XML from model only (debug/verification use).

Default mode is chosen to maximize compatibility with the full OOXML feature set, including
elements not yet fully materialized in the editable model.

## 2) What Is Preserved in Default Mode

- Slide-level namespaces and extension attributes.
- Unmanaged nodes in `p:spTree` (including unknown future extensions).
- GraphicFrame payloads for unsupported/partially supported objects.
- Existing relationships not touched by edited elements.

## 3) What Is Explicitly Materialized

- Shape/Text (`p:sp`)
- Connector/Line (`p:cxnSp`)
- Picture (`p:pic`)
- Table (`p:graphicFrame/a:tbl`)
- Custom geometry (`a:custGeom`) for shape paths (partial command coverage)

For these nodes, explicit style members (`fill`, `line`, text run styles, etc.) are written back from model values.

## 4) Practical Meaning of "Spec Coverage"

The OOXML specification is broad; this library uses:

- A strict part graph based on official PresentationML structure.
- Theme-resolved editable model for browser-side vector editing.
- Non-destructive writeback to avoid dropping unsupported schema branches.

When adding support for new XML branches, update:

- `docs/presentationml-reference.md`
- `docs/shape-support.md`
- this file (`docs/spec-conformance.md`)

Current notable gap:
- `a:custGeom/a:path/a:arcTo` uses approximate rendering (not yet mathematically identical to Office geometry engine).
- `a:custGeom` guide support covers common built-ins and iterative dependencies, but full Office geometry parity is still incomplete.
