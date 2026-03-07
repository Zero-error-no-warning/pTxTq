# PptxThemeDocument

Source-focused repository for working with editable PresentationML models extracted from `.pptx`.

Japanese version: [README.ja.md](./README.ja.md)

## Scope

This repository is intended to publish the implementation under `src/`.

The following are treated as local or verification-only assets and should stay out of the public history unless there is a specific reason to publish them:

- `scripts/`
- `docs/`
- sample `.pptx` files
- local history or workspace metadata

## What `src/` contains

- `src/core/`: package loading, theme resolution, PresentationML graph traversal, editable model
- `src/render/`: SVG and Canvas rendering from the in-memory model
- `src/write/`: round-trip writeback to `.pptx`
- `src/utils/`: XML, units, color, geometry, and object helpers

## Current capabilities

- Load a `.pptx` package into an editable model
- Resolve theme and slide relationships
- Render slides to SVG
- Render slides to Canvas
- Embed one slide into another as editable elements
- Write the edited model back to `.pptx`

## Public API entry point

`src/index.js`

```js
import { PptxThemeDocument } from "./src/index.js";

const doc = await PptxThemeDocument.load("input.pptx");
const svg = await doc.renderSlide(0, { mode: "svg" });
await doc.saveAs("roundtrip.pptx");
```

## Notes

- The source depends on `jszip` and `fast-xml-parser`.
- Verification helpers may rely on additional Node, Python, or PowerShell tooling, but those are outside the intended public scope of this repository.
