# PresentationML Reference Mapping

This project parses `.pptx` using a part-relationship graph aligned to the PresentationML structure.

Primary references:
- http://officeopenxml.com/prPresentation.php
- http://officeopenxml.com/anatomyofOOXML-pptx.php

## Parse Order

1. `ppt/presentation.xml`  
   Read `p:sldIdLst/p:sldId` and relationship ids.
2. `ppt/slides/slideN.xml`  
   Resolve each slide part from `presentation.xml.rels`.
3. `ppt/slideLayouts/slideLayoutN.xml`  
   Resolve via slide relationship type `.../slideLayout`.
4. `ppt/slideMasters/slideMasterN.xml`  
   Resolve via layout relationship type `.../slideMaster`.
5. `ppt/theme/themeN.xml`  
   Resolve via master relationship type `.../theme`.

Implemented in:
- `src/core/presentationGraph.js`
- `src/core/model.js`
- `docs/shape-support.md` (render support matrix)

## Theme-Resolved Model Policy

- Resolve theme color map (`clrMap`, `clrMapOvr`) before materializing element style.
- Keep element-level explicit style members (`fill`, `line`, text run style).
- Parse geometry metadata (`a:prstGeom` and `a:custGeom`) into editable `geometry` model.
- Preserve source xml fragments (`raw`) for round-trip safety.
- Preserve original slide xml (`_sourceXml`) and patch known nodes during writeback.

## Browser Editing Goal

- Model is editable in-memory on client side.
- `embedSlideIntoSlide()` supports vector-like slide embedding with transform.
- Renderers (`canvas` / `svg`) use model only (no screenshot paste workaround).

## Writeback Policy

- Default: non-destructive writeback (`preserveSourceXml: true`)
  - Keep unmanaged XML nodes and extensions.
  - Patch model-managed node families only.
- Optional: full rebuild (`preserveSourceXml: false`) for debug/verification.
