# Shape Support Matrix

This document tracks which PowerPoint XML shape kinds are currently rendered to SVG/Canvas.

References:
- http://officeopenxml.com/prPresentation.php
- http://officeopenxml.com/anatomyofOOXML-pptx.php
- Drawing shape model in PresentationML (`p:sp`, `a:prstGeom/@prst`)

## 1) Object-Level Support

| XML source | Model type | Canvas | SVG | Notes |
|---|---|---:|---:|---|
| `p:sp` (`a:prstGeom`) | `shape` / `text` | Yes | Yes | Preset shape support depends on `@prst` (see section 2). |
| `p:cxnSp` | `line` | Yes | Yes | Dash styles and line-end arrows are supported. |
| `p:pic` | `image` | Yes | Yes | Embedded image (`dataUri`) is rendered. |
| `p:graphicFrame/a:tbl` | `table` | Yes | Yes | Basic cell fill, border, text. |
| `p:graphicFrame/c:chart` | `chart` | Yes | No | Pie chart is implemented on Canvas only. |
| `p:graphicFrame/dgm:*` (SmartArt) | `diagram` | Yes (simplified) | No | Extracted as simplified drawable child shapes. |

## 2) Preset Geometry (`a:prstGeom/@prst`) Support

Current explicit shape handlers:

| `@prst` value | Canvas | SVG | Status | Notes |
|---|---:|---:|---|---|
| `rect` | Yes | Yes | Implemented | Default rectangle. |
| `roundRect` | Yes | Yes | Implemented (approx) | Corner radius is approximated. |
| `ellipse` | Yes | Yes | Implemented | Ellipse/circle box fit. |
| `triangle` | Yes | Yes | Implemented | Isosceles triangle primitive. |
| `rtTriangle` | Yes | Yes | Implemented | Right triangle primitive. |
| `diamond` | Yes | Yes | Implemented | Rhombus primitive. |
| `parallelogram` | Yes | Yes | Implemented | Slanted quadrilateral primitive. |
| `trapezoid` | Yes | Yes | Implemented | Top-narrow trapezoid approximation. |
| `pentagon` | Yes | Yes | Implemented | Regular-ish 5-side approximation. |
| `hexagon` | Yes | Yes | Implemented | Symmetric 6-side approximation. |
| `chevron` | Yes | Yes | Implemented | Basic chevron approximation. |
| `line` / `*Connector*` presets | Yes | Yes | Implemented | Includes `straightConnector1` and connector-like presets. |

Fallback behavior:
- Any other `@prst` value is currently rendered as `rect` while keeping style/text.
- This keeps editability and visibility, but does not preserve exact geometry.

## 3) Line Style (`a:ln`) Support

| Feature | Canvas | SVG | Notes |
|---|---:|---:|---|
| `a:ln/@w` (stroke width) | Yes | Yes | Rendered using EMU-based line width. |
| `a:ln/@cap` (`flat`/`sq`/`rnd`) | Yes | Yes | Mapped to renderer line cap. |
| `a:prstDash` | Yes | Yes | Supports `dot`, `dash`, `lgDash`, `dashDot`, `sysDashDotDot`, etc. |
| `a:custDash` | Yes | Yes | Parsed and rendered as custom dash pattern. |
| `a:headEnd` / `a:tailEnd` | Yes | Yes | Supports `triangle`, `stealth`, `diamond`, `oval`, `arrow`. |

## 4) Custom Geometry (`a:custGeom`) Support

| Feature | Canvas | SVG | Notes |
|---|---:|---:|---|
| `a:custGeom/a:avLst` guides | Yes (partial) | Yes (partial) | Common formula operators, angle constants, and iterative guide resolution are supported. |
| `a:path` + `moveTo` / `lnTo` | Yes | Yes | Implemented. |
| `a:path` + `quadBezTo` / `cubicBezTo` | Yes | Yes | Implemented. |
| `a:path` + `close` | Yes | Yes | Implemented. |
| `a:path` + `arcTo` | Yes (approx) | Yes (approx) | Uses OOXML-angle conversion for elliptical arcs; still not fully Office-identical. |
| Roundtrip writeback | Yes | Yes | `a:custGeom` is preserved/written from model geometry. |

## 5) Rendering Parity Notes

- Canvas is the primary target for fidelity improvements.
- SVG currently supports the core shape/text/image/table path, but advanced objects (chart/diagram) are partial.
- When adding new `@prst` handlers, update both renderers and this table together.
