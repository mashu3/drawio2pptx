# Feature Status

This document provides a detailed checklist of implemented and planned features for drawio2pptx.

## Core Features

### File Operations

- [x] Load drawio files (.drawio / .xml)
- [x] Output to PowerPoint files (.pptx)
- [ ] Support for compressed drawio files (.drawio)
- [ ] Support for encrypted drawio files

### Diagram Processing

- [x] Support for multiple diagrams (each diagram becomes a separate slide)
- [x] Automatic page size configuration (pageWidth, pageHeight)
- [x] Z-order / stacking order (shapes and connectors drawn in draw.io document order; connectors kept above endpoint shapes when needed)
- [ ] Page orientation (landscape/portrait)
- [ ] Page margin settings
- [ ] Page background color
- [ ] Grid display settings

### Conversion Processing

- [x] Conversion option settings (ConversionConfig fully integrated into main conversion flow)
- [x] Conversion error handling and logging (ConversionLogger implemented)
- [ ] Conversion progress display
- [ ] Batch conversion (convert multiple files at once)

## Color Conversion

### Color Formats

- [x] Hexadecimal format (#RRGGBB)
- [x] Short hexadecimal format (#RGB)
- [x] RGB format (rgb(r,g,b))
- [x] light-dark format (uses light mode color)
- [ ] RGBA format (rgba(r,g,b,a))
- [ ] HSL format (hsl(h,s,l))
- [ ] HSLA format (hsla(h,s,l,a))
- [ ] Named colors (red, blue, green, etc.)
- [ ] PowerPoint theme colors (schemeClr)

### Fill Color (fillColor)

- [x] Color specification (fillColor=#RRGGBB, etc.)
- [x] Transparent (fillColor=none)
- [x] Default (fillColor=default/auto) - uses PowerPoint theme color
- [x] Gradient colors (gradientColor, gradientDirection)
- [ ] Pattern colors

### Stroke Color (strokeColor)

- [x] Color specification (strokeColor=#RRGGBB, etc.)
- [x] Default (PowerPoint default when strokeColor is not specified)
- [x] Transparent (strokeColor=none)
- [ ] Gradient colors

### Font Color (fontColor)

- [x] Extraction from style attribute
- [x] Extraction from style attribute within HTML tags (Square/Circle support)
- [x] Direct attribute extraction (fontColor attribute)
- [x] Default (PowerPoint default when fontColor is not specified)

### Other Colors

- [ ] Shadow color (shadowColor)
- [x] Text background color (labelBackgroundColor → highlight)

## Shape Properties

### Position & Size

- [x] Position (x, y)
- [x] Size (width, height)
- [x] Fixed aspect ratio (aspect=fixed) - Square/Circle support
- [ ] Minimum size (minWidth, minHeight)
- [ ] Maximum size (maxWidth, maxHeight)
- [ ] Size constraints

### Fill

- [x] Fill enabled/disabled
- [x] Fill color (fillColor)
  - [x] Color specification (fillColor=#RRGGBB, etc.)
  - [x] Transparent (fillColor=none)
  - [x] Default (fillColor=default/auto)
- [ ] Transparency (opacity / fillOpacity)
- [x] Gradient (gradientColor, gradientDirection)
  - [x] Linear gradient
  - [ ] Radial gradient
- [ ] Pattern fill
- [ ] Image fill

### Stroke

- [x] Stroke enabled/disabled
- [x] Stroke color (strokeColor)
- [x] Stroke width (strokeWidth) - supported for connectors/edges
- [x] Stroke width (strokeWidth) - supported for regular shapes (rectangles, ellipses, etc.)
- [x] Stroke styles
  - [x] Solid (solid)
  - [x] Dashed (dashed)
  - [x] Dotted (dotted)
  - [x] Dash-dot (dashDot)
  - [x] Dash-dot-dot (dashDotDot)
- [ ] Stroke transparency (strokeOpacity)
- [ ] Stroke line cap (strokeLinecap: round, square, flat)
- [ ] Stroke line join (strokeLinejoin: round, bevel, miter)
- [ ] Stroke miter limit (strokeMiterlimit)

### Shadow & Effects

- [x] Shadow enabled/disabled (cell level)
- [x] Shadow enabled/disabled (mxGraphModel level)
- [ ] Shadow color (shadowColor)
- [ ] Shadow offset (shadowOffsetX, shadowOffsetY)
- [ ] Shadow blur (shadowBlur)
- [ ] Glow effect
- [ ] Blur effect

### Other Properties

- [x] Rotation (rotation)
- [ ] Skew (skew)
- [ ] Transform matrix (transform)
- [x] Corner radius (rounded / arcSize)
- [x] White space handling (whiteSpace: wrap, nowrap)
- [x] Aspect ratio (aspect: fixed for Square/Circle)
- [ ] Auto-size (autosize)
- [ ] Resizable flag (resizable)
- [ ] Movable flag (movable)
- [ ] Editable flag (editable)
- [ ] Selectable flag (selectable)

## Text Properties

### Text Content

- [x] Text content
  - [x] Plain text
  - [x] Extraction from HTML tags (Square/Circle support)
- [x] Rich text (partial HTML format support - font, b, i, u, strong, em tags)
- [x] Text line breaks
- [ ] Special characters and symbols

### Font

- [x] Font size (points / fontSize)
- [x] Font family (fontFamily)
- [x] Bold (bold / fontStyle=1)
- [x] Italic (italic / fontStyle=2)
- [x] Underline (underline / fontStyle=4)
- [ ] Strikethrough
- [ ] Superscript
- [ ] Subscript
- [ ] Letter spacing (letterSpacing)
- [ ] Text transformation (uppercase/lowercase)

### Text Alignment

- [x] Horizontal alignment (left, center, right / align)
- [x] Vertical alignment (top, middle, bottom / verticalAlign)
- [x] Padding (top, bottom, left, right / spacingTop, spacingBottom, spacingLeft, spacingRight)
- [x] Text wrapping (whiteSpace=wrap)
- [ ] Line spacing adjustment (lineSpacing)
- [ ] Paragraph spacing (paragraphSpacing)
- [ ] Indentation (indent)
- [ ] Bullet points (listStyle)
- [ ] Numbered lists

### Text Color

- [x] Font color (fontColor)
  - [x] Extraction from style attribute
  - [x] Extraction from style attribute within HTML tags
  - [x] Direct attribute extraction (fontColor attribute)
  - [x] Default (when fontColor is not specified)
- [x] Text background color (labelBackgroundColor → highlight)
- [ ] Text gradient

### Other Text Properties

- [ ] Text auto-size (autosize)
- [ ] Text clipping
- [ ] Text rotation
- [ ] Text transformation

## Shape Types

### Basic Shapes

- [x] Rectangle (RECTANGLE / rect / rectangle)
- [x] Square (SQUARE / square)
  - [x] Processed as rectangle
  - [x] Fixed aspect ratio (aspect=fixed) support
- [x] Ellipse (ELLIPSE / ellipse)
- [x] Circle (CIRCLE / circle)
  - [x] Processed as ellipse
  - [x] Fixed aspect ratio (aspect=fixed) support
- [x] Rounded rectangle (rounded / rounded=1)
- [x] Triangle (TRIANGLE / triangle) — isosceles triangle support
- [x] Right triangle (rightTriangle)
- [x] Step (step) — mapped to PowerPoint Chevron; step size (style `size`) supported
- [x] Hexagon (HEXAGON / hexagon)
- [x] Octagon (OCTAGON / octagon)
- [x] Rhombus (RHOMBUS / rhombus)
- [x] Parallelogram (PARALLELOGRAM / parallelogram)
- [x] Trapezoid (TRAPEZOID / trapezoid)
- [x] Pentagon (PENTAGON / pentagon)
- [x] Star (STAR / star) — 4-point, 5-point, 6-point, 8-point star support
- [x] Smiley (smiley / mxgraph.basic.smiley)
- [ ] Cross (CROSS / cross)
- [ ] Plus (PLUS / plus)
- [x] Arrow (ARROW / arrow) — vertex shape (mxgraph.arrows2.arrow → right_arrow)
- [ ] Double arrow (doubleArrow)
- [ ] Curved arrow (curvedArrow)

### 3D Shapes

- [x] Cylinder (CYLINDER / cylinder / cylinder3)
- [x] Cube (CUBE / cube) — mxgraph.infographic.shadedcube with 3D rotation
- [ ] 3D Box (box3d)

### Special Shapes

- [x] Cloud (CLOUD / cloud)
- [x] BPMN shape (mxgraph.bpmn.shape) - rendered as rhombus; parallel gateway symbol (parallelGw) supported
- [ ] Actor (ACTOR / actor)
- [ ] Text label (TEXT / text)
- [ ] Image (IMAGE / image)
- [x] Swimlane (swimlane) - horizontal and vertical swimlanes with header divider support
- [ ] Container (container)

### Flowchart Shapes

- [x] Process (process)
- [x] Decision (decision / diamond) - processed as rhombus
- [x] Data (data / mxgraph.flowchart.data)
- [x] Document (document)
- [ ] Multi-document (multiDocument)
- [x] Predefined process (predefinedProcess)
- [x] Internal storage (internalStorage)
- [ ] Sequential data (sequentialData)
- [ ] Direct access storage (directAccessStorage)
- [x] Manual input (manualInput / mxgraph.flowchart.manual_input)
- [ ] Manual operation (manualOperation)
- [ ] Preparation (preparation)
- [ ] Connector (connector)
- [x] Off-page connector (offPageConnector)
- [ ] Card (card)
- [x] Punched tape (punchedTape / tape)
- [ ] Summing junction (summingJunction)
- [ ] OR (or)
- [ ] Collate (collate)
- [ ] Sort (sort)
- [x] Extract (extract / mxgraph.flowchart.extract)
- [x] Merge (merge / mxgraph.flowchart.merge_or_storage)
- [ ] Offline storage (offlineStorage)
- [ ] Online storage (onlineStorage)
- [ ] Magnetic tape (magneticTape)
- [ ] Display (display)
- [ ] Delay (delay)
- [ ] Alternate process (alternateProcess)
- [x] Stored data (storedData / dataStorage)
- [ ] Terminator (terminator)

### Connectors/Edges

- [x] Basic connector/edge support
- [x] Straight line (straight) - basic implementation
- [x] Orthogonal (orthogonal) - basic implementation
- [x] Elbow connector (elbow) - processed as orthogonal
- [ ] Curved line (curved) - not supported (converted to straight polyline)
- [x] Line styles (dashed, dotted, etc.) - basic implementation
- [x] Connection point settings
  - [x] Start connection point (exitX, exitY, exitDx, exitDy)
  - [x] End connection point (entryX, entryY, entryDx, entryDy)
  - [x] Accurate connection point calculation on shape boundaries
- [x] Line color (strokeColor)
- [x] Line width (strokeWidth)
- [x] Shadow settings (shadow=0/1)
- [x] Arrow settings
  - [x] Start arrow (startArrow)
  - [x] End arrow (endArrow)
  - [x] Arrow fill settings (startFill, endFill); open oval (startFill/endFill=0) emulated with overlay shape
  - [x] Arrow size (startSize/endSize) mapped to PowerPoint sm/med/lg
  - [x] Default end arrow when mxGraphModel arrows=1 and endArrow omitted (classic)
  - [x] Arrow type mapping (triangle, oval, diamond, open, stealth, etc.)
- [x] Labeled connectors (labels rendered as standalone text elements; z-order preserved)

### Other Shapes

- [ ] Polygon (POLYGON / polygon)
- [ ] Freehand (freehand)
- [ ] Curve (curve)
- [ ] Bezier curve (bezier)
- [ ] Grouped shapes (group)
- [ ] Table (table)
- [ ] Shape combination (Union, Subtract, Intersect, Exclude)

## Other Features

### Z-order & Stacking

- [x] Draw order from mxGraphModel (document / parent-child traversal)
- [x] Connectors drawn above endpoint shapes when edge order would otherwise place them behind
- [x] Connector labels keep same z-index as connector

### Grouping & Layers

- [ ] Grouped shapes (group)
- [ ] Individual shape processing within groups
- [ ] Layer support (layer)
- [ ] Layer show/hide
- [ ] Layer lock

### Transformation & Rotation

- [x] Rotation (rotation)
- [ ] Skew (skew)
- [ ] Transformation (transform)
- [x] Flip (horizontal/vertical flip)

### Links & Hyperlinks

- [ ] Link (link) — intermediate model has TextRun.link; not yet applied in PPTX output
- [ ] Hyperlink (hyperlink)
- [ ] Internal links (links to pages within the same file)
- [ ] External links (URLs)

### Tables

- [ ] Table (table)
- [ ] Table cells
- [ ] Table styles
- [ ] Table merged cells

### Shape Combination & Boolean Operations

- [ ] Union
- [ ] Subtract
- [ ] Intersect
- [ ] Exclude

### Images & Media

- [x] Image embedding (image) — ImageElement written to PPTX
- [x] SVG image support — SVG images are automatically converted to PNG format for embedding
  - [x] SVG to PNG conversion using resvg library
  - [x] High-quality rendering with configurable DPI (default 192 DPI)
  - [x] Optimal DPI calculation to ensure minimum 100px short edge
  - [x] Support for SVG images from data URIs and file paths
- [x] Image size adjustment — images maintain their size from draw.io
- [x] Image aspect ratio maintenance — aspect ratio preserved during conversion

### Other

- [ ] Comments & annotations
- [ ] Template shapes
- [ ] Custom shapes
- [ ] Style inheritance
- [ ] Theme application

