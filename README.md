# drawio2pptx
[![License: MIT](https://img.shields.io/pypi/l/drawio2pptx)](https://opensource.org/licenses/MIT)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/drawio2pptx)](https://pypi.org/project/drawio2pptx)
[![PyPi Version](https://img.shields.io/pypi/v/drawio2pptx?color=yellow)](https://pypi.org/project/drawio2pptx/)

**Convert your draw.io diagrams to PowerPoint presentations!** 🎨➡️📊

## 📖 Overview

drawio2pptx is a Python package that converts draw.io (diagrams.net) files to PowerPoint (.pptx) presentations. It performs conversion from **mxGraph** (the underlying format used by draw.io) to **PresentationML** (the XML format used by PowerPoint).

**Important**: One draw.io file corresponds to one PowerPoint presentation. Each page/diagram within the draw.io file becomes a separate slide in the resulting PowerPoint presentation.

---

## ✨ Features

### 🔧 Core Functionality
- ✅ Convert draw.io files (.drawio, .xml) to PowerPoint (.pptx)
- ✅ **One file = One presentation**: One draw.io file becomes one PowerPoint presentation
- ✅ **One page/diagram = One slide**: Each page/diagram in a draw.io file becomes a separate slide
- ✅ Support for multiple pages/diagrams in a single file
- ✅ Automatic page size configuration (pageWidth, pageHeight)

### 🔷 Shape Support
- **Basic shapes**: Rectangle, Square, Ellipse, Circle, Rounded Rectangle
- **Polygons**: Triangle, Right Triangle, Hexagon, Octagon, Pentagon, Rhombus, Parallelogram, Trapezoid, Star (4/5/6/8-point)
- **3D shapes**: Cylinder
- **Flowchart shapes**: Process, Decision, Document, Predefined Process, Internal Storage, Punched Tape, Stored Data, Cloud
- **Connectors/Edges**: Straight lines, Orthogonal connectors with arrow support

### 🎨 Styling & Formatting
- **Colors**: Hexadecimal (#RRGGBB), Short hex (#RGB), RGB format, light-dark format
- **Fill**: Solid colors, gradients, transparent, default theme colors
- **Stroke**: Color, width, styles (solid, dashed, dotted, dash-dot, dash-dot-dot)
- **Text**: Font size, family, style (bold, italic, underline), alignment (horizontal/vertical), padding, wrapping
- **Effects**: Shadow support, text background color (highlight)

### 📝 Text Support
- Plain text and rich text (partial HTML format: font, b, i, u, strong, em tags)
- Text line breaks
- Font color extraction from style attributes and HTML tags

### 📊 Feature Status

This project is under active development. The core conversion functionality is working, and many basic shapes and styling features are supported.

For a detailed checklist of implemented and planned features, see [FEATURES.md](FEATURES.md).

---

## 📦 Installation

### Requirements

- Python 3.8 or higher
- **python-pptx >= 0.6.21**: Used for creating and writing PowerPoint (.pptx) files in PresentationML format
- **lxml >= 4.6.0**: Used for parsing and processing XML/mxGraph data from draw.io files, and for directly editing PresentationML XML elements that are not supported by python-pptx (e.g., gradients, highlights, advanced styling)

### Install Dependencies

```bash
pip install python-pptx lxml
```

### Install as Package (Development Mode)

Install the package in development mode to use the `drawio2pptx` command:

```bash
pip install -e .
```

Or install from PyPI:

```bash
pip install drawio2pptx
```

---

## 🚀 Usage

### Command Line Interface

After installation, use the `drawio2pptx` command:

```bash
drawio2pptx sample.drawio sample.pptx
```

### Example

```bash
drawio2pptx sample.drawio sample.pptx
```

### Alternative: Python Module

If the command is not found, you can run it as a Python module:

```bash
python -m drawio2pptx.main sample.drawio sample.pptx
```

### Analysis Mode

You can use the `--analyze` (or `-a`) option to display analysis results after conversion:

```bash
drawio2pptx sample.drawio sample.pptx --analyze
```

---

## 📄 Sample files

The sample `.drawio` file included in this repository was created by the author
for demonstration and testing purposes only.

It does not include any source code or assets from diagrams.net (draw.io).
Any third-party icons used in the diagram remain the property of their
respective owners.

---

## 🤝 Contributing

Bug reports and feature requests are welcome. Please open an issue on the repository.

---

## 📝 License

MIT License

See LICENSE file for details (or check pyproject.toml for license information).

---

## 👨‍💻 Author

[mashu3](https://github.com/mashu3)