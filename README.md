# drawio2pptx
[![License: MIT](https://img.shields.io/pypi/l/drawio2pptx)](https://opensource.org/licenses/MIT)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/drawio2pptx)](https://pypi.org/project/drawio2pptx)
[![GitHub Release](https://img.shields.io/github/release/mashu3/drawio2pptx?color=orange)](https://github.com/mashu3/drawio2pptx/releases)
[![PyPi Version](https://img.shields.io/pypi/v/drawio2pptx?color=yellow)](https://pypi.org/project/drawio2pptx/)
[![Downloads](https://static.pepy.tech/badge/drawio2pptx)](https://pepy.tech/project/drawio2pptx)

**Convert your draw.io diagrams to PowerPoint presentations!** 🎨➡️📊

## 📖 Overview

drawio2pptx is a Python package that converts draw.io (diagrams.net) files to PowerPoint (.pptx) presentations. It performs conversion from **mxGraph** (the underlying format used by draw.io) to **PresentationML** (the XML format used by PowerPoint).

**Important**: One draw.io file corresponds to one PowerPoint presentation. Each page/diagram within the draw.io file becomes a separate slide in the resulting PowerPoint presentation.

---

## ✨ Features

### 🔧 Core Functionality
- ✅ Convert draw.io files (.drawio, .xml) to PowerPoint (.pptx)
- ✅ **One file = One presentation**: One draw.io file becomes one PowerPoint presentation
- ✅ **One page/diagram = One slide**: Each page/diagram becomes a separate slide
- ✅ Support for multiple pages/diagrams in a single file
- ✅ Automatic page size configuration (pageWidth, pageHeight)
- ✅ Z-order preserved (shapes and connectors drawn in draw.io order; connectors kept above endpoints when needed)

### 🔷 Shape Support
- **Basic**: Rectangle, Square, Ellipse, Circle, Rounded Rectangle, Triangle (isosceles), Right Triangle, Hexagon, Octagon, Pentagon, Rhombus, Parallelogram, Trapezoid, Star (4/5/6/8-point), Smiley
- **3D**: Cylinder
- **Special**: Cloud, Swimlane (horizontal/vertical with header), BPMN (rhombus / parallel gateway)
- **Flowchart**: Process, Decision, Data, Document, Predefined Process, Internal Storage, Punched Tape, Stored Data, Manual Input, Extract, Merge
- **Connectors**: Straight and orthogonal (elbow) lines; connection points (exit/entry); line styles (dashed, dotted, etc.); arrows (type, size, fill; open oval supported)

### 🎨 Styling & Formatting
- **Colors**: Hex (#RRGGBB, #RGB), RGB, light-dark format
- **Fill**: Solid, gradient, transparent, default theme; corner radius (rounded rectangles)
- **Stroke**: Color, width, styles (solid, dashed, dotted, dash-dot, dash-dot-dot)
- **Text**: Font size, family, bold/italic/underline, horizontal/vertical alignment, padding, wrapping; plain and rich text (HTML: font, b, i, u, strong, em); line breaks; font color from style/HTML
- **Effects**: Shadow, text background color (highlight)

### 📊 Feature Status

This project is in **alpha** and under active development. For a detailed checklist of implemented and planned features, see [FEATURES.md](FEATURES.md).

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

The sample `.drawio` files included in this repository were created by the author
for demonstration and testing purposes only.

They do not include any source code or assets from diagrams.net (draw.io).
Any third-party icons used in the diagrams remain the property of their
respective owners.

---

## 🤝 Contributing

Issues are welcome and encouraged for reporting bugs, proposing features, or sharing ideas.

Currently, pull requests are not accepted, as development is being handled solely by the maintainer.

---

## 📝 License

MIT License

See LICENSE file for details (or check pyproject.toml for license information).

---

## 👨‍💻 Author

[mashu3](https://github.com/mashu3)