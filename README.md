# Markdown to PowerPoint Converter

A Python application that converts Markdown files into PowerPoint presentations. This tool supports various Markdown elements including headers, bullet points, tables, and bold text formatting.

## Features

- Convert Markdown text or files to PowerPoint (.pptx) format
- Support for slide separation using `---`
- Header levels mapping to slide titles and content headers
- Bullet points with indentation support
- Table parsing and rendering
- Bold text formatting (`**bold**`)
- Web-based UI using Streamlit
- Automatic file naming to avoid overwrites

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
```

2. Install required dependencies:
```bash
pip install python-pptx streamlit
```

## Usage

### Command Line Usage

```bash
python mdtopptx.py
```

By default, the script looks for a markdown file at `.\input\sample.md` and outputs the presentation to the `./output` directory.

### Web Interface Usage

Run the Streamlit web interface:

```bash
streamlit run webui.py
```

The web interface provides two ways to convert Markdown to PowerPoint:
1. **Text Input**: Paste your Markdown content directly into the text area
2. **File Upload**: Upload a [.md](file://c:\workspace\pycodespace\abc\input\sample.md) or `.markdown` file

## Markdown Syntax Support

The converter supports the following Markdown elements:

### Slide Separators
Use horizontal rules (`---`) to separate slides:
```markdown
# First Slide Content

---

## Second Slide Content
```

### Headers
- `#` and `##` create new slides with the header as title
- `###`, `####`, etc. become content headers within slides

### Bullet Points
Support for both `-` and `*` bullet styles with indentation:
```markdown
- First item
- Second item
  - Indented item
    - Double indented item
```

### Tables
GitHub Flavored Markdown table syntax:
```markdown
| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |
```

### Text Formatting
Bold text using double asterisks:
```markdown
This is **bold text**.
```

### Paragraphs
Regular text paragraphs will be added to slides as normal text.

## Example Markdown

```markdown
# My Presentation

---

## Introduction

Welcome to my presentation!

- Point 1
- Point 2
  - Sub-point
- **Important point**

---

## Data Overview

| Metric | Value |
|--------|-------|
| Users  | 1000  |
| Rating | 4.5   |
```

## Project Structure

```
├── MarkdownToPPTX             # Core
│   └── MarkdownToPPTX.py      # Core conversion logic
├── webui                      # webui 
│   └── webui.py     # Streamlit web interface
├── input/           # Default input directory
│   └── sample.md    # Sample markdown file
├── output/          # Default output directory
└── README.md        # This file
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contact

For questions or feedback, please contact:
- GitHub: [myklabs](https://github.com/myklabs)

- Email: mikkel03@gmail.com
