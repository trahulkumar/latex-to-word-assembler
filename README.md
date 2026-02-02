# latex-to-word-assembler

Automated assembly pipeline that converts fragmented LaTeX chapter files into a unified, structured Microsoft Word manuscript, driven by a simple metadata outline. Built with Python and Pandoc for seamless "Book-ready" document compiling.

## Architecture

The project consists of two main stages:

1.  **Metadata Generation**:
    *   **Input**: `input/Master Production Manuscript.txt` (Markdown format).
    *   **Process**: Parses the manuscript file to extract the book structure (Book Title, Chapters, Sections).
    *   **Output**: `input/metadata.json`. This JSON file acts as the blueprint for the book, mapping logical sections to physical LaTeX files.

2.  **Conversion**:
    *   **Input**: `input/metadata.json` and the corresponding `.tex` files in `input/latex_files/`.
    *   **Process**:
        *   Reads the JSON blueprint.
        *   Validates that the required `.tex` files exist.
        *   Uses `pypandoc` (and the bundled `pandoc` binary) to convert each file and combine them into a single document.
    *   **Output**: A single `.docx` file in the `output/` directory (e.g., `Predictive_Analytics_with_Python.docx`).

## Folder Structure

```
Latex-to-docx beautify/
├── 01_generate_metadata.bat     # Script 1: Generates JSON from text metadata
├── 02_convert_to_docx.bat       # Script 2: Converts Latex files to Docx
├── input/
│   ├── Master Production Manuscript.txt  # Your book outline file
│   ├── metadata.json            # Generated blueprint
│   └── latex_files/
│       ├── Chapter_1/
│       │   ├── Section_1.1.tex
│       │   └── ...
│       ├── Chapter_2/
│       │   └── ...
│       └── ...
├── output/
│   └── [Book_Title].docx        # Final Output
├── src/
│   ├── generate_metadata.py     # Python script for Step 1
│   └── convert_to_docx.py       # Python script for Step 2
├── .venv/                       # Virtual Environment (managed by uv)
├── pyproject.toml               # Dependency definitions
└── uv.lock
```

## Prerequisites

*   **Windows OS** (as per current setup)
*   **uv**: A fast Python package installer and resolver.
    *   Installation: `powershell -c "irm https://astral.sh/uv/install.ps1 | iex"` (or via pip: `pip install uv`)

## How to Use

### Step 1: Prepare Metadata
1.  Open `input/Master Production Manuscript.txt`.
2.  Ensure your book structure follows the Markdown format:
    ```markdown
    **Book Title:** My Book
    ### Chapter 1: Introduction
    * **1.1 First Section**
    ```

### Step 2: Prepare LaTeX Files
1.  Create your `.tex` files in the `latex_files` directory.
2.  The folder structure must match what the metadata generator expects:
    *   `input/latex_files/Chapter_N/Section_N.M.tex`
    *   Example: For "Section 1.1", the file should be at `input/latex_files/Chapter_1/Section_1.1.tex`.

### Step 3: Generate Metadata
1.  Run **`01_generate_metadata.bat`**.
2.  Check `input/metadata.json` to verify the structure and file paths are correct.

### Step 4: Convert
1.  Run **`02_convert_to_docx.bat`**.
2.  The script will verify if the files listed in the JSON exist.
3.  If successful, find your generated Word document in the `output/` folder.

## Customization

*   **Metadata Format**: If you change the format of the manuscript file, update `src/generate_metadata.py` to match the new parsing logic.
*   **Pandoc Options**: To add custom styles (e.g., specific fonts, margins), edit `src/convert_to_docx.py` and modify the `extra_args` list (e.g., adding `--reference-doc=custom_style.docx`).
