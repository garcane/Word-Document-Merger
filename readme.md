# 📄 Word Document Merger

Combine multiple `.docx` files into one master document and export it as **DOCX**, **PDF**, and/or **Markdown** – complete with automatic Table of Contents, page breaks, image/table preservation, and progress bars.

---

## 🚀 Quick Start

### 1. Install Python dependencies
```bash
pip install python-docx pypandoc tqdm
```

### 2. Install system tools for PDF & Markdown
- **Pandoc** (required for Markdown, optional for PDF)  
  → [pandoc.org/installing.html](https://pandoc.org/installing.html)

- **PDF engine** (choose one):
  - **Microsoft Word** (Windows/Mac) + `pip install docx2pdf`  
  - **LibreOffice** (free, cross‑platform) – `libreoffice` command must be in PATH  
  - **Pandoc + WeasyPrint** (`pip install weasyprint`) or **wkhtmltopdf**

### 3. Run the merger
```bash
python word_merger.py
```

A folder picker will open (or you can type a path). Then choose your output formats.

---

## ✨ Features

- 📁 **Folder picker** – select any folder containing `.docx` files (or drag & drop)
- 🔤 **Smart sorting** – by filename (alphabetically) or last modified date
- 📑 **Automatic Table of Contents** – inserts a Word TOC field (update it in Word/LibreOffice)
- 📄 **Preserves formatting** – images, tables, headers, text styles are kept intact
- 📄 **Page breaks** – optionally insert a page break before each appended document
- 📊 **Progress bars** – visual feedback during merging (if `tqdm` is installed)
- 🌍 **Cross‑platform** – works on Windows, macOS, and Linux
- ⚙️ **Multiple output formats** – DOCX, PDF, Markdown (all three, or any subset)

---

## 🔧 Usage

### Interactive mode (no arguments)
```bash
python word_merger.py
```
A dialog asks you to:
- Pick a folder
- Sort by `name` or `date`
- Choose output: `1` PDF only, `2` Markdown only, `3` PDF + Markdown, `4` DOCX only, `5` All three
- Add a Table of Contents? (Y/n)
- Insert page breaks? (Y/n)

### Command‑line mode
```bash
python word_merger.py "C:\Documents\Reports" --sort date --formats pdf md
```

#### All CLI options
```
usage: word_merger.py [-h] [-o OUTPUT] [-f {docx,pdf,md} [{docx,pdf,md} ...]]
                      [--sort {name,date}] [--no-toc] [--no-page-breaks]
                      [--interactive]
                      [folder]

positional arguments:
  folder                Folder containing .docx files

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Output directory (default: <input>/Output)
  -f {docx,pdf,md} [{docx,pdf,md} ...], --formats {docx,pdf,md} [{docx,pdf,md} ...]
                        Output formats (default: all three)
  --sort {name,date}    Sorting order (default: name)
  --no-toc              Do not insert a Table of Contents
  --no-page-breaks      Do not insert page breaks between documents
  --interactive         Force interactive mode even with a folder argument
```

### Output structure
```
InputFolder/
├── Chapter1.docx
├── Chapter2.docx
└── Output/
    ├── Combined.docx
    ├── Combined.pdf
    └── Combined.md
```

If Markdown is generated, images are extracted into an `Output/media` subfolder.

---

## 📦 Full Installation Guide

<details>
<summary><b>Windows</b></summary>

1. Install Python from [python.org](https://www.python.org/downloads/) (✅ Add to PATH).
2. Open **Command Prompt** and install Python packages:
   ```bash
   pip install python-docx pypandoc tqdm
   ```
3. Install Pandoc:
   - Download installer from [pandoc.org](https://pandoc.org/installing.html) and run it.
4. Choose a PDF method:
   - **If you have Microsoft Word**, install `docx2pdf`:
     ```bash
     pip install docx2pdf
     ```
   - **If not**, install LibreOffice from [libreoffice.org](https://www.libreoffice.org/download/) and make sure it’s in your PATH.
</details>

<details>
<summary><b>macOS</b></summary>

```bash
# Install Homebrew if you haven't
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install Python and required libraries
brew install python pandoc
pip3 install python-docx pypandoc tqdm

# For PDF: either use Microsoft Word + docx2pdf, or LibreOffice
brew install --cask libreoffice
# or
pip3 install docx2pdf   # requires Word installed
```
</details>

<details>
<summary><b>Linux (Debian/Ubuntu)</b></summary>

```bash
sudo apt update
sudo apt install python3 python3-pip pandoc libreoffice
pip3 install python-docx pypandoc tqdm
# Optional: WeasyPrint for headless PDF
sudo apt install weasyprint
pip3 install weasyprint
```
</details>

---

## 🛠️ Troubleshooting

- **`'Section' object has no attribute 'headers'`**  
  The latest version of the script has been fixed. If you still see it, make sure you’re using the corrected `merge_documents()` function from this repository.

- **Pandoc not found**  
  Run `pandoc --version` in your terminal. If it’s missing, install it from [pandoc.org](https://pandoc.org/installing.html) and restart your terminal.

- **PDF generation fails**  
  1. **docx2pdf**: requires Microsoft Word to be installed (Windows/Mac).  
  2. **LibreOffice**: ensure `libreoffice` is available in your PATH.  
  3. **Pandoc fallback**: install `weasyprint` (`pip install weasyprint`) and its system dependencies (e.g., `libpango`, `libcairo`).  
  The script tries all three automatically and reports which one succeeded.

- **Markdown images missing**  
  Ensure you have Pandoc ≥ 2.17 and that the `--extract-media` flag works (automatically set). Images appear in `Output/media`.

---

## 📝 License

This project is provided as‑is, free to use for personal or commercial purposes.

---

## 🤝 Contributing

Pull requests welcome. If you find a bug or have a feature request, open an issue with details.

---

**Enjoy effortless document merging!** 🎉
```

This README covers the essentials, installation steps for all platforms, CLI flags, troubleshooting, and the fix for the earlier error.
