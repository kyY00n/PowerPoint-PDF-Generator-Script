# README for PowerPoint PDF Generator Script

## Overview

This Python script automates the creation of personalized PowerPoint presentations and their conversion into PDF format. It dynamically reads names from a selected CSV file, applies these names to placeholders in PowerPoint templates, and generates a PDF for each customized presentation.

## Key Features

- **Automated Personalization**: Uses names from a CSV file to personalize PowerPoint templates.
- **PDF Conversion**: Converts each customized PowerPoint presentation into a PDF file.
- **Template Selection**: Chooses between short and long name templates based on name length.
- **Dynamic CSV File Selection**: Lists all CSV files in the current directory for user selection.
- **Interactive Directory Confirmation**: Allows users to confirm or change the working directory.

## Requirements

- Python with the following packages:
  - `os` (standard library)
  - `comtypes`
  - `colorama`
  - `python-pptx`
  - `pandas`

## Installation

To install the necessary Python packages, use pip:

```bash
pip install python-pptx pandas colorama
```

**Note**: The `comtypes` package may require installation with administrative privileges. To install `comtypes`, run:

```bash
pip install comtypes
```

If you encounter permission issues, try running the command in an administrative command prompt or add the `--user` flag to install it in the user directory:

```bash
pip install comtypes --user
```

## For Mac users

If you are using MacOS, you might get an error like below.
```
ImportError: COM technology not available (maybe it's the wrong platform).
Note that COM is only supported on Windows.
For more details, please check: https://learn.microsoft.com/en-us/windows/win32/com
```

### Requirements

To avoid COM technology error, I revised converting logic in function `create_pdf_from_ppt` with `libreoffice`. So we don't need `comtypes` anymore.

`libreoffice`
  - can be installed by homebrew: `brew install --cask libreoffice`
  - FYI: How to use in command line : `soffice [command] [argument]`

#### pip install

```bash
pip install -r 'requirements-mac.txt'
```

## Usage

1. **Ensure Required Files**: Place the `required.pptx` template and `required.csv` are in the same directory as the script.
2. **Setup Environment**: Ensure you are running this script on a windows machine with PowerPoint installed. Install all other dependancies.
3. **Run the Script**: Execute `python converter.py` in your command line.
4. **Select a CSV File**: Choose a CSV file from the listed options in the current directory.
5. **Select a PPTX File**: Choose a PPTX file from the listed options in the current directory.
6. **Check Output**: The script generates PDFs in an `output_pdfs` folder within the same directory.

## Script Workflow

- **CSV File Selection**: Lists all CSV files in the current directory for the user to select.
- **PPTX File Selection**: Lists all PPTX files in the current directory for the user to select.
- **Name Processing**: Reads names and converts to PDF

