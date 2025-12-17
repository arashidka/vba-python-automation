# VBA Automation Modules

This directory contains VBA (Visual Basic for Applications) modules for automating Microsoft Office documents, particularly Microsoft Word.

## Available Modules

### 1. Standalone Mapping Module (`standalone_mapping.bas`)

A standalone tool for mapping document sections and normalizing heading styles in Word documents.

**Key Features:**
- Maps all sections and headings in a document
- Normalizes heading styles to standard manuscript formatting
- Generates detailed reports of document structure
- Can execute independently without other dependencies

**Documentation:** See [standalone_mapping_README.md](standalone_mapping_README.md) for detailed usage instructions.

**Quick Start:**
1. Import `standalone_mapping.bas` into Word VBA Editor
2. Run `MapSectionsStandalone()` macro
3. View results in the Immediate Window (Ctrl+G)

### 2. Manuscript Template Module (`manuscript_template.bas`)

A comprehensive manuscript template automation tool for creating formatted documents with standard sections, styles, and page numbering.

**Status:** Available in PR #1

## Installation

### General VBA Installation Steps

1. Open Microsoft Word
2. Press `Alt+F11` to open the VBA Editor
3. Go to **File → Import File**
4. Select the desired `.bas` file
5. The module will be added to your document or template

### Enabling Macros

If macros are disabled:
1. Go to **File → Options → Trust Center**
2. Click **Trust Center Settings**
3. Select **Macro Settings**
4. Choose **Enable all macros** (or configure as needed)
5. Click **OK**

## Usage

Run macros by:
- Pressing `Alt+F8` and selecting the macro
- Using the VBA Editor (Alt+F11) and pressing F5
- Adding macros to the Quick Access Toolbar

## Requirements

- Microsoft Word 2010 or later
- VBA/Macros must be enabled
- Windows or Mac with Office for Mac

## Project Structure

```
vba/
├── README_vba.md                      # This file
├── standalone_mapping.bas              # Standalone mapping tool
├── standalone_mapping_README.md        # Detailed documentation
└── manuscript_template.bas             # Full manuscript template (PR #1)
```

## Contributing

When adding new VBA modules:
1. Include comprehensive error handling
2. Add detailed comments and documentation
3. Follow the established code structure
4. Create a separate README for complex modules

## Support

For issues or questions, please refer to the individual module documentation or open an issue in the repository.
