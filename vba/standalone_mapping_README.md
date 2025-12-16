# Standalone Mapping Module

## Overview
The `standalone_mapping.bas` module provides a standalone tool for mapping document sections and normalizing heading styles in Microsoft Word documents. This module can be executed independently without requiring the full manuscript template workflow.

## Purpose
- Maps all sections and headings in the active Word document
- Normalizes heading styles to match standard manuscript formatting conventions
- Generates a detailed report of all headings and their locations
- Can be used as a diagnostic tool or as part of document preparation workflows

## Installation

1. Open Microsoft Word
2. Press `Alt+F11` to open the VBA Editor
3. Go to **File → Import File**
4. Select `standalone_mapping.bas`
5. The module will be added to your document or template

## Usage

### Running the Tool

There are several ways to execute the standalone mapping:

#### Method 1: From VBA Editor (Recommended for testing)
1. Open the VBA Editor (`Alt+F11`)
2. Open a Word document that you want to map
3. In the VBA Editor, press `F5` or go to **Run → Run Sub/UserForm**
4. Select `MapSectionsStandalone` from the list
5. Click **Run**

#### Method 2: From Macro Dialog
1. In Word, press `Alt+F8` to open the Macros dialog
2. Select `MapSectionsStandalone`
3. Click **Run**

#### Method 3: Assign to Quick Access Toolbar
1. Right-click the Quick Access Toolbar
2. Select **Customize Quick Access Toolbar**
3. Choose **Macros** from the dropdown
4. Select `MapSectionsStandalone`
5. Click **Add** then **OK**

## Output

The module produces two types of output:

### 1. Immediate Window Report (Detailed)
Press `Ctrl+G` in the VBA Editor to view the Immediate Window, which displays:
- List of all normalized headings
- Section number for each heading
- Applied Word style for each heading
- Summary statistics (total headings and sections)

Example output:
```
=== DOCUMENT HEADING MAP ===

Section 1: Murtida iyo Maadda (Title)
Section 2: Dedication (Heading 1)
Section 2: Acknowledgments (Heading 2)
Section 3: Table of Contents (Heading 1)
Section 4: Preface (Heading 1)
Section 5: Wisdom Tales (Heading 1)
Section 6: Glossary (Heading 1)
Section 7: About the Author (Heading 1)
Section 8: Copyright Notice (Heading 1)

=== SUMMARY ===
Total headings found and normalized: 9
Total sections: 8
```

### 2. Message Box (Summary)
A popup message confirms completion and provides:
- Number of headings normalized
- Instructions to check the Immediate Window for the full report

## Recognized Headings

The module recognizes and normalizes the following headings:

| Original Text (case-insensitive) | Normalized Form | Applied Style |
|----------------------------------|-----------------|---------------|
| MURTIDA IYO MAADDA | Murtida iyo Maadda | Title |
| DEDICATION | Dedication | Heading 1 |
| ACKNOWLEDGMENTS / ACKNOWLEDGMENTS: | Acknowledgments | Heading 2 |
| TABLE OF CONTENTS / TOC | Table of Contents | Heading 1 |
| PREFACE | Preface | Heading 1 |
| WISDOM TALES | Wisdom Tales | Heading 1 |
| GLOSSARY | Glossary | Heading 1 |
| ABOUT THE AUTHOR | About the Author | Heading 1 |
| COPYRIGHT NOTICE | Copyright Notice | Heading 1 |

## Features

### Independent Execution
- Does not require any other VBA modules or dependencies
- Can be imported and run in any Word document
- No configuration or setup required

### Safe Operation
- Does not modify document content (only normalizes heading styles)
- Includes comprehensive error handling
- Provides clear feedback to the user

### Text Normalization
The module performs the following text cleaning operations:
- Removes line breaks and carriage returns
- Removes tab characters
- Trims leading and trailing whitespace
- Removes trailing colons from heading text

## Use Cases

1. **Document Preparation**: Quickly standardize heading styles across a document before final formatting
2. **Quality Assurance**: Verify that all expected sections are present in a manuscript
3. **Document Analysis**: Generate a structural map of a document's organization
4. **Batch Processing**: Can be incorporated into larger automation workflows

## Technical Details

### Module Structure
- `MapSectionsStandalone()`: Main entry point for the standalone tool
- `NormalizeHeadingParagraph()`: Identifies and normalizes individual headings
- `CleanParagraphText()`: Performs text cleaning and normalization
- `HandleError()`: Centralized error handling and reporting

### Word Styles Used
- `wdStyleTitle`: For the main document title
- `wdStyleHeading1`: For major section headings
- `wdStyleHeading2`: For subsection headings (e.g., Acknowledgments)

## Requirements

- Microsoft Word 2010 or later
- VBA/Macros must be enabled in Word
- An active document must be open when running the tool

## Troubleshooting

**Problem**: "No active document found" error
- **Solution**: Open a Word document before running the macro

**Problem**: Can't see the Immediate Window output
- **Solution**: In VBA Editor, press `Ctrl+G` or go to **View → Immediate Window**

**Problem**: Macro doesn't appear in the list
- **Solution**: Ensure the module was imported correctly and that macros are enabled

## Related Files

- `manuscript_template.bas`: Full manuscript template automation (includes this functionality as part of a larger workflow)
- `README_vba.md`: General VBA documentation for the project

## License

This code is part of the vba-python-automation project.
