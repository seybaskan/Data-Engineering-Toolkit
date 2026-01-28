# Universal PDF-to-Excel Data Synchronizer 📊📄

A professional Python automation tool designed to synchronize data between unstructured PDF documents and structured Excel databases through high-performance pattern matching.

## ✨ Core Advantages
- **Industry Agnostic:** Easily adaptable to construction, finance, or legal data by simply modifying the Regex patterns in `CONFIG`.
- **High Performance:** Uses dictionary-based mapping for O(1) time complexity during the matching phase.
- **Data Integrity:** Standardizes Excel headers and handles missing values to ensure clean data migration.

## 🛠 Configuration
To customize the tool for your specific needs, modify the `CONFIG` block in `data_synchronizer.py`:
- `MATCH_KEYS`: Define your primary and secondary identifiers in Excel.
- `REGEX_PATTERNS`: Update these to match the text structure of your specific PDF documents.
- 
## How It Works: Step-by-Step
Environment Setup: The script initializes by reading the CONFIG dictionary, which defines the source files, folder paths, and specific Regex patterns required for the project.

PDF Scanning & Data Mining:

The tool iterates through every PDF file in the specified folder.

It utilizes pdfplumber to extract raw text and applies Regular Expressions (Regex) to isolate unique identifiers (Keys) and the target data (Value).

Extracted data is stored in a high-speed Python dictionary {(Key1, Key2): Value} for efficient lookup.

Excel Standardization:

The source Excel file is loaded into a pandas DataFrame.

Column headers are standardized (trimmed and capitalized) to prevent matching errors caused by hidden spaces or casing.

Data Matching & Merging:

The script performs a row-by-row iteration over the Excel data.

It checks if the current row's keys exist in the PDF database.

If a match is found and the target cell in Excel is empty, it populates the cell with the verified value from the PDF.

Output & Formatting:

The synchronized data is exported to a new Excel file.

(Optional) Updated rows are highlighted in yellow using openpyxl to allow for quick manual verification.
## 🚀 Author
**Şeyda BAŞKAN**
*Civil Engineer & M.Sc. Candidate | Automation Specialist*
