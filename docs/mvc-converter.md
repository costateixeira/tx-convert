# README: FSH File Generator from Excel Data

This Python script processes ValueSet information from an Excel file and generates FSH (FHIR Shorthand) files based on provided metadata and data mappings. It is designed for transforming structured healthcare terminology data into a standardized format that can be used in FHIR implementations.

## Functional Description

The script performs the following functions:

1. **Load Metadata**: Reads metadata from an Excel file (`MVC801-metadata.xlsx`) that contains mappings of ValueSet names to titles, descriptions, and CodeSystem URLs.
2. **Process ValueSet Sheets**: Reads a main Excel file (`MVC801.xlsx`) that contains multiple sheets, each representing a ValueSet. The script processes only sheets whose names start with "eHDSI".
3. **Generate FSH Content**:
   - Constructs FSH files using the ValueSet information from the main Excel file and the metadata.
   - Replaces double quotes in descriptions with single quotes to ensure proper formatting in FSH.
   - Adds row-specific concept codes and descriptions from the main Excel data.
4. **Handle Missing Data**:
   - Collects ValueSet names and CodeSystem OIDs that could not be found in the metadata and writes them to CSV files for review.

## Requirements

- Python 3.6 or higher
- The following Python packages:
  - `pandas`
  - `openpyxl`

To install the required packages, you can run:

```bash
pip install pandas openpyxl
```

## How to Use

1. **Setup**: Ensure you have both `MVC801.xlsx` and `MVC801-metadata.xlsx` in the same directory as the script.
   - `MVC801.xlsx`: The main data file containing multiple sheets of ValueSet information.
   - `MVC801-metadata.xlsx`: The metadata file that maps ValueSet names to titles, descriptions, and CodeSystem URLs.

2. **Running the Script**:
   - Open a terminal or command prompt and navigate to the directory containing the script and the Excel files.
   - Run the script using Python:
     ```bash
     python mvc-convert.py
     ```
   - The script will process all sheets in `MVC801.xlsx` that start with "eHDSI" and generate corresponding `.fsh` files.

3. **Output**:
   - The script generates an FSH file for each processed sheet, saved in the same directory as the script.
   - If any ValueSet names or CodeSystem OIDs cannot be found in the metadata, they will be logged in `unknown_names.csv` and `unknown_oids.csv`, respectively.

## Script Logic Overview

1. **Loading Metadata**:
   - The script reads `MVC801-metadata.xlsx` to create two dictionaries:
     - `metadata`: Maps ValueSet names to titles, descriptions, and the package indicator.
     - `codesystem_lookup`: Maps CodeSystem OIDs to their corresponding URLs.
   - Double quotes in descriptions are replaced with single quotes to ensure proper FSH formatting.

2. **Processing ValueSet Sheets**:
   - The script iterates through all sheets in `MVC801.xlsx` that start with "eHDSI".
   - For each sheet, it extracts the ValueSet name and checks if it exists in the `metadata` dictionary and if the package indicator is `1`.
   - If valid, it constructs the FSH file content using the title, description, and concept codes from the data.

3. **Generating FSH Files**:
   - Each FSH file includes:
     - ValueSet name, ID, title, description, and identifier system.
     - Concept codes and descriptions from the main data, mapped to the appropriate CodeSystem URL.
   - If a CodeSystem OID from the data is not found in the metadata, it is logged.

4. **Handling Missing Data**:
   - Unmatched ValueSet names are written to `unknown_names.csv`.
   - Unmatched CodeSystem OIDs are written to `unknown_oids.csv`.

## Example Output

### Sample FSH Content

For a ValueSet named `eHDSIActiveIngredient`, the generated FSH file might look like this:

```fsh
ValueSet: eHDSIActiveIngredient
Id: eHDSIActiveIngredient
Title: "ATC Classification"
Description: "The Value Set description here"

* ^identifier.system = "urn:ietf:rfc:3986"
* ^identifier.value = "urn:uuid:eHDSIActiveIngredient"

* http://www.whocc.no/atc#A01AA03 "olaflur"
* http://www.whocc.no/atc#A01AB12 "hexetidine"
* http://www.whocc.no/atc#A01AD "Other agents for local oral treatment"
* http://www.whocc.no/atc#A01AD08 "becaplermin"
* http://www.whocc.no/atc#A02AA "Magnesium compounds"
```

### Generated CSV Files

1. **unknown_names.csv**: Lists ValueSet names that could not be found in the metadata.
2. **unknown_oids.csv**: Lists CodeSystem OIDs that could not be found in the metadata.

## Troubleshooting

- **Missing Dependencies**: If you get an error about missing modules, ensure you have installed all dependencies using `pip install pandas openpyxl`.
- **Incorrect File Paths**: Ensure the Excel files are in the correct location and that the paths are specified correctly in the script.

## Customization

- **Modify the Metadata File**: You can update `MVC801-metadata.xlsx` to add or correct ValueSet mappings.
- **Adjust Script Logic**: If you need to change the criteria for processing ValueSets or how the FSH content is constructed, you can modify the script accordingly.

## License

This script is provided under the Creative Commons License.
