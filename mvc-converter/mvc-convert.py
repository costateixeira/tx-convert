import pandas as pd

def load_metadata_from_xlsx(file_path):
    # Load the metadata from the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Remove rows that are completely empty
    df = df.dropna(how='all')

    # Strip whitespace and ensure consistent data types for Value Set names and OIDs
    df['Value Set name'] = df['Value Set name'].str.strip()
    df['CodeSystem OID'] = df['CodeSystem OID'].astype(str).str.strip()

    # Create dictionaries for Value Set and CodeSystem lookups
    metadata = {
        str(row['Value Set name']): {
            'title': str(row['ValueSet Title']),
            'description': str(row['ValueSet Description']).replace('"', "'"),  # Replace double quotes with single quotes
            'package': row['Package']
        }
        for index, row in df.iterrows() if not pd.isna(row['Value Set name'])
    }

    codesystem_lookup = {
        str(row['CodeSystem OID']): str(row['CodeSystem URL'])
        for index, row in df.iterrows() if not pd.isna(row['CodeSystem OID']) and not pd.isna(row['CodeSystem URL'])
    }

    return metadata, codesystem_lookup

def create_fsh_files_from_xlsx(file_path, metadata_file):
    # Load metadata and codesystem lookup from the metadata Excel file
    metadata, codesystem_lookup = load_metadata_from_xlsx(metadata_file)

    # Load the main Excel file
    xlsx = pd.ExcelFile(file_path)
    
    # Initialize sets to store unique names and OIDs not found in the metadata
    unknown_names = set()
    unknown_oids = set()

    # Iterate through the sheets
    for sheet_name in xlsx.sheet_names:
        if sheet_name.startswith("eHDSI"):
            # Load the sheet into a DataFrame without headers
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            print("Processing tab: " + sheet_name)

            # Extract the Value Set name from the data
            value_set_name = str(df.iloc[1, 1]).strip() if not pd.isna(df.iloc[1, 1]) else "UNKNOWN"

            # Check if the Value Set name is in the metadata and if the Package value is 1
            if value_set_name in metadata and metadata[value_set_name]['package'] == 1:
                # Get title and description from metadata
                title = metadata[value_set_name]['title']
                description = metadata[value_set_name]['description']

                # Construct the initial FSH content
                fsh_content = (
                    f"ValueSet: {value_set_name}\n"
                    f"Id: {value_set_name}\n"
                    f"Title: \"{title}\"\n"
                    f"Description: \"{description}\"\n\n"
                    f"Description: \"* ^experimental = false\"\n"
                    

                )

                # Always include the identifier system line
                fsh_content += "* ^identifier.system = \"urn:ietf:rfc:3986\"\n"
                fsh_content += f"* ^identifier.value = \"urn:uuid:{value_set_name}\"\n"

                # Append a new line after the extensions
                fsh_content += "\n"

                # Add row data from row 8 onwards, using the CodeSystem OID from column A
                for index, row in df.iterrows():
                    if index >= 7 and not pd.isna(row[0]):
                        code_system_oid = str(row[0]).strip()  # OID from column A
                        code_system_url = codesystem_lookup.get(code_system_oid, "UNKNOWN_CS")

                        if code_system_url == "UNKNOWN_CS":
                            unknown_oids.add(code_system_oid)

                        # Extract concept code and description from columns B and C
                        concept_code = str(row[2])
                        description_fsn = str(row[3]).replace('"', "'")  # Replace double quotes with single quotes
                        fsh_content += f"* {code_system_url}#{concept_code} \"{description_fsn}\"\n"

                # Save the FSH content to a file
                file_name = f"{sheet_name}.fsh"
                with open(file_name, "w", encoding="utf-8") as file:
                    file.write(fsh_content)
            else:
                print(f"Skipping ValueSet: {value_set_name} (Package != 1 or not found in metadata)")
                unknown_names.add(value_set_name)

    # Save the list of unknown Value Set names and OIDs to CSV files
    if unknown_names:
        unknown_names_df = pd.DataFrame(list(unknown_names), columns=["Unknown Value Set Name"])
        unknown_names_df.to_csv("unknown_names.csv", index=False)
        print("Unknown Value Set names saved to unknown_names.csv")

    if unknown_oids:
        unknown_oids_df = pd.DataFrame(list(unknown_oids), columns=["Unknown CodeSystem OID"])
        unknown_oids_df.to_csv("unknown_oids.csv", index=False)
        print("Unknown CodeSystem OIDs saved to unknown_oids.csv")

# Example usage
create_fsh_files_from_xlsx("MVC801.xlsx", "MVC801-metadata.xlsx")
