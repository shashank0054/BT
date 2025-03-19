import pandas as pd
import json

# Function to extract data from Excel and convert it to JSON
def extract_json_from_excel(file_path, output_path="output.json"):
    # Load the Excel file with the first two rows as headers
    df = pd.read_excel(file_path, sheet_name=0, header=[0, 1])

    # Extract top-level fields from row 0 (first row)
    top_level_columns = [col[0] for col in df.columns if col[0] not in ["this_period", "ytd_period"]]

    # Extract nested fields from row 1 (second row)
    nested_columns = {
        "this_period": [col for col in df.columns if col[0] == "this_period" and col[1]],
        "ytd_period": [col for col in df.columns if col[0] == "ytd_period" and col[1]],
    }

    # Convert top-level fields to JSON format
    top_level_json = df[top_level_columns].dropna(how="all").apply(lambda row: row.dropna().astype(str).to_dict(), axis=1).tolist()

    # Convert nested fields to JSON format
    def row_to_json(row):
        return {
            section: {field[1]: str(row[field]) for field in cols if pd.notna(row[field])}
            for section, cols in nested_columns.items()
        }

    nested_json = df.apply(row_to_json, axis=1).tolist()

    # Merge both JSON structures
    final_json = [{**top, **nested} for top, nested in zip(top_level_json, nested_json)]

    # Save the JSON to a file
    with open(output_path, "w") as json_file:
        json.dump(final_json, json_file, indent=4)

    print(f"✅ JSON extracted and saved as '{output_path}'")
    return output_path

# Example usage
if __name__ == "__main__":
    input_excel = "goldenset.xlsx"  # Replace with your actual file
    output_json = "extracted_data.json"
    extract_json_from_excel(input_excel, output_json)
