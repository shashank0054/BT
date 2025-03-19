import pandas as pd
import json

# Define input and output file names
file_path = "goldenset.xlsx"
output_top_level = "top_level_data.json"
output_nested = "nested_data.json"
output_merged = "merged_data.json"

# Step 1: Extract Top-Level Fields
def extract_top_level():
    df = pd.read_excel(file_path, sheet_name="Sheet1", skiprows=2)
    df.columns = pd.read_excel(file_path, sheet_name="Sheet1", nrows=1).iloc[0].fillna("").str.lower().str.replace(" ", "_")

    top_level_fields = [col for col in df.columns if "this_period" not in col and "ytd_period" not in col]

    json_data = df[top_level_fields].dropna(how="all").apply(lambda row: row.dropna().astype(str).to_dict(), axis=1).tolist()

    with open(output_top_level, "w") as json_file:
        json.dump(json_data, json_file, indent=4)
    
    print(f"✅ Top-level data saved as '{output_top_level}'")

# Step 2: Extract Nested Fields
def extract_nested():
    df = pd.read_excel(file_path, sheet_name="Sheet1", header=[0, 1], skiprows=1)
    
    df.columns = pd.MultiIndex.from_tuples([
        (str(col[0]).strip().lower().replace(" ", "_"), 
         str(col[1]).strip().lower().replace(" ", "_") if isinstance(col[1], str) else "")
        for col in df.columns
    ])

    nested_columns = {
        "this_period": [col for col in df.columns if col[0] == "this_period" and col[1]],
        "ytd_period": [col for col in df.columns if col[0] == "ytd_period" and col[1]],
    }

    def row_to_json(row):
        return {
            section: {field[1]: str(row[field]) for field in cols if pd.notna(row[field])}
            for section, cols in nested_columns.items()
        }

    json_data = df.apply(row_to_json, axis=1).tolist()

    with open(output_nested, "w") as json_file:
        json.dump(json_data, json_file, indent=4)

    print(f"✅ Nested data saved as '{output_nested}'")

# Step 3: Merge JSON Files
def merge_json():
    with open(output_top_level, "r") as top_file:
        top_level_data = json.load(top_file)

    with open(output_nested, "r") as nested_file:
        nested_data = json.load(nested_file)

    if len(top_level_data) != len(nested_data):
        raise ValueError("Mismatch in record count between top-level and nested data.")

    merged_data = [{**top, **nested} for top, nested in zip(top_level_data, nested_data)]

    with open(output_merged, "w") as output_file:
        json.dump(merged_data, output_file, indent=4)

    print(f"✅ Merged JSON saved as '{output_merged}'")

# Run all steps in sequence
if __name__ == "__main__":
    print("🔄 Processing Excel file...")
    extract_top_level()
    extract_nested()
    merge_json()
    print("🎉 All steps completed successfully!")
