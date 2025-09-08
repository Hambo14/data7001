import pandas as pd
import glob
import os

# Step 1: Get a list of all Excel (.xlsx) files in the current folder
file_list = glob.glob("*.xlsx")

# Step 2: Create a new folder "cleaned_data" to save the cleaned CSV files
output_dir = "cleaned_data"
os.makedirs(output_dir, exist_ok=True)

# Step 3: Loop through each Excel file and clean it
for file in file_list:
    print(f"Processing file: {file}")

    # Read Excel file into a DataFrame
    df = pd.read_excel(file)

    # Show basic info about the dataset (columns, types, missing values)
    print(df.info())

    # Standardize column names: lowercase + strip spaces + replace spaces with underscores
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace(r"[^\w_]+", "", regex=True)  # remove special characters
    )

    # Remove duplicate rows
    df = df.drop_duplicates()

    # Handle missing values:
    # - Numeric columns: fill with mean
    # - Other columns: fill with "Unknown"
    for col in df.columns:
        if df[col].dtype in ["int64", "float64"]:
            df[col] = df[col].fillna(df[col].mean())
        else:
            df[col] = df[col].fillna("Unknown")

    # Clean string values: remove extra spaces
    df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

    # Save cleaned file as CSV in the "cleaned_data" folder
    output_file = os.path.join(output_dir, file.replace(".xlsx", "_clean.csv"))
    df.to_csv(output_file, index=False, encoding="utf-8-sig")

    print(f"✅ Cleaned file saved as: {output_file}\n")

print("🎉 All Excel files have been cleaned and saved in 'cleaned_data' folder!")
