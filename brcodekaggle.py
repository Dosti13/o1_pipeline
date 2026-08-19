import pandas as pd

input_file = r"C:\Users\HP\Downloads\upc_corpus.csv\upc_corpus.csv"
input_file1 = r"C:\Users\HP\Downloads\categoriesandfinal_brand_standardized (1).xlsx"

# Read files
df1 = pd.read_csv(input_file, low_memory=False)
df2 = pd.read_excel(input_file1)

# Normalize barcodes
df1["barcode1"] = (
    df1["ean"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.lstrip("0")
)

df2["barcode"] = (
    df2["barcode"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.lstrip("0")
)

# Keep only required columns
df1 = df1[["barcode1", "name"]]
df2 = df2[["barcode", "material_name"]]

# Remove duplicates
df1 = df1.drop_duplicates(subset=["barcode1"])
df2 = df2.drop_duplicates(subset=["barcode"])

# Merge
matched = pd.merge(
    df1,
    df2,
    left_on="barcode1",
    right_on="barcode",
    how="inner"
)

# Rename columns

# Save
matched.to_excel(r"C:\Users\HP\Downloads\matched_materials.xlsx", index=False)

print(matched.head())