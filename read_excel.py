import pandas as pd

# Read Excel file
file_path = "Interarch Building Solutions Ltd l Valuation template l Dec 25 (1).xlsx"
xl = pd.ExcelFile(file_path)

with open("excel_structure.txt", "w", encoding="utf-8") as f:
    f.write("=" * 60 + "\n")
    f.write("SHEET NAMES IN EXCEL FILE:\n")
    f.write("=" * 60 + "\n")
    for i, name in enumerate(xl.sheet_names):
        f.write(f"  {i}: {name}\n")
    
    f.write("\n" + "=" * 60 + "\n")
    f.write("PREVIEWING EACH SHEET:\n")
    f.write("=" * 60 + "\n")
    
    for sheet_name in xl.sheet_names:
        f.write(f"\n{'='*60}\n")
        f.write(f"SHEET: {sheet_name}\n")
        f.write(f"{'='*60}\n")
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        f.write(f"Shape: {df.shape} (rows, cols)\n\n")
        # Show first 60 rows to capture all relevant data
        f.write(df.head(60).to_string())
        f.write("\n\n")

print("Output saved to excel_structure.txt")
