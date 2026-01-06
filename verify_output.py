import pandas as pd

# Read and compare Factsheet from updated file
file_path = "Interarch Building Solutions Ltd l Valuation template l Dec 25 (1)_updated.xlsx"
df = pd.read_excel(file_path, sheet_name='Factsheet', header=None)

print("=" * 60)
print("FACTSHEET DATA (Rows 1-60) - UPDATED FILE")
print("=" * 60)
print(df.head(60).to_string())
