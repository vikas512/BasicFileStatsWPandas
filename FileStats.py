import os
import pandas as pd
import xlsxwriter as xw

# Read the CSV file with pipe delimiter
df = pd.read_csv(r"C:\Users\parevi02\pythongpt\ADD_Output.csv", sep='|')

# Check if the file already exists
file_path = "fill_rate_chart.xlsx"
if os.path.exists(file_path):
    os.remove(file_path)

# Create a new workbook and add a worksheet
workbook = xw.Workbook(file_path)
worksheet = workbook.add_worksheet()

# Write the headers to the worksheet
worksheet.write_row("A1", ["column name", "total count", "fill count",  "percent_fill_rate", "blank count","percent blank count", "zero count","percent zero count"])

# Iterate through the columns and calculate fill rate for each column
for i, col in enumerate(df.columns):
    filled_count = df[col].count()
    total_count = len(df)
    fill_rate = filled_count / total_count
    blank_count = df[col].isna().sum()
    zero_count = (df[col] == 0).sum()
    blank_rate = blank_count / total_count
    zero_rate = zero_count / total_count

    # Write the data to the worksheet
    worksheet.write(i + 1, 0, col)
    worksheet.write(i + 1, 1, total_count)
    worksheet.write(i + 1, 2, filled_count)
    worksheet.write(i + 1, 3, fill_rate)
    worksheet.write(i + 1, 4, blank_count)
    worksheet.write(i + 1, 5, blank_rate)
    worksheet.write(i + 1, 6, zero_count)
    worksheet.write(i + 1, 7, zero_rate)

# Save the workbook
workbook.close()
