#code to take the input of cross reference 8A8Z and search in excel file, highlight and copy into interchange column
import pandas as pd

# Load the Excel file
file_path = 'test1.xlsx'  # Update this with the actual file path
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Define the search pattern
pattern = r'8A8Z'  # Look for '8A8Z' occurring anywhere in the column values

# Initialize a boolean mask for matching rows
matches = pd.Series([False] * len(df))

# Iterate through each column and find matches
for column in df.columns:
    # Check if the column contains the pattern (case-insensitive)
    column_matches = df[column].astype(str).str.contains(pattern, case=False, na=False)
    
    # For rows where a match is found, copy the matching value to 'Cross Reference'
    df.loc[column_matches, 'Cross Reference'] = df.loc[column_matches, column]
    
    # Update the overall matches
    matches |= column_matches

# Save the modified dataframe to a new Excel file
output_path = 'modified_file.xlsx'  # This will be the name of the output file
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Apply highlighting only to rows that matched
    format_highlight = workbook.add_format({'bg_color': 'yellow'})
    for row_num, match in enumerate(matches, start=1):
        if match:
            worksheet.set_row(row_num, cell_format=format_highlight)

print(f"Modified file saved at: {output_path}")
