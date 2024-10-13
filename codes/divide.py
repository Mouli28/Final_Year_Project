import pandas as pd
import re


file_path = 'E:/ai check/phase2/combined_42_51.xlsx'
df = pd.read_excel(file_path)


columns_to_process = ['Part No From Site', 'Cross Reference'] 

patterns = {
    'SKU': r'SKU\s*#\s*(\S+)',
    'Same As': r'Same As\s*([A-Za-z0-9,\s.-]+)',
    'Part Interchanges': r'Part Interchanges\s*([A-Za-z0-9,\s.-]+)',
    'OE Numbers': r'OE Numbers\s*([A-Za-z0-9,\s.-]+)',
    'OE Cross Reference': r'OE Cross Reference\s*([A-Za-z0-9,\s.-]+)'
}


def extract_data(text):
    extracted = {}
    
    for key, pattern in patterns.items():
        match = re.search(pattern, str(text))
        if match:
            extracted[key] = match.group(1).strip()
    
    return extracted

for column in columns_to_process:
    df[f'Extracted_{column}'] = df[column].apply(lambda x: ', '.join([f'{k}: {v}' for k, v in extract_data(x).items()]) if extract_data(x) else None)

output_file = 'output_file.xlsx'
df.to_excel(output_file, index=False)

print("New Excel file with extracted data saved as:", output_file)
