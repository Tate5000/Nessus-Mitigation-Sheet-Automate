import pandas as pd
import re
from collections import defaultdict
import spacy
import openpyxl
import tkinter as tk
from tkinter import filedialog

# Prompt user for file path
def prompt_for_file_path(title="Select a file"):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title=title)
    return file_path

# Basic extraction function
def extract_data(pattern, description):
    match = re.search(pattern, description)
    return match.group(1).strip() if match else None

def extract_keywords(text):
    words = re.findall(r'\b\w+\b', text.lower())
    common_words = ["the", "and", "for", "with", "that", "from", "this", "are", "not", "all", "has", "can", "should"]
    keywords = [word for word in words if len(word) > 3 and word not in common_words]
    return set(keywords)

def infer_800_53_control(description, control_keywords_mapping, nlp_model=None):
    words = set(re.findall(r'\b\w+\b', description.lower()))
    control_matches = defaultdict(int)
    if nlp_model:
        doc = nlp_model(description.lower())
        entities = {ent.lemma_ for ent in doc.ents}
        words = words.union(entities)
    for control, keywords in control_keywords_mapping.items():
        for keyword in keywords:
            if keyword in words:
                control_matches[control] += 1
    return max(control_matches, key=control_matches.get, default="Inferred Control")

def process_data(df, control_keywords_mapping):
    df['Type'] = ''
    df['ID'] = df['Plugin ID']
    df['CAT'] = ''
    df['STIG-ID'] = ''
    df['800-53 Control'] = ''

    critical_mask = df['Risk'] == 'Critical'
    df.loc[critical_mask, 'Type'] = 'Compliance'
    df.loc[critical_mask, 'CAT'] = 'I'
    
    failed_mask = df['Risk'] == 'FAILED'
    df.loc[failed_mask, 'STIG-ID'] = df.loc[failed_mask, 'Description'].apply(lambda x: extract_data(r'STIG-ID\|([^,]+)', x))
    df.loc[failed_mask, '800-53 Control'] = df.loc[failed_mask, 'Description'].apply(lambda x: extract_data(r'800-53\|([^,]+)', x))
    df.loc[failed_mask, 'CAT'] = df.loc[failed_mask, 'Description'].apply(lambda x: extract_data(r'CAT\|([^,]+)', x))
    
    high_mask = df['Risk'] == 'High'
    df.loc[high_mask, 'CAT'] = df.loc[high_mask, 'Description'].apply(lambda x: extract_data(r'CAT\|([^,]+)', x))

    df['ID'] = df['ID'].combine_first(df['STIG-ID'])
    df['800-53'] = df['Description'].apply(lambda x: infer_800_53_control(x, control_keywords_mapping))
    
    return df

def generate_output_excel_modified(df):
    output_df = pd.DataFrame()
    
    output_df['POAM Number'] = range(1, len(df) + 1)
    output_df['ID'] = df['ID']
    output_df['Control'] = df['800-53 Control']
    output_df['Title'] = df['Name']
    output_df['Pub/Mod Date'] = df['Plugin Publication Date']
    output_df['Weakness/Info'] = df['Synopsis'] + ' ' + df['Description']
    output_df['Solution'] = df['Solution']
    output_df['CAT'] = df['CAT']
    output_df['Date Added to POAM'] = ''
    output_df['Status'] = ''
    output_df['Plugin Output'] = df['Plugin Output']
    output_df['Comments'] = ''
    output_df['Mitigation'] = ''
    output_df['Path Forward'] = ''
    output_df['Screenshot'] = ''
    output_df['Ken\'s Comments'] = ''
    
    return output_df

def load_data(file_path):
    return pd.read_excel(file_path)

# File paths
control_catalog_path = prompt_for_file_path("Select the Control Catalog file (e.g., control-catalog.xlsx)")
nessus_path = prompt_for_file_path("Select the Nessus file (e.g., nessus.xlsx)")
output_path = prompt_for_file_path("Select where to save the Output file (e.g., output.xlsx)")

# Load the Control Catalog
control_catalog = pd.read_excel(control_catalog_path)
control_catalog['Combined Text'] = control_catalog['Control (or Control Enhancement) Name'] + ' ' + control_catalog['Control Text']
control_keywords_mapping = {control: extract_keywords(text) for control, text in zip(control_catalog['Control Identifier'], control_catalog['Combined Text'])}

# Example usage
df = load_data(nessus_path)

# Use NER enhancement
nlp = spacy.load("en_core_web_sm")
processed_df = process_data(df, control_keywords_mapping)
final_output_df = generate_output_excel_modified(processed_df)

# Save the Excel
final_output_df.to_excel(output_path, index=False, sheet_name='Report', engine='openpyxl')

# Open the Excel file and adjust column width and row height
book = openpyxl.load_workbook(output_path)
worksheet = book['Report']
for column_cells in worksheet.columns:
    length = max(len(str(cell.value)) for cell in column_cells)
    worksheet.column_dimensions[column_cells[0].column_letter].width = max(length, 10)  # Adjust column width
for row_cells in worksheet.iter_rows():
    for cell in row_cells:
        if cell.row != 1:  # Skip header
            cell.value = cell.value if cell.value != '' else ' '  # Ensure no empty cells
    worksheet.row_dimensions[cell.row].height = 15  # Adjust row height

# Save the formatted Excel
book.save(output_path)