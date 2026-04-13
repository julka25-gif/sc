import ocrmypdf
import camelot
import pandas as pd
import os
import argparse

# Define command line arguments
def parse_args():
    parser = argparse.ArgumentParser(description='Combine scanned statement PDFs into an Excel file.')
    parser.add_argument('--input', type=str, default='.', help='Input folder containing PDF files')
    return parser.parse_args()

# Function to extract data from PDFs
def extract_data(input_folder):
    combined_data = []
    pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        cache_pdf_path = os.path.join(input_folder, 'cache', pdf_file)
        # Perform OCR on the PDF
ocrmypdf.ocr(pdf_path, cache_pdf_path)
        # Extract tables using camelot
        tables = camelot.read_pdf(cache_pdf_path, pages='all', flavor='stream')
        for i, table in enumerate(tables):
            df = table.df
            if df.iloc[0].equals(df.iloc[1]):  # Check for repeated header
                df = df[1:]  # Drop the first header
            df.columns = df.iloc[0]  # Normalize header
            df = df[1:]  # Drop the header row
            df['source_file'] = pdf_file
            df['table_index'] = i
            combined_data.append(df)
    return pd.concat(combined_data, ignore_index=True)

# Save the combined DataFrame to Excel
def save_to_excel(df):
    df.to_excel('combined.xlsx', index=False, sheet_name='combined')

# Main function
if __name__ == '__main__':
    args = parse_args()
    os.makedirs(os.path.join(args.input, 'cache'), exist_ok=True)
    combined_df = extract_data(args.input)
    save_to_excel(combined_df)
