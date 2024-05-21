import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import landscape
import os

def filter_and_save_to_pdf(input_excel):
    try:
        # Read Excel file
        df = pd.read_excel(input_excel)
        
        # Print all column names
        print("Column names in the Excel file:")
        print(df.columns)
        
        # Check if column 'Issuing bank' exists in the DataFrame
        if 'Issuing bank' not in df.columns:
            raise KeyError("Column 'Issuing bank' not found in the Excel file.")
        
        # Get unique values in column 'Issuing bank'
        unique_values = df['Issuing bank'].unique()
        
        # Create a directory to store PDFs if it doesn't exist
        output_dir = "filtered_pdfs"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Filter DataFrame based on unique values in column 'Issuing bank' and save to PDFs
        for value in unique_values:
            filtered_df = df[df['Issuing bank'] == value]
            
            # Generate PDF filename
            pdf_filename = f"{output_dir}/Column_I_{value}.pdf"
            
            # Generate PDF
            generate_pdf(filtered_df, pdf_filename)
    
    except Exception as e:
        print(f"An error occurred: {e}")



def generate_pdf(df, filename):
    c = canvas.Canvas(filename, pagesize=landscape(letter))
    
    # Define starting position
    x_start = 0.5 * inch  # Left margin
    y_start = 7.5 * inch   # Top margin in landscape mode
    
    # Set column width
    col_width = 1.5 * inch
    
    # Write column headers
    x = x_start
    y = y_start
    for col in df.columns:
        c.drawString(x, y, str(col))
        x += col_width
    y -= 0.25 * inch  # Move to next row
    
    # Write DataFrame content to PDF
    for _, row in df.iterrows():
        x = x_start
        for col in df.columns:
            c.drawString(x, y, str(row[col]))
            x += col_width
        y -= 0.25 * inch  # Move to next row
    
    c.save()
    print(f"PDF saved: {filename}")





# Example usage
input_excel_file = "path to excel file"
filter_and_save_to_pdf(input_excel_file)
