import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import os

def filter_and_save_to_pdf(input_excel, column_name, header_info):
    try:
        # Read Excel file
        df = pd.read_excel(input_excel)
        
        # Check if the specified column exists in the DataFrame
        if column_name not in df.columns:
            raise KeyError(f"Column '{column_name}' not found in the Excel file.")
        
        # Get unique values in the specified column
        unique_values = df[column_name].unique()
        
        # Create a directory to store PDFs if it doesn't exist
        output_dir = "filtered_pdfs"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Filter DataFrame based on unique values and save to PDFs
        for value in unique_values:
            filtered_df = df[df[column_name] == value]
            
            # Generate PDF filename
            pdf_filename = f"{output_dir}/{column_name}_{value}.pdf"
            
            # Generate PDF
            generate_pdf(filtered_df, pdf_filename, header_info)
    
    except Exception as e:
        print(f"An error occurred: {e}")

def generate_pdf(df, filename, header_info):
    c = canvas.Canvas(filename, pagesize=landscape(letter))
    
    # Define starting position and page layout constants
    x_start = 0.25 * inch  # Left margin
    y_start = 8.0 * inch   # Top margin in landscape mode
    col_width = 1.2 * inch # Width of the columns
    row_height = 0.20 * inch #Height of the rows
    footer_y_position = 0.25 * inch #Footer location & size
    
    # Calculate the number of rows that can fit on one page
    max_rows_per_page = int((y_start - 2 * row_height) / row_height)  # Leave space for header and footer
    
    total_count = 0
    total_sum = 0
    current_row = 0
    total_rows = len(df) #Check the length of the Dataframe rows
    
    # Function to write header
    def write_header():
        c.drawString(x_start, y_start + row_height, header_info)
    
    # Function to write footer
    def write_footer(count, total):
        footer_text = f"Count: {count}   Total: {total:,.2f}"
        c.drawString(x_start, footer_y_position, footer_text)
    
    while current_row < total_rows:
        count = 0
        page_sum = 0
        
        write_header()
        x = x_start
        y = y_start
        
        # Write column headers
        for col in df.columns:
            c.drawString(x, y, str(col))
            x += col_width
        y -= row_height  # Move to next row
        
        # Write DataFrame content to PDF with pagination
        for idx, (_, row) in enumerate(df.iloc[current_row:].iterrows()):
            if idx >= max_rows_per_page:
                break  # Break if reached max rows per page
            x = x_start
            for col in df.columns:
                if col == "Amount":  # Convert 'Amount' column to float without commas
                    cell_value = float(row[col].replace(",", ""))
                else:
                    cell_value = row[col]
                c.drawString(x, y, str(cell_value))
                x += col_width
            y -= row_height  # Move to next row
            count += 1
            page_sum += float(row['Amount'].replace(",", ""))  # Convert 'Amount' column to float without commas
        
        write_footer(count, page_sum)
        c.showPage()  # Start a new page
        
        total_count += count
        total_sum += page_sum
        current_row += max_rows_per_page
    
    # Write the final footer with the total count and sum
    c.drawString(x_start, footer_y_position, f"Total Count: {total_count}   Total Sum: {total_sum:,.2f}")
    
    c.save()
    print(f"PDF saved: {filename}")

# Example usage
input_excel_file = "/Users/apple/Downloads/300489_live_outward_transactions_20240517_162243.xlsx"
column_to_filter = "FI Name"  # Change this to your specific column
header_info = "Business date: 5/19/2024                Transaction report                     BIGPAY"
filter_and_save_to_pdf(input_excel_file, column_to_filter, header_info)
