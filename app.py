import openpyxl as xl
from openpyxl.chart import BarChart, Reference


def process_workbook(filename, sheet_name='Sheet1'):
    """
    Processes an Excel workbook to adjust prices in a specific column
    and adds a bar chart based on the corrected prices.

    Args:
    - filename (str): The path to the Excel file.
    - sheet_name (str): The name of the sheet to process (default is 'Sheet1').
    """
    try:
        wb = xl.load_workbook(filename)
        
        if sheet_name not in wb.sheetnames:
            print(f"Sheet '{sheet_name}' not found in the workbook.")
            return
        
        sheet = wb[sheet_name]

        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row, 3) 
            if isinstance(cell.value, (int, float)):  
                corrected_price = cell.value * 0.9  
                corrected_price_cell = sheet.cell(row, 4) 
                corrected_price_cell.value = corrected_price
            else:
                print(f"Warning: Non-numeric value in row {row}, column 3.")

        values = Reference(
            sheet,
            min_row=2,
            max_row=sheet.max_row,
            min_col=4, 
            max_col=4
        )

        chart = BarChart()
        chart.add_data(values, titles_from_data=False)
        chart.title = "Corrected Prices Chart"
        chart.x_axis.title = "Product"
        chart.y_axis.title = "Corrected Price"
        
       
        sheet.add_chart(chart, 'G2')
        wb.save(filename)
        print(f"Workbook '{filename}' processed successfully and saved.")
    
    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

process_workbook('transactions.xlsx')  