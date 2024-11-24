import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

def csv_to_excel():
    # Prompt user to select a CSV file
    Tk().withdraw()  # Hide the root Tkinter window
    csv_file = askopenfilename(filetypes=[("CSV files", "*.csv")], title="Select CSV File")
    
    if not csv_file:
        print("No file selected. Exiting.")
        return
    
    # Read the CSV file into a DataFrame
    try:
        data = pd.read_csv(csv_file)
    except Exception as e:
        print(f"Error reading the CSV file: {e}")
        return
    
    # Prompt user to select the save location for the Excel file
    excel_file = asksaveasfilename(defaultextension=".xlsx", 
                                   filetypes=[("Excel files", "*.xlsx")], 
                                   title="Save Excel File As")
    
    if not excel_file:
        print("No save location selected. Exiting.")
        return
    
    # Save the DataFrame to an Excel file
    try:
        data.to_excel(excel_file, index=False, engine='openpyxl')
        print(f"CSV successfully converted to Excel and saved as: {excel_file}")
    except Exception as e:
        print(f"Error saving the Excel file: {e}")

if __name__ == "__main__":
    csv_to_excel()
