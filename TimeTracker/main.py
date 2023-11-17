import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import openpyxl

# Function to process the Excel file
def process_excel_file():
    # Prompt user to select Excel file
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not excel_file_path:
        return

    # Prompt user for sheet name and column name
    sheet_name = entry_sheet.get()
    column_name = entry_column.get()

    try:
        # Read the specified sheet and column from the Excel file
        data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        column_data = data[column_name]

        # Check if the column is already in datetime format
        if not pd.api.types.is_datetime64_dtype(column_data):
            # Convert column data to datetime format
            column_data = pd.to_datetime(column_data, errors='coerce')

        # Get today's date
        today = pd.to_datetime(datetime.now().date())

        # Calculate the difference in days between today and column data
        days_left = (column_data - today).dt.days

        # Add the "Days Left" column to the data
        data["Days Left"] = days_left

        # Filter products based on categories and days left
        category_1 = data[(days_left >= 90) & (days_left <= 120)]
        category_2 = data[(days_left >= 0) & (days_left <= 30)]
        category_3 = data[(days_left >= 30) & (days_left <= 90)]
        category_4 = data[(days_left >= 0) & (days_left <= -15)]

        # Prompt user to select output file path
        output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not output_file_path:
            return

        # Create a new Excel file with the categories
        with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
            category_1.to_excel(writer, sheet_name="Category 1", index=False)
            category_2.to_excel(writer, sheet_name="Category 2", index=False)
            category_3.to_excel(writer, sheet_name="Category 3", index=False)
            category_4.to_excel(writer, sheet_name="Category 4", index=False)

            # Adjust the column width for the Expiration Date and Days Left columns
            workbook = writer.book
            for category_num in range(1, 5):
                sheet_name = f"Category {category_num}"
                worksheet = writer.sheets[sheet_name]
                worksheet.column_dimensions['C'].width = 12  # Adjust the column width for Expiration Date
                worksheet.column_dimensions['D'].width = 10  # Adjust the column width for Days Left

        # Display success message
        messagebox.showinfo("Success", f"The results have been saved in {output_file_path}.")
    except Exception as e:
        # Display error message
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create the Tkinter window
window = tk.Tk()
window.title("Excel Processing")
window.geometry("700x600")
window.configure(bg="lightgreen")

# Create the labels and entry fields
label_sheet = tk.Label(window, text="Sheet Name:")
label_sheet.pack()
entry_sheet = tk.Entry(window)
entry_sheet.pack()

label_column = tk.Label(window, text="Column Name:")
label_column.pack()
entry_column = tk.Entry(window)
entry_column.pack()

# Create the Excel file selection button
button_select_file = tk.Button(window, text="Choose File!", command=process_excel_file)
button_select_file.pack()

# Create the process button
button_process = tk.Button(window, text="Proceed", command=process_excel_file)
button_process.pack()

# Start the Tkinter event loop
window.mainloop()
