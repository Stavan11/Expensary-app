import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import categorize_expenses

# Function to browse and load Excel file
def browse_file():
    global input_file  # Declare input_file as global

    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filename:
        try:
            print(f"Selected file: {filename}")  # Debugging statement

            # Read the Excel file using pandas
            input_file = filename

            # Display the file path in the entry widget
            entry.delete(0, tk.END)
            entry.insert(0, filename)

            # Enable generate button once file is selected
            generate_button.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the file: {e}")
            print(f"Error details: {e}")  # Debugging statement

# Function to generate summary report
def generate_summary():
    global input_file

    if not input_file:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        # Specify the output file path
        output_file = os.path.join(os.path.dirname(input_file), "expenditure_summary.docx")

        # Placeholder function for actual report generation
        # Call visualize_expenditure_summary with the input and output file paths
        categorize_expenses.categorize_expenditure(input_file, 'Remarks', 'Amount','Deposit Amt.')

        categorize_expenses.generate_report()

        # Open the generated report file
        #os.startfile(output_file)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate summary: {e}")
        print(f"Error details: {e}")  # Debugging statement

# Initialize global variables
input_file = None

# Create the main window
root = tk.Tk()
root.title("Expenditure Summary Generator")

# Create a label and entry for file selection
label = tk.Label(root, text="Select Excel file:")
label.grid(row=0, column=0, padx=10, pady=10)
entry = tk.Entry(root, width=50)
entry.grid(row=0, column=1, padx=10, pady=10)

# Browse button to select Excel file
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# Generate button to create summary report
generate_button = tk.Button(root, text="Generate Summary", command=generate_summary, state=tk.DISABLED)
generate_button.grid(row=1, column=1, padx=10, pady=10)

# Run the main event loop
root.mainloop()
