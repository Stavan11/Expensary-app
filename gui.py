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

        # Call functions from categorize_expenses module
        categorize_expenses.categorize_expenditure(input_file, 'Remarks', 'Amount','Deposit Amt.')
        categorize_expenses.generate_report()

        messagebox.showinfo("Success", "Summary report generated successfully!")

        # Open the generated report file
        # os.startfile(output_file)  # Uncomment if you want to open the report automatically

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate summary: {e}")
        print(f"Error details: {e}")  # Debugging statement

# Initialize global variables
input_file = None

# Create the main window
root = tk.Tk()
root.title("ExpenSary App Version 1.0")
root.configure(bg="#f0f0f5")  # Set background color to light blue-gray

# Create a title label with custom font and color
title_label = tk.Label(root, text="Welcome to ExpenSary", font=("Arial", 36, "bold"), fg="navy", bg="#f0f0f5")
title_label.grid(row=0, column=0, columnspan=3, pady=20)  # Centered across columns

# Create a version label at the bottom right with smaller font
version_label = tk.Label(root, text="Version 1.0", font=("Arial", 12), fg="gray", bg="#f0f0f5")
version_label.grid(row=3, column=2, sticky="se", padx=10, pady=10)

# Create a label and entry for file selection
label = tk.Label(root, text="Select Excel file:", bg="#f0f0f5")
label.grid(row=1, column=0, padx=10, pady=10)
entry = tk.Entry(root, width=50)
entry.grid(row=1, column=1, padx=10, pady=10)

# Browse button to select Excel file
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=1, column=2, padx=10, pady=10)

# Generate button to create summary report (initially disabled)
generate_button = tk.Button(root, text="Generate Summary", command=generate_summary, state=tk.DISABLED)
generate_button.grid(row=2, column=1, padx=10, pady=10)

# Run the main event loop
root.mainloop()
