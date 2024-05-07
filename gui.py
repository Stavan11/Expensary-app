import tkinter as tk
import os
from tkinter import filedialog
from categorize_expenses import visualize_expenditure_summary
import pandas as pd

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filename:
        # Try reading the Excel file using different encodings
        for encoding in ['utf-8', 'latin-1', 'cp1252']:
            try:
                # Read the Excel file using pandas with the current encoding
                df = pd.read_excel(filename, encoding=encoding)
                # Display the file path in the entry widget
                entry.delete(0, tk.END)
                entry.insert(0, filename)
                break  # Exit the loop if successful
            except Exception as e:
                continue  # Try the next encoding if reading fails
    


def generate_summary():
    # Get the input file path
    input_file = entry.get()
    
    if not input_file:
        tk.messagebox.showerror("Error", "Please select an Excel file.")
        return
    
    # Specify the output file path
    output_file = os.path.join(os.path.dirname(input_file), "expenditure_summary.docx")
    
    # Call visualize_expenditure_summary with the input and output file paths
    visualize_expenditure_summary(input_file, output_file)


# Create the main window
root = tk.Tk()
root.title("Expenditure Summary Generator")

# Create a label and entry for file selection
label = tk.Label(root, text="Select Excel file:")
label.grid(row=0, column=0, padx=10, pady=10)
entry = tk.Entry(root, width=50)
entry.grid(row=0, column=1, padx=10, pady=10)
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# Create a button to generate the summary
generate_button = tk.Button(root, text="Generate Summary", command=generate_summary)
generate_button.grid(row=1, column=1, padx=10, pady=10)

# Run the main event loop
root.mainloop()
