import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from categorize_expenses import visualize_expenditure_summary

def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filename:
        try:
            print(f"Selected file: {filename}")  # Debugging statement

            # Check file extension to use appropriate engine
            if filename.endswith('.xlsx'):
                df = pd.read_excel(filename, engine='openpyxl')
            elif filename.endswith('.xls'):
                df = pd.read_excel(filename, engine='xlrd')
            else:
                raise ValueError("Unsupported file format. Please select an Excel file with .xlsx or .xls extension.")
            
            # Display the file path in the entry widget
            entry.delete(0, tk.END)
            entry.insert(0, filename)
            
            # Verify the dataframe is read correctly
            print(df.head())
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the Excel file: {e}")
            print(f"Error details: {e}")  # Debugging statement

def generate_summary():
    # Get the input file path
    input_file = entry.get()
    
    if not input_file:
        messagebox.showerror("Error", "Please select an Excel file.")
        return
    
    # Specify the output file path
    output_file = os.path.join(os.path.dirname(input_file), "expenditure_summary.docx")
    
    try:
        print(f"Generating summary for file: {input_file}")  # Debugging statement
        print(f"Output file: {output_file}")  # Debugging statement

        # Call visualize_expenditure_summary with the input and output file paths
        visualize_expenditure_summary(input_file, output_file)
        messagebox.showinfo("Success", f"Expenditure summary generated successfully at {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate summary: {e}")
        print(f"Error details: {e}")  # Debugging statement

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
