from cx_Freeze import setup, Executable
import os

# Define the output directory
output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ExpenSary Version1")

# Define the build options
build_exe_options = {
    "packages": ["tkinter", "os", "pandas", "categorize_expenses"],
    "include_files": [],
    "build_exe": output_dir
}

# Define the base
base = None
if os.name == 'nt':
    base = 'Win32GUI'  # Use 'Win32GUI' to hide the console window (only for Windows)

# Define the executable
executables = [Executable("gui.py", base=base)]

# Setup the cx_Freeze
setup(
    name="Expensary",
    version="1.0",
    description="Expenditure Summary Generator",
    options={"build_exe": build_exe_options},
    executables=executables
)
