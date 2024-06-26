import cx_Freeze
import sys
from cx_Freeze import setup, Executable

base = None

# Only use the base parameter for Windows systems
if sys.platform == "win32":
    base = "Win32GUI"

# Define the setup parameters
setup(
    name="Expensary",
    version="1.0",
    description="Expenditure Summary Generator",
    options={"build_exe": {"packages": ["tkinter", "os", "pandas", "categorize_expenses"],
                           "include_files": []}},
    executables=[Executable("gui.py", base=base)]
)
