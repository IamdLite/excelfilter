from cx_Freeze import setup, Executable

# Specify the program name and source file
program_name = "My Excel Filter"
source_file = "excel_filter.py"

# Create the executable
executable = Executable(script=source_file, targetName=program_name)

# Define options for the setup
options = {
    'build_exe': {
        'packages': ['pandas', 'unicodedata', 'tkinter', 'openpyxl'],
        'include_files': []
    }
}

# Setup the application
setup(name=program_name, version='1.0', description='My program', options=options, executables=[executable])
