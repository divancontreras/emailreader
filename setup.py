from cx_Freeze import setup, Executable

base = None

executables = [Executable("Email Reader.py", base=base)]

packages = ["win32com", "openpyxl", "datetime", "time", "copy"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "Email Reader",
    options = options,
    version = "1.0",
    description = 'First exe realease of Email Parser, a python scripts that parses data on emails. Author: Diego Contreras.',
    executables = executables
)
