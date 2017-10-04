from cx_Freeze import setup, Executable

base = None

executables = [Executable("Def_Email.py", base=base)]

packages = ["win32com", "openpyxl", "datetime", "time", "copy"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "Email Parser",
    options = options,
    version = "1.0",
    description = 'First exe realease of Email Parser.',
    executables = executables
)
