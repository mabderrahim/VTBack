from cx_Freeze import setup, Executable

base = None
target_name = "technical_visit.exe"
icon = "data/aryatowers.ico"
include_files = ['data', 'technical_visits']
packages = ["flask", "flask_restful", "functools", "re", "json", "os", "flask_cors", "passlib"]
source_file = "app.py"

options = {
    'build_exe': {
        'packages': packages,
        'include_files': include_files
    }
}
executables = [Executable(source_file, target_name=target_name, base=base, icon=icon)]

setup(
    name="Mon Programme",
    options=options,
    version="1.0",
    description='Mon programme',
    executables=executables
)