import sys
from cx_Freeze import setup, Executable

main_python_file = "GradeGUI.py"
application_title = "NCKU Grade"
application_description = "App to generate NCKU grade report"

includes = []
excludes = []
packages = []
include_files = []
build_exe_options = {"includes": includes,
                     "excludes": excludes,
                     "packages": packages,
                     "include_files": include_files}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(name=application_title,
      version="0.1",
      description=application_description,
      author="LeeW",
      author_email="cl87654321@gmail.com",
      url="https://github.com/Lee-W/NCKU_Grade",
      options={"build_exe":  build_exe_options},
      executables=[Executable(main_python_file, base=base)]
      )
