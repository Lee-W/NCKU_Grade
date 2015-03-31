from setuptools import setup

main_python_file = "GradeGUI.py"
application_title = "NCKU Grade"
application_description = "App for generating NCKU grade report"

includes = ["Ncku_grade_crawler.NckuGradeCrawler"]
excludes = []
packages = []
include_files = []
build_exe_options = {"includes": includes,
                     "excludes": excludes,
                     "packages": packages,
                     "include_files": include_files}

setup(app=[main_python_file],
      setup_requires=["py2app"],
      name=application_title,
      version="0.1",
      description=application_description,
      author="LeeW",
      author_email="cl87654321@gmail.com",
      url="https://github.com/Lee-W/NCKU-course-checker",
      options={"build_exe":  build_exe_options},
      )
