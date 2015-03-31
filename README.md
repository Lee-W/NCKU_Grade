# NCKU_Grade
Calculate GPA and generate all NCKU grades information to xlsx file

# Requirements
requests, XlsxWriter

```sh
pip3 install -r requirements.txt
```

# Clone
```sh
git clone --recursive https://github.com/Lee-W/NCKU_Grade.git
```
Since, there is a submodule reference to another repo.  
Recursive clone is needed.

# USAGE
```sh
python3 Ncku_grade_crawler.py
```
Then input your student ID and password.  
"Grade Summary.xlsx" will then be generated under current directory.

# Todo
Parse Course Name in English from other site

# AUTHORS
[Lee-W](https://github.com/Lee-W/)

# LICENSE
MIT

