# WulantAIAssignment
## Project Struct
1. export_excels: The assignment result will output to this folder
2. test_pdfs: The testing pdf should put in this folder
3. assignment1.py/assignment2.py/assignment3.py: Assignment runfile
4. requirements-dev.txt: pypi packages for developing env
## Assignment1
### Version1
command:
```
# args:
# "--src_pdf" or "-sp": PDF file name
# "--expt_name" or "-en": output excel file name
# "--ver" or "-ver": default=1, ver: 1 without using opencv; ver: 2 using opencv(not finished)
# "--pg_start" or "-pgs": default=0, PDF start page
# "--pg_end" or "-pge": default=0, PDF end page
python assignment1.py -sp='keppel-corporation-limited-annual-report-2018.pdf' -en='test1.xlsx' -ver=1 -pgs=12 -pge=12
```



