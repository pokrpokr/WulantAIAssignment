# WulantAIAssignment
## Project Struct
1. export_excels: The assignment result will output to this folder
2. test_pdfs: The testing pdf should put in this folder
3. assignment1.py or assignment2.py or assignment3.py: Assignment runfile
4. requirements-dev.txt: pypi packages for developing env
## Milestone
1. Assignment1 
    - Finished read pdf text with useing PyMuPDF package
    - Can find text contours of the image of the PDF page with opencv but not combine them
2. Assignment2 
    - Finished read the table of the specific page of the example PDF
3. Assignment3 
    - TODO

## Known Issues
1. Assignment1
    - Footer and some tags are not follow the order of PDF reading order(due to the sort() method logic, can be refactored)
2. Assignment2
    - When using the read_pdf method of camelot, if using lattice as value of the flavor attribute, the table will not be found. However using the stream as value of the flavor attribute, it will read text blocks as table columns.
    - Using camelot to read PDF file page is different with using PyMuPDF.For the requirement that should read the table of page 69 in PDF file,but what I realised it's actually page 77 for the camelot reading result. 
3. Assignment3
    - TODO
## Assignment1
### How to run
```
# args:
# "--src_pdf" or "-sp": PDF file name
# "--expt_name" or "-en": output excel file name
# "--ver" or "-ver": default=1, ver: 1 without using opencv; ver: 2 using opencv(not finished)
# "--pg_start" or "-pgs": default=0, PDF start page
# "--pg_end" or "-pge": default=0, PDF end page
python assignment1.py -sp='keppel-corporation-limited-annual-report-2018.pdf' -en='assignment1.xlsx' -ver=1 -pgs=12 -pge=12
```
## Assignment2
### How to run

```
# args:
# "--src_pdf" or "-sp": PDF file name
# "--expt_name" or "-en": output excel file name
# "--ver" or "-ver": default=1, ver: 1 using camelot;
# "--pg_start" or "-pgs": default=0, PDF start page
# "--pg_end" or "-pge": default=0, PDF end page
# when pg_start equal pg_end and it starts with 0 will mean all pages
python assignment2.py -sp='keppel-corporation-limited-annual-report-2018.pdf' -en='assignment2.xlsx' -ver=1 -pgs=77 -pge=77
```
## Assignment3
### How to run

To Do