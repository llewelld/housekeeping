#housekeeping

This python script will do some housekeeping work when marking students' programming assignments (4001COMP):
 1. Change the name of "feedback_username_taskN.docx" to "feedback_'real_username'_taskN.docx"
 2. Add student's name (FirstName, LastName) to the correct table cell in the "feedback_'real_username'_taskN.docx"
 3. *nix only; unzip all zipped files

## Usage

All the submissions should be in a folder named by the marker's initials. The housekeeping.py script should be in the folder containing this folder, alongside the Word feedback template and Excel marking sheet.

Open a console in this folder and run:

housekeeping.py <task> <initials>

Where <task> is the task number (e.g. 1, 2, etc.) and <initials> is the marker's initials (e.g. PH).

## Dependencies:
 1. xlrd; sudo pip install xlrd
 2. python-docx; pip install python-docx

## Installing dependencies on Windows

 1. Install setuptools if it's not already:
   - Follow installation instructions at https://pypi.python.org/pypi/setuptools
   - Add Python\Scripts to PATH (e.g. C:\Python27\Scripts)

 2. xlrd: easy_install xlrd
 3. python-doc: easy_install python-docx
