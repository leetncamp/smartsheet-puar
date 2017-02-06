##Paur##

Paur opens an Excel spreadsheet containing data on smartsheet shares and owners.  It emails HP owners of smartsheets that are
shared to non-HP people.

##Windows Development Setup##

1. Install python. Available from http://python.org
2. Install `git` version control software. Available from https://git-scm.com/download/win
3. open a command prompt and change directories to your Desktop
4. execute   `git clone https://github.com/leetncamp/paur.git`  This retrieves the code from it's public repository.
5. Cloning will create a folder called paur. Change directories to paur.
6. execute   `pip install -r requirement.txt`  This installs the necessary python modules to run paur.py
7. To build a windows binary that doesn't depend on a python installation, execute `build.bat`.  See paur.exe. 
