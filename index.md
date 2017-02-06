##Paur##

Paur opens an Excel spreadsheet containing data on smartsheet shares and owners.  It emails HP owners of smartsheets that are
shared to non-HP people.

This repository contains a Windows binary, paur.exe.  Execute `paur.exe --help` to see full options. In short,
download `UserAccessReport.xlsx` from smartseets into the same folder as paur.exe. Run `paur.exe`

Some things can be modified by changing the content of configuration.py.

<pre>
Subject = u"Smartsheets"  #The subject of the email.
From = u"do-not-reply@nips.cc"  #The from address of the email
default_filename = u"UserAccessReport.xlsx"  #If you dont' supply the name of an excel spreadsheet, this one is opened. 
smartsheet_manager = u"Terrence Gaines <terrence.gaines@hp.com>"  #Currently not used. 
log_format = ".isoformat()"  #A log file is written each time paur.exe is run. This is the datestamp format. Make it HP compatible for logging. 
date_format = u"%Y-%m-%d %H:%M"  #This currently is not used, but if we displayed the modified date of the sheet, this controls the format.
redirect_emails_to = u"terrence.gaines@hp.com"  #Normally an empty string.  Set to your email address to send all output to you. 
stop_after = 1  #Set this to zero to send all emails. 
</pre>


###Using the program###
2. Install `git` version control software. Available from https://git-scm.com/download/win
3. open a command prompt and change directories to your Desktop
5. Cloning will create a folder called paur. Change directories to paur.
8. Copy configuration.py.sample to configuration.py.  Make any modifications to configuration.py you want. 
9. Download UserAccessReport.xlsx from smartsheets to the same folder as paur.exe

##Windows Development Setup##

1. Install python. Available from http://python.org
2. Install `git` version control software. Available from https://git-scm.com/download/win
3. open a command prompt and change directories to your Desktop
4. execute   `git clone https://github.com/leetncamp/paur.git`  This retrieves the code from it's public repository.
5. Cloning will create a folder called paur. Change directories to paur.
6. execute   `pip install -r requirement.txt`  This installs the necessary python modules to run paur.py
7. To build a windows binary that doesn't depend on a python installation, execute `build.bat`.  See paur.exe. 
8. Copy configuration.py.sample to configuration.py.  Make any modifications to configuration.py you want. 
9. Download UserAccessReport.xlsx from smartsheets to the same folder as paur.exe
