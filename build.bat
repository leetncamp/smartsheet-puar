c:\Python27\Scripts\pyinstaller.exe -F --add-data fonts --add-data hp.png paur.py 
move dist\paur.exe
del /s/q dist
del /s/q build