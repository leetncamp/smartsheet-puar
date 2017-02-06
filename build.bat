c:\Python27\Scripts\pyinstaller.exe -F --data-data fonts --add-binary hp.png paur.py 
move dist\paur.exe
del /s/q dist
del /s/q build