del paur.exe
pyinstaller.exe -F --add-data fonts;fonts --add-data hp.png;hp.png paur.py 
move dist\paur.exe
del /s/q dist
del /s/q build