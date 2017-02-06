del paur.exe
pyinstaller.exe -F paur.py 
move dist\paur.exe
del /s/q dist
del /s/q build