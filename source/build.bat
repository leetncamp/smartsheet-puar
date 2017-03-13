del puar.exe
pyinstaller.exe -F puar.py 
move dist\puar.exe
del /s/q dist
del /s/q build