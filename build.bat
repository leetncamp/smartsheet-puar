c:\Python27\Scripts\pyinstaller.exe paur.py -F --hidden-import=jinja2
move dist\paur.exe
del /s/q dist
del /s/q build