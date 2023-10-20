cd c:\projects\trimark
erase dist /q
.\venv\Scripts\pyinstaller.exe --onefile --icon=trimark.ico --noconfirm --clean main.py


