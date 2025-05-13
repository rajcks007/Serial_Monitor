to make a .exe file below are the command
pip install pyinstaller
pyinstaller --onefile --windowed --icon=fast.ico --name="FAST-Serial Monitor" AM60_TB_FAST.py

where,  AM60_TB_FAST.py                 is name of pythone file
        --icon=fast.ico                 to set a icon for .exe file
        --name="FAST-Serial Monitor"    to set a name for .exe file