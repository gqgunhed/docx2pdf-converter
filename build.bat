@echo off
REM Use "--noconsole/--windowed" to suppress the black console window when building *.py files
REM This is automatically set, when building *.pyw files
REM 
REM --clean optional
REM
REM https://pyinstaller.org/en/stable/usage.html

pyinstaller --onedir --onefile --noconsole --name Word2PDF_converter converter.pyw
