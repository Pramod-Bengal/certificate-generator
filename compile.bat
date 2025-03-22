@echo off
echo Installing required packages...
pip install -r requirements.txt

echo Creating virtual environment...
python -m venv venv
call venv\Scripts\activate

echo Installing packages in virtual environment...
pip install -r requirements.txt

echo Compiling the application...
pyinstaller --clean generate.spec

echo Compilation complete!
echo The executable can be found in the 'dist' folder.
pause 