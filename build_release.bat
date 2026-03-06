@echo off
chcp 65001 >nul

echo [1/4] Cleaning old build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [2/4] Building exe...
pyinstaller --noconsole --onefile --name ppt_translator --add-data "phrases.json;." --add-data "lexicon.json;." app.py

echo [3/4] Copying config files...
copy phrases.json dist\ >nul
copy lexicon.json dist\ >nul

echo [4/4] Copying Argos model...
if not exist dist\models mkdir dist\models
copy models\*.argosmodel dist\models\ >nul

echo.
echo Build finished.
echo Release folder: dist
pause
