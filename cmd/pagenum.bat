@echo off
pushd "%~dp0.."

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] .venv activation failed
    popd
    pause
    exit /b 1
)

python "src\add_page_numbers.py"

if errorlevel 1 (
    echo [FAILED] add_page_numbers.py
) else (
    echo [DONE] add_page_numbers.py
)

popd
pause
