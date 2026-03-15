@echo off
pushd "%~dp0.."

call ".venv\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] .venv activation failed
    popd
    pause
    exit /b 1
)

python "src\merge_pptx.py"

if errorlevel 1 (
    echo [FAILED] merge_pptx.py
) else (
    echo [DONE] merge_pptx.py
)

popd
pause
