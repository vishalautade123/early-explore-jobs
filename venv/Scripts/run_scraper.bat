@echo off
REM Launcher for Task Scheduler. Adjust the path below if your project
REM folder is somewhere other than where this .bat file lives -- by
REM default it just uses its own folder, so you usually don't need to
REM change anything.

cd /d "%~dp0"

REM If you're using a virtual environment, uncomment and adjust:
REM call venv\Scripts\activate.bat

python main11.py

REM Keep the window open briefly on error so you can see what happened
REM when testing manually (Task Scheduler itself won't show this window).
if errorlevel 1 (
    echo.
    echo Script exited with an error - see output\%DATE%\scraper.log
    timeout /t 15
)
