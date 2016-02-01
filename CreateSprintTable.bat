@echo off
setlocal

:: Usage:
:: Create CreateSprintTable.bat Sprint_Name
:: e.g.
::      CreateSprintTable.bat "Sprint 40[FY16]" 
::
:: Output:
::      data.htm
::

if %1 == '' (
    echo Please input sprint name!
    exit 1
    )

:: create json
python.exe GenerateJson.py --config=user.cfg %1

:: create table
python.exe convert.py data.json

