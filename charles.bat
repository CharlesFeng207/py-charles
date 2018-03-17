@echo off

if exist %~dp0%1.bat (
  call %~dp0%1.bat
   ) ELSE (
   %~dp0python27\python.exe %~dp0%1.py
   )

pause