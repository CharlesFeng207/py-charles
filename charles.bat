@echo off

if exist %~dp0cmd\%1.bat (
  call %~dp0cmd\%1.bat %2 %3
   ) else (
   %~dp0python27\python.exe %~dp0cmd\%1.py %2 %3
   )

pause