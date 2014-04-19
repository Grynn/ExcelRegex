@echo off
setlocal
set MSBUILD="C:\Windows\Microsoft.NET\Framework\v4.0.30319\msbuild.exe"


%MSBUILD% libExcelRegex.sln /property:Configuration=Release /verbosity:q /nologo
if %errorlevel% neq 0 exit /b %errorlevel%

set CD=%~dp0%
ExcelDna\Distribution\ExcelDnaPack.exe ExcelRegex.dna /Y /O %CD%ExcelRegex.xll
endlocal