@echo off
rem ---------------------------------------------------------------------------
rem Build the reconstructed reference implementation with the .NET Framework 4.0
rem compiler shipped with Windows. No Visual Studio / NuGet / SDK required.
rem
rem All paths here are ASCII on purpose: .bat files are read with the system ANSI
rem code page, so non-ASCII filenames would be garbled.
rem
rem Output: ..\dist\PsdToPng.rebuilt.exe
rem The original untouched binary in ..\dist\ is left as-is.
rem ---------------------------------------------------------------------------
setlocal
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo [ERROR] .NET Framework 4.x csc.exe not found.
    exit /b 1
)

pushd "%~dp0"
"%CSC%" /nologo /target:winexe /optimize+ /codepage:65001 ^
    /out:"..\dist\PsdToPng.rebuilt.exe" ^
    /r:System.dll /r:System.Core.dll /r:System.Drawing.dll ^
    /r:System.Windows.Forms.dll /r:Microsoft.CSharp.dll ^
    PsdToPng.cs
set "RC=%ERRORLEVEL%"
popd
if not "%RC%"=="0" (
    echo [ERROR] build failed with code %RC%
    exit /b %RC%
)
echo [OK] built ..\dist\PsdToPng.rebuilt.exe
endlocal
