@echo off
rem This version is optimized for sbcl use with WordMat for speed
rem begin localisation of Environment Variables
setlocal enableextensions
set LANGUAGE=

set "version=5.47.0"

rem Where is Maxima installed? Set %prefix% to exactly that directory.
rem The path is determined by the path of the executed maxima.bat - using the special variable: %~dp0
rem The following command removes the filename and the last path component (bin/maxima.bat),
FOR %%A IN ("%~dp0.") DO SET "prefix=%%~dpA"
rem remove trailing "\":
if "%prefix:~-1%"=="\" set "prefix=%prefix:~0,-1%"

rem change "\" to "/" in that path. Works too and causes less troubles (e.g. with xmaxima).
set "prefix=%prefix:\=/%"
set "maxima_prefix=%prefix%"
set "package=maxima"

rem if not exist "%maxima_prefix%" set maxima_prefix=D:\PROGRA~1\WordMat\MAXIMA~1.1
rem if not exist "%maxima_prefix%" set maxima_prefix=C:\PROGRA~1\WordMat\Maxima-5.47.0
rem if not exist "%maxima_prefix%" set maxima_prefix=C:\PROGRA~2\WordMat\Maxima-5.47.0
rem if not exist "%maxima_prefix%" set maxima_prefix=%ProgramFiles(x86)%\WordMat\Maxima-5.47.0
rem if not exist "%maxima_prefix%" set maxima_prefix=%ProgramFiles%\WordMat\Maxima-5.47.0

set MAXIMA_INFODIR=

set "path=%maxima_prefix%/gnuplot;%maxima_prefix%/gnuplot/bin;%maxima_prefix%/bin;%maxima_prefix%/vtk;%path%"

rem set userdir and tempdir - replace "\" with "/"
if not defined MAXIMA_USERDIR set "MAXIMA_USERDIR=%USERPROFILE:\=/%/maxima"
if not defined MAXIMA_TEMPDIR set "MAXIMA_TEMPDIR=%TEMP:\=/%"
if not exist "%MAXIMA_TEMPDIR%"/ set "MAXIMA_TEMPDIR=%USERPROFILE:\=/%"

rem if exist "%maxima_prefix%/bin/maxima_longnames.exe" for /F "tokens=* USEBACKQ" %%F in (`maxima_longnames.exe "%MAXIMA_TEMPDIR%"`) do set "MAXIMA_TEMPDIR=%%F"

if defined MAXIMA_INITIAL_FOLDER cd "%MAXIMA_INITIAL_FOLDER%"

set "maxima_imagesdir=%MAXIMA_PREFIX%/src"
set "maxima_imagesdir=%MAXIMA_PREFIX%/lib/%package%/%version%"
set "maxima_image_base=%maxima_imagesdir%/binary-sbcl/maxima"

rem sbcl.exe --core "%maxima_image_base%.core" --noinform %lisp_options% --end-runtime-options --eval %start_maxima% --end-toplevel-options %*
sbcl.exe --core "%maxima_image_base%.core" --noinform
