@echo off

rem begin localisation of Environment Variables
setlocal enableextensions


rem Uncomment the line below and set required value to MAXIMA_LANG_SUBDIR
rem to get localized describe in command line Maxima
rem set MAXIMA_LANG_SUBDIR=es

rem The variable LANGUAGE is usually not set on Windows, but if Maxima is running in
rem Wine/Linux (e.g. for testing purposes) it is inherited from the Linux environment
rem and causes troubles (at least with Clisp, maybe other Lisps too),
rem so unset this variable.
set LANGUAGE=

set "lisp=sbcl"
set lisp_options=
set "version=5.45.1"

title Maxima Open Source Computer Algebra System %version%

rem For what directory was Maxima configured? Set %prefix% to exactly that directory, if
rem Maxima was installed in the default path (= %configureprefix%/bin/maxima.bat exists).
rem %~dp0/.. should be the same, but is something like the 8.3 short path - and that
rem causes problems with Xmaxima and Clisp.
rem (the combination will not work if Maxima is installed somewhere else, but that is not recommended)
set "prefix=%~dp0/.."
set "configureprefix=C:/maxima-5.45.1"
if exist "%configureprefix%/bin/maxima.bat" set "prefix=%configureprefix%"

rem change "\" to "/" in that path. Works too and causes less troubles (e.g. with xmaxima).
set "prefix=%prefix:\=/%"
set "maxima_prefix=%prefix%"
set "package=maxima"
set verbose=false
set "path=%maxima_prefix%/gnuplot;%maxima_prefix%/gnuplot/bin;%maxima_prefix%/bin;%maxima_prefix%/vtk;%path%"

rem set userdir and tempdir - replace "\" with "/"
if not defined MAXIMA_USERDIR set "MAXIMA_USERDIR=%USERPROFILE:\=/%/maxima"
if not defined MAXIMA_TEMPDIR set "MAXIMA_TEMPDIR=%TEMP:\=/%"
rem there were reports, that the tempdir could not be found (maybe because of special characters, ut8-chars, ...)
rem so set a fallback (as it was before)
if not exist "%MAXIMA_TEMPDIR%"/ set "MAXIMA_TEMPDIR=%USERPROFILE:\=/%"

rem change 'short names' (e.g. "C:/Users/TESTTE~1/AppData/Local/Temp" to
rem long names (e.g. "C:/Users/test test/AppData/Local/Temp" - 
rem CLISP has a problem with these short names.
rem Do this only if maxima_longnames.exe exists in the Maxima binary dir.
rem (it is not compiled and installed for example when a Lisp-only build is done).
if exist "%maxima_prefix%/bin/maxima_longnames.exe" for /F "tokens=* USEBACKQ" %%F in (`maxima_longnames.exe "%MAXIMA_TEMPDIR%"`) do set "MAXIMA_TEMPDIR=%%F"


rem read maximarc and set the Lisp type, if the user did specify MAXIMA_LISP
rem currently no comments or other (Unix shell) commands are understood, only (e.g):
rem LISP=sbcl
if EXIST "%MAXIMA_USERDIR%/maximarc" (
    setlocal enabledelayedexpansion
    for /F usebackq %%L in ("%MAXIMA_USERDIR%/maximarc") do (
        set %%L
        set lisp=!MAXIMA_LISP!
    )
)

rem In order to allow maxima to find out where its files are on MS Windows maxima needs
rem to be started in the directory it can be found in. Which in the case of
rem maxima-portable might be a different folder every time.
rem The user, though, might want it to start in the folder the worksheet lies in, as
rem read_matrix, write_data etc. should be relative to that folder.
rem common lisp doesn't provide such a function and clisp has problems with shortnames
rem while sbcl has problems with non-ascii-non-unicode folder names.
rem The variable MAXIMA_INITIAL_FOLDER tells which folder to start maxima in
rem after searching for the files it is composed of.
if defined MAXIMA_INITIAL_FOLDER cd "%MAXIMA_INITIAL_FOLDER%"

:startparseargs
if x%1 == x-l goto foundlisp
if x%1 == x--lisp goto foundlisp
if x%1 == x-u goto foundversion
if x%1 == x--use-version goto foundversion
if x%1 == x-v goto foundverbose
if x%1 == x--verbose goto foundverbose
if x%1 == x--lisp-options goto foundlispoptions
if x%1 == x-X goto foundlispoptions

:continueparseargs
shift
if not x%1 == x goto startparseargs
goto endparseargs

:foundlispoptions
set lisp_options=%~2
shift
goto continueparseargs

:foundlisp
set lisp=%2
shift
goto continueparseargs

:foundversion
set version=%2
shift
goto continueparseargs

:foundverbose
set verbose=true
goto continueparseargs

:endparseargs

if "%MAXIMA_LAYOUT_AUTOTOOLS%" == "" goto defaultlayout
set layout_autotools=true
goto endlayout

:defaultlayout
set layout_autotools=true

:endlayout

if "%MAXIMA_PREFIX%" == "" goto defaultvars
if "%layout_autotools%" == "true" goto maxim_autotools
set "maxima_imagesdir=%MAXIMA_PREFIX%/src"
goto endsetupvars

:maxim_autotools
set "maxima_imagesdir=%MAXIMA_PREFIX%/lib/%package%/%version%"
goto endsetupvars

:defaultvars
if "%layout_autotools%" == "true" goto defmaxim_autotools
set "maxima_imagesdir=%prefix%/src"
goto endsetupvars

:defmaxim_autotools
set "maxima_imagesdir=%prefix%/lib/%package%/%version%"
goto endsetupvars

:endsetupvars

set "maxima_image_base=%maxima_imagesdir%/binary-%lisp%/maxima"

if "%verbose%" == "true" @echo on
if "%lisp%" == "gcl" goto dogcl
if "%lisp%" == "clisp" goto doclisp
if "%lisp%" == "ecl" goto doecl
rem Allow ccl as an alias of openmcl
if "%lisp%" == "ccl" set lisp=openmcl
if "%lisp%" == "openmcl" goto doopenmcl
if "%lisp%" == "sbcl" goto dosbcl
if "%lisp%" == "abcl" goto doabcl

@echo Maxima error: Lisp %lisp% not known.
goto end

:dogcl
set "path=%maxima_prefix%/lib/gcc-lib/mingw32/undefined;%path"
"%maxima_imagesdir%/binary-gcl/maxima.exe" -eval "(cl-user::run)" %lisp_options% -f -- %*
goto end

:doclisp
if exist "%maxima_imagesdir%/binary-clisp/maxima.exe" goto doclisp_exec

if "%layout_autotools%" == "true" goto clisp_autotools
clisp %lisp_options% -q -M "%maxima_image_base%.mem" "" -- %*
goto end

:clisp_autotools
set "path=%maxima_imagesdir%/binary-clisp;%maxima_prefix%/clisp-2.49/base;%path%"
lisp.exe %lisp_options% -q -M "%maxima_image_base%.mem" "" -- %*
goto end

:clisp_exec
"%maxima_imagesdir%/binary-clisp/maxima.exe" %lisp_options% -q "" -- %*
goto end

:doecl
ecl -load %maxima_image_base%.fas %lisp_options% -eval "(user::run)" -- %*
goto end

rem SBCL Steel Bank Common Lisp
:dosbcl
rem run executable image if it exists
if exist "%maxima_imagesdir%/binary-sbcl/maxima.exe" goto dosbcl_exec
if "%MAXIMA_SIGNALS_THREAD%" == "" (
  set start_maxima="(cl-user::run)"
) else ( 
  set start_maxima="(progn (load (maxima::$sconcat (namestring (pathname (maxima::maxima-getenv \"MAXIMA_PREFIX\"))) \"/bin/win_signals.lisp\")) (cl-user::run))"
)

rem At least in the crosscompiled installer, SBCL_HOME is not defined - set it here
if not defined SBCL_HOME if exist "%maxima_prefix%/bin/sbcl.exe" set "SBCL_HOME=%~dp0"

rem Only on 64bit: if using sbcl (and the user does not specify anything else(!))
rem set the dynamic-space-size to 2GB so that load("lapack") will work.
rem The check for a 64 bit build is done by checking the name of the Mingw libgcc name
rem (which is libgcc_s_seh-1.dll for 64 and libgcc_s_sjlj-1.dll for 32 bit)
if exist "%maxima_prefix%/bin/libgcc_s_seh-1.dll" (
    if "%lisp_options%"=="" set lisp_options=--dynamic-space-size^ 2000
) else (
    rem I am a 32 bit build - the SBCL workaround for load(lapack) will not work
)

sbcl.exe --core "%maxima_imagesdir%/binary-sbcl/maxima.core" --noinform %lisp_options% --end-runtime-options --eval %start_maxima% --end-toplevel-options %*
goto end
:dosbcl_exec
"%maxima_imagesdir%/binary-sbcl/maxima.exe" %lisp_options% --noinform --end-runtime-options --eval "(cl-user::run)" --end-toplevel-options %*
goto end

:doopenmcl
if "%MAXIMA_SIGNALS_THREAD%" == "" (
  set start_maxima="(cl-user::run)"
) else ( 
  set start_maxima="(progn (load (maxima::$sconcat (namestring (pathname (maxima::maxima-getenv \"MAXIMA_PREFIX\"))) \"/bin/win_signals.lisp\")) (cl-user::run))"
)
wx86cl -I "%maxima_image_base%.image" %lisp_options% -e %start_maxima% -- %*
goto end

:doabcl
set "MAXIMA_IMAGESDIR_BIN=%maxima_imagesdir%/binary-abcl"

java -jar "%maxima_imagesdir%/binary-abcl/abcl.jar" %lisp_options% --eval "(load \"%maxima_imagesdir%/binary-abcl/defsystem.lisp\")" --eval "(mk:add-registry-location \"%maxima_imagesdir%/binary-abcl\")" --eval "(funcall (intern (symbol-name :operate-on-system) :mk) \"maxima\" :load :verbose nil)" --eval "(cl-user::run)" -- %*
goto end

:end

rem Restore environment variables
endlocal
