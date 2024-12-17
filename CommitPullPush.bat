@echo off
rem Dette script syncer med remote repository.

setlocal enabledelayedexpansion

rem Change directory to your repository
rem cd C:\path\to\your\repository

rem User inputs the message for commit message
set /p commitMessage="Enter a commit message (default: WordMat updated): "

if "%commitMessage%"=="" set commitMessage=WordMat updated

rem Git add, commit, pull, and push
cmd /k "git add . && git commit -m "!commitMessage!" && git pull && git push && pause && exit"