@echo off
setlocal enabledelayedexpansion
for /f "delims=" %%i in ('dir /b ".\情况反馈"') do (
set a=%%i
set b=!a:~0,-5!
copy .\情况反馈\%%i .\学生\!b!
echo !b! &echo Completed.
)
pause
