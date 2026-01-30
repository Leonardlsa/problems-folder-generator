@echo off
setlocal EnableDelayedExpansion
set /p var=移动第_周错题集（数字）:
md .\第%var%周
for /f "delims=" %%i in ('dir /b/s ".\学生"') do (
set a=%%i
set b=!a:~9!
set c=!b:~0,-11!
echo %%i | findstr "第%var%周错题集" &&  md .\第%var%周\!c! 
echo %%i | findstr "第%var%周错题集" && copy %%i .\第%var%周\!c!\
)

pause
