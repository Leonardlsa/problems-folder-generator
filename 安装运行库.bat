@echo off
::等号中间没有空格
.\Scripts\pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
echo Below are module you're going to install:
for /f %%j in (list.txt) do echo %%j
echo Installation started...
for /f %%i in (list.txt) do .\Scripts\pip install %%i
.\Scripts\pip install -r list.txt
pause
echo Those are your module
.\Scripts\pip list
pause
echo Installation Completed.
pause