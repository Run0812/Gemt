@echo off
set /p a="请输入要汇总的文件夹路径: "
.\env\python.exe src/gather_all_salary.py %a%
pause