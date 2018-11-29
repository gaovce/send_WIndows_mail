@echo off
echo "正在获取操作系统信息......"
echo The Computer Information: >info.txt
cscript //Nologo Get_Machine.vbs >> info.txt
