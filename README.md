# send_Windows_mail
Windows下简单的使用VBS自动发送邮件带附件
1.	设置cscript为指定编译器:
    在DOS中执行:  cscript   //h:cscript  //s
2.	在记事本中编写vbs脚本,另存为sendmail.vbs

3. 在记事本中编写bat脚本,另存为sendmail.bat,内容如下:
           Call sendmail.vbs
4. 可用windows定时任务来调用sendmail.bat。

