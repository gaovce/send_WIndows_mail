REM '获取IP地址'
REM '判断DNS是否为空，判断IP地址开头是否为10或192'
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=TRUE",,48) 
For Each objItem in colItems 
    If isNULL(objItem.DNSServerSearchOrder) Then
    Else
        IPX=objItem.IPAddress(0) 
        LefIP=split(IPX,".")(0)
        If LefIP="10" OR LefIP="192" Then
           IP=IPX
           Wscript.Echo "ip:" & IP
        End If
    End If
Next
  
REM '获取SN号'
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_BIOS",,48) 
For Each objItem in colItems 
    SN=objItem.SerialNumber
    Wscript.Echo "Sn: " & SN
Next
  
