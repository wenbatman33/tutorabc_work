<%
sysinfo=Request.ServerVariables("HTTP_USER_AGENT")


x_sysinfo=split(sysinfo,";")

IF instr(sysinfo,"Windows NT 5.0")> 0 Then
   sysname="Windows 2000"
   softwarenum=1
ElseIF instr(sysinfo,"Windows NT 5.1")> 0 Then 
   sysname="Windows XP"
   softwarenum=1
ElseIF instr(sysinfo,"Windows NT 5.2")> 0 Then 
   sysname="Windows 2003"
   softwarenum=1
ElseIF instr(sysinfo,"Windows NT 6")> 0 Then 
   sysname="Windows Vista"
   softwarenum=3
Else
   sysname=Replace(sysinfo,"")
   softwarenum=2
End IF
%>
