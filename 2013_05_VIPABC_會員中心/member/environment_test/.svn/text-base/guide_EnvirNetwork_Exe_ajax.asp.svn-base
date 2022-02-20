<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_global.asp"-->
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp" -->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<%
'MuchTalk會員註冊 
Dim returnValue
Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")

'''''''''''''''''''處理request
''Dim p_email : p_email = getRequest("email" , 1)

'''''''''''''''''''處理request end
Dim envirService 
set envirService = new EnvironmentTestService
returnValue = envirService.saveEnvirNetworkState ( objPCInfo )

objPCInfo.close
set objPCInfo = nothing
set envirService = nothing
%><%=returnValue%>