<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_global.asp"-->
<!--#include virtual="/program/class/environment_test/EnvironmentTestLogService.asp" -->
<!--#include virtual="/lib/include/testLog.inc"-->
<%
'影響的欄位名稱
Dim envLogColName:envLogColName = getRequest("envLogColName", CONST_DECODE_NO)
'影響的欄位值
Dim envLogColVal:envLogColVal = getRequest("envLogColVal", CONST_DECODE_NO)
'影響類型 欄位名稱
Dim envLogKindColName:envLogKindColName = getRequest("envLogKindColName", CONST_DECODE_NO)
'影響類型 欄位值
Dim envLogKindColVal:envLogKindColVal = getRequest("envLogKindColVal", CONST_DECODE_NO)

Dim testLogKind:testLogKind = getRequest("testLogKind", CONST_DECODE_NO)

Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")

set logService = new EnvironmentTestServiceLog
if testLogKind = "FTPClickStep0" then
	'紀錄首測重測 進入點  Log Insert
	call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "0" ) )
elseif	testLogKind = "FTPClickStep1" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "1" ) )
elseif	testLogKind = "FTPClickStep2" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "2" ) )
elseif	testLogKind = "FTPClickStep3" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "3" ) )
elseif	testLogKind = "FTPClickStep4" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "4" ) )
	
elseif	testLogKind = "QSB" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("QuestionStep_out" , "QuestionStep_kind") , Array( Now() , "QSB" ) )
elseif	testLogKind = "QRB" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("QuestionStep_out" , "QuestionStep_kind") , Array( Now() , "QRB" ) )
	
elseif	testLogKind = "javaDownload" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("DownloadJAVA_click" ) , Array( Now()  ) )
elseif	testLogKind = "javaFinish" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("DownloadJAVA_finish" ) , Array( Now()  ) )
elseif	testLogKind = "flashDownload" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("DownloadFlash_click" ) , Array( Now()  ) )
elseif	testLogKind = "flashFinish" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("DownloadFlash_finish" ) , Array( Now()  ) )
elseif	testLogKind = "joinNetDownload" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("DownloadJoinet_click" ) , Array( Now()  ) )
elseif	testLogKind = "internetStep_out" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("InternetStep_out" ) , Array( Now()  ) )
elseif  testLogKind = "orderTest" then
	call logService.AddLogBaseArrCol( objPCInfo , Array("OrderTest_in" ) , Array( Now()  ) )
end if

' select case testLogKind
	' case "FTPClickStep0"
		' call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "0" ) )
	' case "FTPClickStep1"
		' call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "1" ) )
	' case "FTPClickStep2"
		' call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "2" ) )
	' case "FTPClickStep3"
		' call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "3" ) )
	' case "FTPClickStep4"
		' call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_click" , "FTP_kind") , Array( Now() , "4" ) )
	' case "QSB"
		' call logService.AddLogBaseArrCol( objPCInfo , Array("QuestionStep_out" , "QuestionStep_kind") , Array( Now() , "QSB" ) )
	' case "QRB"
		' call logService.AddLogBaseArrCol( objPCInfo , Array("QuestionStep_out" , "QuestionStep_kind") , Array( Now() , "QRB" ) )
' end select

objPCInfo.close                               '關閉物件
Set objPCInfo = nothing                       '設定物件為空
%>