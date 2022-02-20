<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"-->
<!--#include virtual="/lib/functions/functions_email_to.asp" --><%'發送mail人員%>
<!--#include virtual="/lib/functions/functions_instant_demo.asp" --><%'發送mail人員%>
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<script src="/program/member/environment_test/js/logService.js"></script>
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--
Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")     

'紀錄首測重測 進入點  Log Insert	
set logService = new EnvironmentTestServiceLog
''  vol1:聽不到聲音  vol2:太大(小)聲  mic1:聽不到聲音  mic2:有雜音  mic3:太大(小)聲
Dim err_code : err_code = getRequest("err_code", CONST_DECODE_NO)
Set envirService = new EnvironmentTestService
if err_code="vol1" then
	call envirService.saveEnvirVolErrReason ( objPCInfoInsert , "未完成 / 聽不到聲音")
	call envirService.saveEnvirMicErrReason ( objPCInfoInsert , "未完成")
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Headset_out" , "volReason" , "Headset_status") , Array( Now() ,  "未完成 / 聽不到聲音" , 2 ) )
	call logService.UpdateTotalStatus( objPCInfoInsert , "" , "" , "" ,"")
elseif err_code="vol2" then
	call envirService.saveEnvirVolErrReason ( objPCInfoInsert , "未完成 / 太大(小)聲")
	call envirService.saveEnvirMicErrReason ( objPCInfoInsert , "未完成")
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Headset_out" , "volReason" , "Headset_status" ) , Array( Now() ,  "未完成 / 太大(小)聲" , 2 ) )
	call logService.UpdateTotalStatus( objPCInfoInsert , "" , "" , "" ,"")
elseif err_code="mic1" then
	call envirService.saveEnvirMicErrReason ( objPCInfoInsert , "未完成 / 聽不到聲音")
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Mic_out" , "micReason" , "Mic_status" ) , Array( Now() ,  "未完成 / 聽不到聲音" , 2) )
	call logService.UpdateTotalStatus( objPCInfoInsert , "" , "" , "" ,"")
elseif err_code="mic2" then
	call envirService.saveEnvirMicErrReason ( objPCInfoInsert , "未完成 / 有雜音")
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Mic_out" , "micReason" , "Mic_status"  ) , Array( Now() ,  "未完成 / 有雜音" , 2 ) )
	call logService.UpdateTotalStatus( objPCInfoInsert , "" , "" , "" ,"")
elseif err_code="mic3" then
	call envirService.saveEnvirMicErrReason ( objPCInfoInsert , "未完成 / 太大(小)聲")
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Mic_out" , "micReason" , "Mic_status"  ) , Array( Now() ,  "未完成 / 太大(小)聲" , 2 ) )
	call logService.UpdateTotalStatus( objPCInfoInsert , "" , "" , "" ,"")
end if
set logService = nothing
%>
<%
'================是否新增tutormeet table
sql="SELECT * from fms_client_audio_setting_info  where client_sn='" & intclient_sn & "' and comp_status='abc' "
x = objPCInfoInsert.excuteSqlStatementEach( sql , Array() , CONST_TUTORMEET_CONN)
if objPCInfoInsert.RecordCount <= 0 and not isEmptyOrNull(intclient_sn) then
	sql="insert into fms_client_audio_setting_info (client_sn, cname, ename, nlevel, profile, comp_status, client_status, lang,class_type) values('" & intclient_sn & "','"&session("client_cname")&"','"&session("client_fname")&"','"&session("client_nlevel")&"','首/','vip','首/',2,9) "
	x = objPCInfoInsert.excuteSqlStatementEach( sql , Array() , CONST_TUTORMEET_CONN)
end if

call envirService.testFailIMS(objPCInfoInsert)
set envirService = nothing
objPCInfoInsert.close                               '關閉物件
Set objPCInfoInsert = nothing                       '設定物件為空
' ' ' ' Set obj_conn_fms = Server.CreateObject("ADODB.Connection")
' ' ' ' 'set rs2=server.createobject("adodb.recordset")
' ' ' ' obj_conn_fms.connectionString = tutormeet_connectionString			
' ' ' ' obj_conn_fms.open	

' ' ' ' 'strUpdateMeterial = " UPDATE fms_client_audio_setting_info SET volume_setting = "&testVol&" , mic_setting="&testMicVol&" WHERE sn =(select top 1 sn from fms_client_audio_setting_info with(nolock) where client_sn= "&intclient_sn&" order by sn DESC)"
' ' ' ' 'obj_conn_fms.Execute strUpdateMeterial



' ' ' ' set RoomRS=Server.CreateObject("Adodb.RecordSet")
' ' ' ' '檢查該User的資料是否轉入TutorMeet，若否，則直接insert進一筆
' ' ' ' sql="SELECT * from fms_client_audio_setting_info  where client_sn='" & intclient_sn & "' and comp_status='abc' "

' ' ' ' RoomRS.open sql,obj_conn_fms,1,1
' ' ' ' if RoomRS.eof=true and not isEmptyOrNull(intclient_sn) then
	' ' ' ' '20090722_Jessica Mod by TutorMeet
	' ' ' ' sql="insert into fms_client_audio_setting_info (client_sn, cname, ename, nlevel, profile, comp_status, client_status, lang,class_type) values('" & intclient_sn & "','"&session("client_cname")&"','"&session("mh_userfirstname")&"','"&session("client_nlevel")&"','首/','vip','首/',2,9) "
	' ' ' ' obj_conn_fms.execute sql
' ' ' ' else
	' ' ' ' '20090722_Jessica Mod by TutorMeet
	' ' ' ' ' sql="update fms_client_audio_setting_info set cname='"&session("client_cname")&"', ename='"&session("mh_userfirstname")&"',nlevel = '"&session("mh_nlevel")&"', class_type=9 where client_sn=" & intclient_sn & "  and comp_status='abc' "
	' ' ' ' ' obj_conn_fms.execute sql
' ' ' ' end if
' ' ' ' RoomRS.close
' ' ' ' obj_conn_fms.close
' ' ' ' set obj_conn_fms = nothing	

session("bolFirstTestQuestionaryhaveTestOk") = "E"  '與  E:\web\www.vipabc.com\program\member\reservation_class\normal_and_vip.asp 和 E:\web\www.vipabc.com\program\member\reservation_class\include\include_prepare_and_check_data.inc 連動相關
'session("bolFirstTestQuestionaryPassed") = "1" 
%>
<!--內容start-->
<div id="menber_membox">
			<div align="center">
			<!-- ------------>
			<div class="new_guide_start"><img src="/images/environment_test/title_test_new.jpg" /><br />
			<div class="new_quide_test">
			<div class="step-box"><img src="/images/environment_test/new_test_step4.png" usemap="#planetmap"></div>
			<%
			Dim firstTest : firstTest =getCookies("firstTest",CONST_DECODE_NO)																
			if firstTest ="N" then
			%>
			<map name="planetmap">
			  <area shape="rect" coords="0,0,140,65" href="guide_EnvirSurvey.asp?" alt="填写电脑环境问卷" />
			  <area shape="rect" coords="140,0,280,65" href="guide_EnvirDownload.asp?" alt="下载上课软件平台" />
			  <area shape="rect" coords="280,0,420,65" href="guide_EnvirNetwork.asp?" alt="网络联机测试" />
			  <area shape="rect" coords="420,0,600,65" href="guide_EnvirVolumn.asp?" alt="耳机&麦克风测试" />
			</map>
			<%end if%>
			<div class="opps-box001"> 抱歉!测试未通过，建议您重新测试或在线预约测试。</div>
			<div class="opps-box000"> 
			  <span style="font-weight:border">为什么要线上预约测试呢?</span><br/><br/>
			  为确保您拥有良好的上课品质，我们希望为您测试耳机&麦克风，并再次检查电脑、网络连接相关设定是否良好，故建议您现在可直接在线预约测试，我们将安排专人主动于预约时间致电给您，为您服务! 
			  <div class="t-left">
				<input type="button" onclick="javascript:location.href='guide_EnvirVolumn.asp'" class="btn-org" value="耳机&amp;麦克风测试" id="button2" name="button2">				
				<%'<input type="button" onclick="javascript:location.href='/program/member/first_session.asp'" class="btn-org" value="我要预约测试" id="button2" name="button2">%>
				<input type="button" onclick="javascript:location.href='/program/member/environment_test/guide-2firstsession.asp'" class="orderTest btn-org" value="我要预约测试" id="button2" name="button2">
			  </div>
			</div>

			</div>
			</div>
			<!-- ------------>                                                                         
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->