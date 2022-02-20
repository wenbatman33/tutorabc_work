<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"-->
<!--#include virtual="/lib/functions/functions_email_to.asp" --><%'發送mail人員%>
<!--#include virtual="/lib/functions/functions_instant_demo.asp" --><%'發送mail人員%>
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--
Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")     

%>
<%
'查詢客戶預約首冊測試時間---結束---

''''''''''''''''''''''''''''''''''''''''''''
'POST
Dim testMicVol : testMicVol=50
Dim testVol : testVol =50
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	'取得麥克風 寫入cookie
	testMicVol = getRequest("testMicVol", CONST_DECODE_NO)
	Response.Cookies("testMicVol").Expires = date - 1
	response.cookies("testMicVol") = testMicVol
	Response.Cookies("testMicVol").Expires = date + 1
	
	testMicVol = getCookies("testMicVol",CONST_DECODE_NO)
	testVol = getCookies("testVol",CONST_DECODE_NO)
	if isEmptyOrNull (testVol) then
		testVol = 0
	end if
	
	if testVol/10=0 or testVol/10=10 then
	else
		testVol =  Cint(testVol/10)+1
	end if
	if testVol>10 then
		testVol = 10
	end if
	
	if testMicVol/10=0 or testMicVol/10=10 then
	else
		testMicVol = Cint(testMicVol/10)+1
	end if
	if testMicVol>10 then
		testMicVol = 10
	end if
	
	sql="SELECT * from fms_client_audio_setting_info  where client_sn='" & intclient_sn & "' and comp_status='abc' "
	x = objPCInfoInsert.excuteSqlStatementEach( sql , Array() , CONST_TUTORMEET_CONN)
	if objPCInfoInsert.RecordCount <= 0  then
		sql="insert into fms_client_audio_setting_info (client_sn, cname, ename, mic_setting, volume_setting, nlevel, profile, comp_status, client_status, lang,class_type) values('" & intclient_sn & "','"&session("client_cname")&"','"&session("client_fname")&"',"&testMicVol&","&testVol&",'"&session("mh_nlevel")&"','首/','abc','首/',2,9) "
		x = objPCInfoInsert.excuteSqlStatementEach( sql , Array() , CONST_TUTORMEET_CONN)
	else
		sql="update fms_client_audio_setting_info set cname='"&session("client_cname")&"', ename='"&session("client_fname")&"',nlevel = '"&session("mh_nlevel")&"', class_type=9 ,   mic_setting="&testMicVol&", volume_setting="&testVol&"  where client_sn=" & intclient_sn & "  and comp_status='abc' "
		x = objPCInfoInsert.excuteSqlStatementEach( sql , Array() , CONST_TUTORMEET_CONN)
	end if
	
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
	' ' ' ' if RoomRS.eof=true then
		' ' ' ' '20090722_Jessica Mod by TutorMeet
		' ' ' ' sql="insert into fms_client_audio_setting_info (client_sn, cname, ename, mic_setting, volume_setting, nlevel, profile, comp_status, client_status, lang,class_type) values('" & intclient_sn & "','"&session("client_cname")&"','"&session("mh_userfirstname")&"',"&testMicVol&","&testVol&",'"&session("mh_nlevel")&"','首/','abc','首/',2,9) "
		' ' ' ' obj_conn_fms.execute sql
	' ' ' ' else
		' ' ' ' '20090722_Jessica Mod by TutorMeet
		' ' ' ' sql="update fms_client_audio_setting_info set cname='"&session("client_cname")&"', ename='"&session("mh_userfirstname")&"',nlevel = '"&session("mh_nlevel")&"', class_type=9 ,   mic_setting="&testMicVol&", volume_setting="&testVol&"  where client_sn=" & intclient_sn & "  and comp_status='abc' "
		' ' ' ' obj_conn_fms.execute sql
	' ' ' ' end if
	' ' ' ' RoomRS.close
	' ' ' ' obj_conn_fms.close
	' ' ' ' set obj_conn_fms = nothing	
	
	
	
	Set envirService = new EnvironmentTestService
	'設定測試成功
	call envirService.setEnvironmentHaveTestOk(objPCInfoInsert)
	'寫測試成功的關懷
	call envirService.testSuccessIMS( objPCInfoInsert )
	
	set logService = new EnvironmentTestServiceLog
	'紀錄首測重測 進入點  Log Insert
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Mic_out" , "Mic_status") , Array( Now() , 1 ) )
	call logService.UpdateTotalStatus( objPCInfoInsert , "" , "" , "" ,"")
	
	'紀錄同時間內，做完首測又做重測，則重測需要新增一筆log紀錄
	call logService.checkFinishIsChangeRepeat( objPCInfoInsert )
	set logService = nothing
end if

objPCInfoInsert.close                               '關閉物件
Set objPCInfoInsert = nothing                       '設定物件為空

'暫時還是N 等到預約完才變Y
session("bolFirstTestQuestionaryhaveTestOk") = "Y"  '與  E:\web\www.vipabc.com\program\member\reservation_class\normal_and_vip.asp 和 E:\web\www.vipabc.com\program\member\reservation_class\include\include_prepare_and_check_data.inc 連動相關
%>
<!--內容start-->
<div id="menber_membox">
			<div align="center">
			<!-- ------------>
			<div class="new_guide_start"><img src="/images/environment_test/title_test_new.jpg" /><br />
			<div class="new_quide_test">
			<div class="step-box"><img src="/images/environment_test/new_test_step4.png"  usemap="#planetmap"></div>
			<%
			Dim firstTest : firstTest =getCookies("firstTest",CONST_DECODE_NO)																
			if firstTest ="N" then
			%>
			<map name="planetmap">
			  <area shape="rect" coords="0,0,140,65" href="guide_EnvirSurvey.asp" alt="填写电脑环境问卷" />
			  <area shape="rect" coords="140,0,280,65" href="guide_EnvirDownload.asp" alt="下载上课软件平台" />
			  <area shape="rect" coords="280,0,420,65" href="guide_EnvirNetwork.asp" alt="网络联机测试" />
			  <area shape="rect" coords="420,0,600,65" href="guide_EnvirVolumn.asp" alt="耳机&麦克风测试" />
			</map>
			<%end if%>
			<div class="opps-box003">
			  <div class="title">麦克风测试完成!</div>
			</div>
			<div class="opps-box000">
				<dl>
					<dd>恭喜您完成麦克风测试!<br/><br/>
					若您尚未完成「DCGS学习偏好设定」或「程度分析」，点选「完成」后系统将自动导页。
					</dd>
					<div class="t-left">
						<input type="button" name="button2" id="button2" value="完成" class="btn-org" onclick="javascript:location.href='/program/member/learning_dcgs.asp'" /> <!--/program/member/member.asp-->
					</div>
				</dl>
			</div>

			</div>
			</div>
			<!-- ------------>                                                                           
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->