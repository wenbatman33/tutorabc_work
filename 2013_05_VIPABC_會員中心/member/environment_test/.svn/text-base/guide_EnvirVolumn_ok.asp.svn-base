<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數


%>
<%
''''''''''''''''''''''''''''''''''''''''''''
'POST
Dim testVol : testVol=50
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	'取得喇叭聲 寫入cookie
	testVol = getRequest("testVol", CONST_DECODE_NO)
	Response.Cookies("testVol").Expires = date - 1
	response.cookies("testVol") = testVol
	Response.Cookies("testVol").Expires = date + 1	
		
	Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	'紀錄首測重測 進入點  Log Insert
	set logService = new EnvironmentTestServiceLog		
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Headset_out" , "Headset_status" ) , Array( Now() , 1 ) )
	set logService = nothing
	objPCInfoInsert.close                               '關閉物件
	Set objPCInfoInsert = nothing                       '設定物件為空
end if
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
			<div class="opps-box003">
			  <div class="title">耳机测试完成!</div>
			</div>
			<div class="opps-box000">
			  <dl>
				<dd>恭喜您完成耳机测试!<br/>接下来请进行麦克风测试。</dd>
				<div class="t-left">
				  <input type="button" style="width:7em" name="button2" id="button2" value="下一步" class="btn-org" onclick="javascript:location.href='guide_EnvirMic.asp'" />
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