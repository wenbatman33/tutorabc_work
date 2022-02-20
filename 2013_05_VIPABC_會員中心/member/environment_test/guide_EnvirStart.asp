<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<script src="/program/member/environment_test/js/logService.js"></script>
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--

Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")

set logService = new EnvironmentTestServiceLog
'紀錄首測重測 進入點  Log Insert
call logService.AddLogBaseArrCol( objPCInfo , Array("FTP_in") , Array( Now() ) )

objPCInfo.close                               '關閉物件
Set objPCInfo = nothing                       '設定物件為空
%>
<!--內容start-->
<div id="menber_membox">
				<div class='page_title_4'><h2 class='page_title_h2'>新手上路</h2></div>
				<div class="box_shadow" style="background-position:-80px 0px;"></div> 
			<div align="center">
				<!--  --->
				<div class="new_guide_start"><br />
				<div class="new_quide_test">
				<h1>亲爱的客户，您好：<br />
				欢迎加入VIPABC！为了确保您的上课品质，请您务必以「日后上课会使用的电脑」，进行以下测试。
				提醒您，未来若更换电脑设备、网络、耳机、麦克风及上课地点或电脑重装，您可随时至「<span class="f-bold">新手上路－测试上课电脑设备</span>」进行测试。</h1>
				<div class="test_bg">
				
				<%
				Dim firstTest : firstTest =getCookies("firstTest",CONST_DECODE_NO)																
				if firstTest ="N" then
				%>
				  <div class="r clearfix">
					<div class="number">Step1</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" style="float:left;width:210px;margin-right:0" href="#">填写电脑环境问卷</span><div class="btn_go" ><a class="FTPStep1" href="guide_EnvirSurvey.asp?">重新测试</a></div></div>
				  </div>
				  <div class="r clearfix">
					<div class="number">Step2</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" style="float:left;width:210px;margin-right:0" href="#">下载上课平台软件</span><div class="btn_go"><a class="FTPStep2" href="guide_EnvirDownload.asp?">重新测试</a></div></div>
				  </div>
				  <div class="r clearfix">
					<div class="number">Step3</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" style="float:left;width:210px;margin-right:0" href="#">网络连接测试</span><div class="btn_go"><a class="FTPStep3" href="guide_EnvirNetwork.asp?">重新测试</a></div></div>
				  </div>
				  <div class="r clearfix">
					<div class="number">Step4</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" style="float:left;width:210px;margin-right:0" href="#">耳机&amp;麦克风测试</span><div class="btn_go"><a class="FTPStep4" href="guide_EnvirVolumn.asp?">重新测试</a></div></div>
				  </div>
				<!--
				<a href="guide_EnvirSurvey.asp?"><img src="/images/environment_test/new_test_start1.png" /></a><br />
				<a href="guide_EnvirDownload.asp?"><img src="/images/environment_test/new_test_start2.png" /></a><br />
				<a href="guide_EnvirNetwork.asp?"><img src="/images/environment_test/new_test_start3.png" /></a><br />
				<a href="guide_EnvirVolumn.asp?"><img src="/images/environment_test/new_test_start4.png" /></a>
				-->
				<%else%>
				  <div class="r clearfix">
					<div class="number">Step1</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" href="#">填写电脑环境问卷</span></div>
				  </div>
				  <div class="r clearfix">
					<div class="number">Step2</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" href="#">下载上课平台软件</span></div>
				  </div>
				  <div class="r clearfix">
					<div class="number">Step3</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" href="#">网络连接测试</span></div>
				  </div>
				  <div class="r clearfix">
					<div class="number">Step4</div>
					<div class="r_r_01" style="line-height:28px;"><span class="title" href="#">耳机&amp;麦克风测试</span></div>
				  </div>
				<!--
				<img src="/images/environment_test/new_test_start1.png" /><br />
				<img src="/images/environment_test/new_test_start2.png" /><br />
				<img src="/images/environment_test/new_test_start3.png" /><br />
				<img src="/images/environment_test/new_test_start4.png" />
				-->
				<%end if%>
				<% if not firstTest ="N" then %>
				<div class="test-sb"><a href="guide_EnvirSurvey.asp?" class="FTPStep0" ><img src="/images/environment_test/new_test_start_btn.png" ></a></div>
				<%end if%>
				</div>
				</div>
				</div>
				<!--  --->
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->