﻿<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<link href="/css/new_guide_jr.css" rel="stylesheet" type="text/css" />
<script src="/program/member/environment_test/js/logService.js"></script>
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--
Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")     

'紀錄首測重測 進入點  Log Insert
set logService = new EnvironmentTestServiceLog	
call logService.AddLogBaseArrCol( objPCInfoInsert , Array("InternetStep_in") , Array( Now() ) )
set logService = nothing
%>
<%

%>
<!--內容start-->
<div id="menber_membox">
			<div align="center">
				<!-- ------------>
				<div class="new_guide_start">

				<div class='page_title_4'><h2 class='page_title_h2'>测试上课电脑设备</h2></div>
      			<div class="box_shadow" style="background-position:-80px 0px;"></div> 
      			
				<div class="new_quide_test">
				<div class="step-box"><img src="/images/environment_test/new_test_step_jr_3.png" usemap="#planetmap"></div>
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
				
				<div class="opps-box003">请点击『开始测试』，系统将自动为您检测网络连接品质。 </div>
				<div class="t-mid">
				  <input type="submit" name="button2" id="button2" value="开始测试" class="btn-org" onclick="javascript:location.href='guide_EnvirNetwork_Exe.asp'" />
						  </div>
				</div>
				</div>
				<!-- ------------>                                                                            
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->