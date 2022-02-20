<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--
Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")     

%>
<%

%>
<!--內容start-->
<div id="menber_membox">
	<div class='page_title_4'><h2 class='page_title_h2'>教学影片</h2></div>
	<div class="box_shadow" style="background-position:-80px 0px;"></div> 
			<div align="center">
				<!--  --->
				<div class="new_guide_start">
				  <div class="new_quide_test">
				  	<br>					
					<p class='guide_p'>观看新手上路教学影片，让您更了解VIPABC的上课方式及相关功能操作
					<br>* 全营幕观赏：请直接在影片上点击左键两次后开启。</p>
					<!--Video START-->
						<div class="guide-vdotutorials clearfix">
						  <div class="vdo-fction-area">
							<div class="fction-bar"><a href="guide_EnvirVideo1.asp">如何订位</a></div>
							<div class="fction-bar"><a href="guide_EnvirVideo2.asp">进入教室上课</a></div>
							<div class="fction-bar"><a href="guide_EnvirVideo3.asp">如何使用小帮手</a></div>
							<div class="fction-bar-on"><a href="guide_EnvirVideo4.asp">课后评鉴及课后作业</a></div>
							<div class="fction-bar"><a href="guide_EnvirVideo5.asp">其他重要须知</a></div>
						  </div>
						  <div class="vdo-player-area">
						  <iframe src="guide_EnvirVideo4_.asp" width="545" height="410" marginwidth="0" marginheight="0" scrolling="no" frameborder="0" ></iframe>
						  <!--<embed id="ShowVideo" src="/video/envTest/step4.wmv" width="400" height="300" type="application/x-mplayer2" wmode="transparent" showstatusbar="0" autostart="true" EnableContextMenu="false" ></embed>-->
						  
						  </div>
						</div>
						<div class='guide_p'>如有任何疑问，欢迎来信<a href="mailto:vipservices@vipabc.com">vipservices@vipabc.com</a> 或来电客服中心4006-30-30-22询问</div>
				  </div>
				</div>
				<!--  --->
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->