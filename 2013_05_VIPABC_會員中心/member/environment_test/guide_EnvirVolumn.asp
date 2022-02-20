<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--
Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")   '建立連線 

%>
<%
'本頁功能 新版首測重測 Johnny Liu 2012-08-16
''''''''''''''''''''''''''''''''''''''''''''
'初始化
Dim envirService 
set envirService = new EnvironmentTestService
''''''''''''''''''''''''''''''''''''''''''''
'GET
if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	'紀錄使用者目前進入 狀態 3
	call envirService.insertStatus( objPCInfo , "4" , "program/member/environment_test/guide_EnvirVolumn.asp" )
	
	'紀錄首測重測 進入點  Log Insert
	set logService = new EnvironmentTestServiceLog	
	call logService.AddLogBaseArrCol( objPCInfo , Array("Headset_in","Headset_status" , "Mic_status") , Array( Now(), 4 , 4  ) )
	call logService.UpdateTotalStatus( objPCInfo , "" , "" , "" ,"")
	set logService = nothing
end if
%>
<script language="JavaScript"  type="text/javascript">

function checkJavaInstall() {
    try {
	var x = checkJava.setMessage();	
	mictest.checkJavaInstall("1");
    } catch(e) {
	mictest.checkJavaInstall("0");
    }
}

function showTutorMeet(lang) {
   popupTutorMeet('1', lang);
}

function openVolControl() {
   document.tutorMeetManager.jsOpenVolControl();
}

function adjustVolume(volume) {
//alert("volume:"+volume);
   document.tutorMeetManager.jsAdjustVolume(volume);
}

function adjustMicVolume(volume) {
   document.tutorMeetManager.jsAdjustMicVolume(volume);
}

function getVolumeReady(vol) {
   mictest.getVolumeReady(vol);
}

function getMicVolumeReady(vol) {
   mictest.getMicVolumeReady(vol);
}

function getVolume() { // speaker volumne
   document.tutorMeetManager.jsGetVolume();
}

function getMicVolume() { // microphone volumne
   document.tutorMeetManager.jsGetMicVolume();
}

function reportJavaError(error) {
   try {
       //tutormeet.reportJavaError(error);
   } catch(e) {
   }
}
</script>
<script>
<%'取得flash 回傳回來聲音的值
' [0]  聲音的值
%>
function SoundDateInfo( arr ){
	$("#testVol").val(arr[0]);
	adjustVolume( arr[0]*0.01 );
}
</script>
<applet code="CheckJava" width="0" height="0" mayscript="mayscript" name="checkJava"></applet>
<OBJECT 
    name = "tutorMeetManager"
    classid = "clsid:CAFEEFAC-0016-0000-FFFF-ABCDEFFEDCBA"
    WIDTH = "0" HEIGHT = "0" ALIGN = "baseline" >
    <PARAM NAME = "name" VALUE = "tutorMeetManager" >
    <PARAM NAME = "code" VALUE = "com.tutormeet.remoteaccess.client.RemoteAccessApplet.class" >
    <PARAM NAME = "codebase" VALUE = "" >
    <PARAM NAME = "archive" VALUE = "" >
    <PARAM NAME = "type" VALUE = "application/x-java-applet;version=1.6">
    <PARAM NAME = "scriptable" VALUE = "true">
    <PARAM NAME = "mayscript" VALUE = "yes">    
    <PARAM NAME = "cache_option" VALUE="Plugin">
    <PARAM NAME = "cache_archive" VALUE="tmra.jar">            
</OBJECT>
<!--內容start-->
<div id="menber_membox">
			<div class='page_title_4'><h2 class='page_title_h2'>新手上路</h2></div>
			<div class="box_shadow" style="background-position:-80px 0px;"></div> 
			<div align="center">
			<!-- ------------>
			<form method="POST" action="guide_EnvirVolumn_ok.asp">
			<div class="new_guide_start"><br />
			<div class="new_quide_test">
			<div class="step-box"><img src="/images/environment_test/new_test_step4.png" usemap="#planetmap"></div>
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
			<!-------------- title --->
			<div  class="opps-box003">
			  <div id="title_1" class="title title_tmp">现在进行耳机测试</div>
			  <div id="title_2" class="title title_tmp" style="display:none">我听不到声音</div>
			  <div id="title_3" class="title title_tmp" style="display:none">耳机太大或太小声</div>
			</div>
			<!-------------- content --->
			<div class="opps-box000">
			  <dl>
				<!-------------- content2 --->
				<div id="content_1" class="content_tmp" >
					<dd>您好! 请您拿出耳麦后，将耳机的绿色接头接在电脑的耳机图示插孔中。<br />
					  <div class="test_pic"><img src="/images/environment_test/ico_headphone_01.gif" /></div>
					  接着，请您戴上耳机。<br />
					  <div class="test_pic"><img src="/images/environment_test/ico_mic_02.gif" /></div>
					  并点选播放钮<img src="/images/environment_test/ico_btn_play.gif" />。<br /><br />
					  请确认是否可听到VIPABC的欢迎词?<br />
					  *若音量太小声，您可试着把音量调高一点，再播放一次试试看! <!--<img src="/images/environment_test/vol_icon.jpg" />-->
					</dd>
				</div>
				<div id="content_2" class="content_tmp" style="display:none" >
					<dd>若您无法听到声音，请确认以下的可能原因，进行调整：</dd>
					<dd>
					  <p><img src="/images/environment_test/ico_headphone_01.gif" /> <img src="/images/environment_test/ico_headphone_02.gif" /></p>
					</dd>
					<dd>
					  <ol>
						<li>耳机的绿色接头是否未接在电脑的绿色插孔或画有耳机图示的插孔上。</li>
						<li>您可通过耳机上的音量控制调整适合您的音量大小。</li>
						<li>荧幕上的音量调整杆是否停在左边的位置，导致音量太小声。请将音量调整杆往右边移动至合适的位置。</li>
					  </ol>
					</dd>
					<div class="qa">请按下测试按钮重新测试：</div>
				</div>
				<div id="content_3" class="content_tmp" style="display:none" >
					<dd>请调整耳机音量</dd>
					<dd>
					  <p><img src="/images/environment_test/ico_headphone_02.gif" /></p>
					</dd>
					<dd>
					  <ol>
						<li>请确认耳机上的音量控制，调整到适合您的音量。</li>
						<li>请确认下方的音量调整杆，将音量调整到合适的位置。</li>
					  </ol>
					</dd>
					<div class="qa">请按下测试钮重新测试：</div>
				</div>
					
				<div class="player">
					<object id="myContent_ie" classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0" width="453" height="78"  align="middle">
					  <param name="allowScriptAccess" value="always" />
					  <param name="movie" value="player.swf?<%=Rnd()%>" />
					  <param name="quality" value="high" />
					  <param name="bgcolor" value="#ffffff" />
					  <param name="wmode" value="transparent" />
					  <embed id="myContent_ff"  wmode="transparent" src="player.swf?<%=Rnd()%>" mce_src="player.swf?<%=Rnd()%>" quality="high" bgcolor="#ffffff" width="453" height="78"  swLiveConnect="true" align="middle" allowScriptAccess="always" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />  
					</object>
				</div>
			  </dl>
			</div>
			<!-------------- select --->
			<div id="select_1" class="select_tmp">
				<div class="qa">请问您听的声音清楚吗?</div>
				<div class="t-left">
				  <input type="submit" name="button2" id="button2" value="清楚" class="btn-org" onclick="javascript:location.href=''" />
				</div>
				<div class="qa">还是有遇到什么问题吗?</div>
				<div class="t-left">
				  <input type="button" name="button2" id="button2" value="我听不到声音" class="btn-org" onclick="changeView(2)" />
				  <input type="button" name="button2" id="button2" value="太大声或太小声" class="btn-org" onclick="changeView(3)" />
				</div>
			</div>															
			<div id="select_2" class="select_tmp" style="display:none">
				<div class="qa">请问耳机声音清楚吗？</div>
				<div class="t-left">
				  <input type="submit" name="button2" id="button2" value="听到声音" class="btn-org" onclick="javascript:location.href=''" />
				  <input type="button" name="button2" id="button2" value="还是听不到声音" class="btn-org" onclick="javascript:location.href='guide_EnvirVolumn_err.asp?err_code=vol1'" />
				</div>
				<div class="qa">还是您有遇到其它问题？</div>
				<div class="t-left">
				  <input type="button" name="button2" id="button2" value="太大声或太小声" class="btn-org" onclick="changeView(3)" />
				</div>
			</div>
			<div id="select_3" class="select_tmp" style="display:none">
				<div class="qa">耳机音量适中吗？</div>
				<div class="t-left">
				  <input type="submit" name="button2" id="button2" value="音量适中" class="btn-org" onclick="javascript:location.href=''" />
				  <input type="button" name="button2" id="button2" value="还是太大(小)声" class="btn-org" onclick="javascript:location.href='guide_EnvirVolumn_err.asp?err_code=vol2'" />
				</div>
			</div>
			

			</div>
			</div>
			<input type="hidden" name="testVol" id="testVol" value="50"/>
			</form>
			<!-- ------------>
			<script>
			///about how to display
			function changeView( num){
				$(".title_tmp").hide();
				$(".content_tmp").hide();
				$(".select_tmp").hide();
				$("#title_"+num).fadeIn(2000);
				$("#content_"+num).fadeIn(2000);
				$("#select_"+num).fadeIn(2000);
			}
			
			</script>	                                                        
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->
<%
objPCInfo.close                               '關閉物件
Set objPCInfo = nothing                       '設定物件為空
%>