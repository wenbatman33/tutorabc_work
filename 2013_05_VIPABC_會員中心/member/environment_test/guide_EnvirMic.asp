<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--


if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")  
	
	set logService = new EnvironmentTestServiceLog
	'紀錄首測重測 進入點  Log Insert
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Mic_in") , Array( Now() ) )
	set logService = nothing
	
	objPCInfoInsert.close                               '關閉物件
	Set objPCInfoInsert = nothing                       '設定物件為空
end if

'POST
Dim testVol : testVol=50
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	'取得喇叭聲 寫入cookie
	testVol = getRequest("testVol", CONST_DECODE_NO)
	Response.Cookies("testVol").Expires = date - 1
	response.cookies("testVol") = testVol
	Response.Cookies("testVol").Expires = date + 1	
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
	
   document.tutorMeetManager.jsAdjustVolume(volume);
}

function adjustMicVolume(volume) {
//alert("volume:"+volume);
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
	//alert(arr[0]);
	$("#testMicVol").val(arr[0]);
	adjustMicVolume( arr[0]*0.01 );
}
</script>
</head>
    <body>
<applet code="CheckJava" width="0" height="0" mayscript="mayscript" name="checkJava">
</applet>

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
				  <area shape="rect" coords="0,0,140,65" href="guide_EnvirSurvey.asp" alt="填写电脑环境问卷" />
				  <area shape="rect" coords="140,0,280,65" href="guide_EnvirDownload.asp" alt="下载上课软件平台" />
				  <area shape="rect" coords="280,0,420,65" href="guide_EnvirNetwork.asp" alt="网络联机测试" />
				  <area shape="rect" coords="420,0,600,65" href="guide_EnvirVolumn.asp" alt="耳机&麦克风测试" />
				</map>
				<%end if%>
				<!-------------- title --->
				<div class="opps-box003">
				  <div id="title_1" class="title title_tmp">现在进行麦克风测试</div>
				  <div id="title_2" class="title title_tmp" style="display:none">麦克风没有声音</div>
				  <div id="title_3" class="title title_tmp" style="display:none">麦克风还是没有声音</div>
				  <div id="title_4" class="title title_tmp" style="display:none">麦克风有杂音</div>
				  <div id="title_5" class="title title_tmp" style="display:none">麦克风还是有杂音</div>
				  <div id="title_6" class="title title_tmp" style="display:none">麦克风收音太大(小)声</div>
				</div>
				<!-------------- content --->
				<div id="content_1" class="content_tmp" >
					<div class="opps-box000">
					  <dl>
						<dd>请依照下列步骤完成麦克风测试。</dd>
						<dd>
						  <ol>
							<li>请将您的麦克风粉红色接头，接在电脑的麦克风图示插孔。 <div class="test_pic"><img src="/images/environment_test/ico_mic_01.gif" /></div></li>
							<li>请戴上一体成型的耳机麦克风，并将麦克风调整到合适的位置（如图示）。<div class="test_pic"><img src="/images/environment_test/ico_mic_03.gif" /></div></li>
						  </ol>
						</dd>
					  </dl>
					</div>
					<div class="qa">请按下录音键 <img src="/images/environment_test/ico_btn_rec.gif" />后念出以下文字，完成后请点选<img src="/images/environment_test/ico_btn_play.gif" />播放您录制的声音。</div>
					<p class="opps-box004">如果我在通话的同时，指标进入黃色区域，代表麦克风已调整到适当的音量。若指标未进入黃色区域，则需要将麦克风位置移近 。</p>
				</div>
				
				<div id="content_2" class="content_tmp" style="display:none">
					<div class="opps-box000">
					  <dl>
						<!--<dt>麥克風沒聲音</dt>-->
						<dd>请检查麦克风插头位置并重新接上插头</dd>
						<dd>
						  <p><img src="/images/environment_test/ico_mic_01.gif" /></p>
						</dd>
						<dd>
						  <ol>
							<li>请检查麦克风接头位置并重新接上插孔。</li>
							<li>请确认荧幕上的音量调整杆是否停在左边的位置，导致音量太小。请将音量调整杆往右边移动至合适的位置。</li>
						  </ol>
						</dd>
					  </dl>
					</div>
					<div class="qa">请按下录音键<img src="/images/environment_test/ico_btn_rec.gif" />后念出以下文字，完成后请点选<img src="/images/environment_test/ico_btn_play.gif" />播放您录制的声音。</div>
					<p class="opps-box004">如果我在通话的同时，指标进入黃色区域，代表麦克风已调整到适当的音量。<br />若指标未进入黃色区域，则需要将麦克风位置移近。</p>
				</div>
				
				<div id="content_3" class="content_tmp" style="display:none">
					<div class="opps-box000">
					<dl>
						<dd>请确认荧幕上的下拉选单是否有其他麦克风装置，如果有其他的麦克风装置，请选择另一个装置后再次点选「声音测试」。</dd>
						<div class="qa">请按下录音键<img src="/images/environment_test/ico_btn_rec.gif" />后念出以下文字，完成后请点选<img src="/images/environment_test/ico_btn_play.gif" />播放您录制的声音。</div>
					</dl>
					<p class="opps-box004">如果我在通话的同时，指标進入黃色区域，代表麦克风已调整到适当的音量。<br />若指标未进入黃色区域，则需要将麦克风位置移近 。</p>
					</div>
				</div>
				
				<div id="content_4" class="content_tmp" style="display:none">
					<div class="opps-box000">
					  <dl>
						<dd>
						  <p><img src="/images/environment_test/ico_mic_04.gif" /> <img src="/images/environment_test/ico_mic_01.gif" /></p>
						</dd>
						<dd>
						  <ol>
							<li>如果您是使笔记型电脑测试，请务必将电池充满，并移除AC电源线，避免产生电源干扰声。</li>
							<li>维持学习的最佳品质，请确认您的周遭环境是否有杂音干扰，务必在安静的球境中上课。</li>
							<li>请先拔掉麦克风的粉紅色接头，并重新将接头接在电脑的粉红色插孔或有麦克风图示的插孔。</li>
						  </ol>
						</dd>
						<div class="qa">请按下录音键<img src="/images/environment_test/ico_btn_rec.gif" />后念出以下文字，完成后请点选<img src="/images/environment_test/ico_btn_play.gif" />播放您录制的声音。</div>
					  </dl>
					  <p class="opps-box004">如果我在通话的同时，指标进入黃色区域，代表麦克风已调整到适当的音量。<br />若指标未进入黃色区域，则需要将麦克风位置移近 。</p>
					</div>
				</div>
				
				<div id="content_5" class="content_tmp" style="display:none" >
					<div class="opps-box000">
					  <dl>
						<dd>请确认荧幕上的下拉选单是否有其他麦克风装置，如果有其他的麦克风装置，请选择另一个装置后再次点选「声音测试」。</dd>
						<div class="qa">请按下录音键<img src="/images/environment_test/ico_btn_rec.gif" />后念出以下文字，完成后请点选<img src="/images/environment_test/ico_btn_play.gif" />播放您录制的声音。</div>
					  </dl>
						<p class="opps-box004">如果我在通话的同时，指标进入黃色区域，代表麦克风已调整到适当的音量。<br />若指标未进入黃色区域，则需要将麦克风位置移近 。</p>
					</div>
				</div>
				
				<div id="content_6" class="content_tmp" style="display:none">
					<div class="opps-box000">
					  <dl>
						<dd>
						  <p><img src="/images/environment_test/ico_mic_03.gif" /></p>
						</dd>
						<dd>
						  <ol>
							<li>请调整麦克风与您嘴巴的距离，太近或太远都会影响到收音效果。</li>
							<li>请确认下方的音量调整杆，将音量调整到合适的位置。</li>
						  </ol>
						</dd>
					  </dl>
					</div>
					<div class="qa">请按下录音键<img src="/images/environment_test/ico_btn_rec.gif" />后念出以下文字，完成后请选择<img src="/images/environment_test/ico_btn_play.gif" />播放您录制的声音。</div>
					<p class="opps-box004">如果我在通话的同时，指标进入黃色区域，代表麦克风已调整到适当的音量。<br />若指标未进入黃色区域，则需要将麦克风位置移近 。</p>
				</div>
				<!-------------- flash --->
				<div class="player">
					<object id="myContent_ie" classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0" width="453" height="150"  align="middle">
					  <param name="allowScriptAccess" value="always" />
					  <param name="movie" value="recorder.swf?<%=Rnd()%>" />
					  <param name="quality" value="high" />
					  <param name="bgcolor" value="#ffffff" />
					  <param name="wmode" value="transparent" />
					  <embed id="myContent_ff" wmode="transparent" src="recorder.swf?<%=Rnd()%>" mce_src="recorder.swf?<%=Rnd()%>" quality="high" bgcolor="#ffffff" width="453" height="150"  swLiveConnect="true" align="middle" allowScriptAccess="always" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />  
					</object>
				</div>
				
				<form method="POST" action="guide_EnvirMic_ok.asp">
					<input type="hidden" name="testMicVol" id="testMicVol" value="50"/>
				<!-------------- select --->
				<div id="select_1" class="select_tmp" >
					<div class="qa">请问您能清楚听到自己录制的声音吗？</div>
					<div class="t-left">
					  <input type="submit" name="button2" id="button2" value="清楚" class="btn-org"  />
					  <input type="button" name="button2" id="button2" value="沒有声音" class="btn-org" onclick="changeView(2)" />
					</div>
					<div class="qa">若不清楚，是否有遇到什么问题?</div>
					<div class="t-left">
					  <input type="button" name="button2" id="button2" value="有杂音" class="btn-org" onclick="changeView(4)" />
					  <input type="button" name="button2" id="button2" value="收音 太大(小)声" class="btn-org" onclick="changeView(6)" />
					</div>
				</div>
				<div id="select_2" class="select_tmp" style="display:none">
					<div class="qa">请问您能清楚听到自己录制的声音吗？</div>
					<div class="t-left">
					  <input type="submit" name="button2" id="button2" value="清楚" class="btn-org"  />
					  <input type="button" name="button2" id="button2" value="还是没有声音" class="btn-org" onclick="changeView(3)" />
					</div>
					<div class="qa">若不清楚，是否有遇到什么问题？</div>
					<div class="t-left">
					  <input type="button" name="button2" id="button2" value="有杂音" class="btn-org" onclick="changeView(4)" />
					  <input type="button" name="button2" id="button2" value="收音太大(小)声" class="btn-org" onclick="changeView(6)" />
					</div>
				</div>
				<div id="select_3" class="select_tmp" style="display:none">
					<div class="qa">请问您能清楚听到自己录制的声音吗？</div>
					<div class="t-left">
					  <input type="submit" name="button2" id="button2" value="清楚" class="btn-org"  />
					  <input type="button" name="button2" id="button2" value="还是没有声音" class="btn-org" onclick="javascript:location.href='guide_EnvirVolumn_err.asp?err_code=mic1'" />
					</div>
					<div class="qa">若不清楚，是否有遇到什麼問題？</div>
					<div class="t-left">
					  <input type="button" name="button2" id="button2" value="有杂音" class="btn-org" onclick="changeView(4)" />
					  <input type="button" name="button2" id="button2" value="收音太大(小)声" class="btn-org" onclick="changeView(6)" />
					</div>																
				</div>
				
				<div id="select_4" class="select_tmp" style="display:none">
					<div class="qa">请问您能清楚听到自己录制的声音吗</div>
					<div class="t-left">
					  <input type="submit" name="button2" id="button2" value="清楚" class="btn-org"  />
					  <input type="button" name="button2" id="button2" value="还是有杂音" class="btn-org" onclick="changeView(5)" />
					</div>
					<div class="qa">若不清楚，是否有遇到什么问题？</div>
					<div class="t-left">
					  <input type="button" name="button2" id="button2" value="收音太大(小)声" class="btn-org" onclick="changeView(6)" />
					</div>
				</div>
				
				<div id="select_5" class="select_tmp" style="display:none">
					<div class="qa">请问您能清楚听到自己录制的声音吗？</div>
					<div class="t-left">
					  <input type="submit" name="button2" id="button2" value="清楚" class="btn-org"  />
					  <input type="button" name="button2" id="button2" value="还是有杂音" class="btn-org" onclick="javascript:location.href='guide_EnvirVolumn_err.asp?err_code=mic2'" />
					</div>
					<div class="qa">若不清楚，是否有遇到什麼問題？</div>
					<div class="t-left">
					  <input type="button" name="button2" id="button2" value="收音太大(小)声" class="btn-org" onclick="changeView(6)" />
					</div>
				</div>
				
				<div id="select_6" class="select_tmp" style="display:none">
					<div class="qa">请问您能清楚的到自己录制的声音吗？</div>
					<div class="t-left">
					  <input type="submit" name="button2" id="button2" value="清楚" class="btn-org"  />
					  <input type="button" name="button2" id="button2" value="还是太大(小)声" class="btn-org" onclick="javascript:location.href='guide_EnvirVolumn_err.asp?err_code=mic3'" />
					</div>
				</div>
				
				</form>		  
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
				<!--  --->
			</div>
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->