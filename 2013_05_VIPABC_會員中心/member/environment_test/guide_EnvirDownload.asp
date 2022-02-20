<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<script src="/program/member/environment_test/js/logService.js"></script>
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--
Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")     
Dim envirService
set envirService = new EnvironmentTestService
if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	'紀錄使用者目前進入 狀態
	call envirService.insertStatus( objPCInfoInsert , "2" , "program/member/environment_test/guide_EnvirDownload.asp" )
	
	'紀錄首測重測 進入點  Log Insert
	set logService = new EnvironmentTestServiceLog	
	call logService.AddLogBaseArrCol( objPCInfoInsert , Array("Downloadstep_in") , Array( Now() ) )
	set logService = nothing
end if
'POST
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	call envirService.saveEnvirDownload( objPCInfoInsert  )

	response.redirect "guide_EnvirNetwork.asp"
end if
%>
<%
'本頁功能 新版首測重測 Johnny Liu 2012-08-16
'''''''''''''''''''''''''''''''''''''''''''''''
' 偵測作業系統，若為MAC電腦則畫面隱藏JoinNet的下載按鈕並直接將喜好教室種類設為TutorMeetOnly，設定方法為先偵測使用者上一次的設定，
' 若已為TutorMeetOnly則不寫入資料，反之則INSERT一筆資料
if Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"mac")>0  or  Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"Mac")>0 then 
	Dim roomType '房間類型
	sql = " SELECT cdc_specific_category_sn  FROM dcgs_client_disagreeable_condition_bas(nolock) WHERE (cdc_category_sn = 50) AND client_sn=@clientSn AND cdc_valid = 1 "
	intSqlWriteResult =objPCInfoInsert.excuteSqlStatementEach(sql, Array(intclient_sn), CONST_TUTORABC_R_CONN) 
	if not objPCInfoInsert.EOF then 
		roomType = objPCInfoInsert("cdc_specific_category_sn")
	end if
	
	if roomType<>2  then  '--TutorMeetOnly : 2
		sql = " UPDATE dcgs_client_disagreeable_condition_bas SET cdc_valid = 0,cdc_edate=GETDATE(),cdc_editor=2,cdc_editor_sn=@clientSn  WHERE client_sn = @clientSn2 AND cdc_valid = 1 "
		intSqlWriteResult =objPCInfoInsert.excuteSqlStatementEach(sql, Array(intclient_sn , intclient_sn), CONST_TUTORABC_R_CONN) 
		
		sql = " INSERT INTO dcgs_client_disagreeable_condition_bas (client_sn ,cdc_category_sn ,cdc_specific_category_sn ,cdc_sdate ,cdc_edate ,cdc_inpdate ,cdc_inputer ,cdc_inputer_sn ,cdc_edtdate ,cdc_editor ,cdc_editor_sn ,cdc_valid ,cdc_learning_note) VALUES(@clientSn,50,2,getdate(),DATEADD(YEAR,20,getdate()),GETDATE(),2,@clientSn2,NULL,NULL,NULL,1,NULL) "
		intSqlWriteResult =objPCInfoInsert.excuteSqlStatementEach(sql, Array(intclient_sn , intclient_sn), CONST_TUTORABC_R_CONN) 
	end if
	
end if

'取得 首測重測資料
call envirService.getEnvironmentTable(objPCInfoInsert)
Dim installJava , installFlash , installJoinNet
if not objPCInfoInsert.EOF then
	installJava = objPCInfoInsert("installJava")
	installFlash = objPCInfoInsert("installFlash")
	installJoinNet = objPCInfoInsert("installJoinNet")
end if
'如果是mac 則預設資料庫已裝
if Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"mac")>0  or  Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"Mac")>0 then 
	installJoinNet = "Y"
end if

objPCInfoInsert.close                               '關閉物件
Set objPCInfoInsert = nothing  
%>
<script type="text/javascript" src="/lib/javascript/deployJava.js"></script>
<script type="text/javascript" src="/lib/javascript/swfobject2.2.js"></script>
<script>
var playerVersion = swfobject.getFlashPlayerVersion();
//support flash
window.isSupportFlash = false;
var MM_contentVersion = 7;
var plugin = (navigator.mimeTypes && navigator.mimeTypes["application/x-shockwave-flash"]) ?
	    navigator.mimeTypes["application/x-shockwave-flash"].enabledPlugin : 0;
if (plugin) {
    var words = navigator.plugins["Shockwave Flash"].description.split(" ");
    for (var i = 0; i < words.length; ++i) {
        if (isNaN(parseInt(words[i])))
            continue;
        var MM_PluginVersion = words[i];
    }
    window.isSupportFlash = MM_PluginVersion >= MM_contentVersion;
} else if (navigator.userAgent && navigator.userAgent.indexOf("MSIE") >= 0 &&
			(navigator.appVersion.indexOf("Win") != -1)) {
    
        window.isSupportFlash = new ActiveXObject("ShockwaveFlash.ShockwaveFlash." + MM_contentVersion) != null;
    
}


</script>
<!--內容start-->
<div id="menber_membox">
			<div align="center">
				<!-- -------------->
				<div class="new_guide_start">


				<div class='page_title_4'><h2 class='page_title_h2'>测试上课电脑设备</h2></div>
      			<div class="box_shadow" style="background-position:-80px 0px;"></div> 


				<div class="new_quide_test">
				<div class="step-box"><img src="/images/environment_test/new_test_step2.png" usemap="#planetmap"></div>
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
				<div class="opps-box001" id="not_download" style="display:none">提醒您，您尚未安装上课所需要的软件，请点选『立即安装』。</div>
				<div class="opps-box002" id="is_download" style="display:none">恭喜您！您已经成功下载上课软件！</div>
				<div style="padding-top:10px">
				<table width="100%" border="0" cellspacing="30" cellpadding="0" class="">
				  <tr>
					<!--用ico_software-on和ico_software-off切換圖示-->
                    <%  '寫死下載位置……判定userAgent給予不同的下載版本
		                Dim userAgent : userAgent = Request.ServerVariables("HTTP_USER_AGENT") 
		                if (instr(userAgent, "NT 5.2")>0) then '判斷NT 5.2有出現在userAgent中，有出現則值代表為位置（Windows XP 64位元為Windows NT 5.2的架構）
                            userAgentWinFlash = "http://152.101.38.107/download/install_flash_player_1102_64bit.exe"
		                elseif (instr(userAgent, "NT 5.1")>0)then 'Windows XP 32位元的架構是NT5.1
                            userAgentWinFlash = "http://152.101.38.107/download/install_flash_player_1102_32bit.exe"
		                else'走正常路線
		                    userAgentWinFlash = "http://get.adobe.com/cn/flashplayer/"
		                end if
	                %>
					<td style="height:100px" valign="top" width="33%" align="center"><div id="java_id" class="ico_software-on software-java" onclick="window.open('http://www.java.com/download')" style="cursor:pointer">&nbsp;</div>
					<input style="display:none" type="button" name="java_download" id="java_download" value="立即安装" class="btn-org javaDownload" onclick="window.open('http://www.java.com/download')"   />
					<div id="java_download_ok" style="display:none" class="sw_done" style="display:block"><img src="/images/environment_test/guide_icon_line_01.gif" /> 完成</div>
					</td>
					<%if Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"mac")>0  or  Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"Mac")>0 then %>
					<td valign="top" width="33%" align="center"><div id="flash_id" class="ico_software-on software-flesh" class="btn-org flashDownload" onclick="window.open('/download/install_flash_player_osx.dmg')" style="cursor:pointer">&nbsp;</div><!-- http://get.adobe.com/tw/flashplayer -->
					<%else%>
					<td valign="top" width="33%" align="center"><div id="flash_id" class="ico_software-on software-flesh" class="btn-org flashDownload" onclick="window.open('<%=userAgentWinFlash%>')" style="cursor:pointer">&nbsp;</div><!-- http://get.adobe.com/tw/flashplayer -->
					<%end if%>
					<%if Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"mac")>0  or  Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"Mac")>0 then %>
					<input style="display:none" type="button" name="flash_download" id="flash_download" value="立即安装" onclick="window.open('/download/install_flash_player_osx.dmg')" class="flashDownload" />
					<%else%>
					<input style="display:none" type="button" name="flash_download" id="flash_download" value="立即安装" onclick="window.open('<%=userAgentWinFlash%>')" class="flashDownload" />
					<%end if%>
					<div id="flash_download_ok" style="display:none" class="sw_done" style="display:block"><img src="/images/environment_test/guide_icon_line_01.gif" /> 完成</div>
					</td>
					
					<!--#include virtual="/lib/functions/functions_checkcompsystem.asp" --><%'判斷使用者作業系統%>
					 <%
						'不同的作業系統要下載不同的下載連線軟體
						IF softwarenum=1 Then
							strDownloadSoftware="client_xp_software_usetcp.exe"
						ElseIF softwarenum=2 Then
							strDownloadSoftware="client_9x_software_usetcp.exe"
						ElseIF softwarenum=3 Then
							strDownloadSoftware="client_vista_software_usetcp.exe"
						END IF
					  %>
					
					<%if Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"mac")<=0 and Instr( Request.ServerVariables("HTTP_USER_AGENT")  ,"Mac")<=0 then %>
					<td valign="top" width="33%" align="center"><div class="ico_software-on software-jnet" onclick="joinNetInstalled();window.open('http://<%=CONST_VIPABC_WEBHOST_NAME%>/download/<%=strDownloadSoftware%>')" style="cursor:pointer">&nbsp;</div>
						<%if installJoinNet="Y" then%>
							<div class="sw_done" style="display:block"><img src="/images/environment_test/guide_icon_line_01.gif" /> 完成</div>
						<%else%>
							<input type="button" name="joinNet_download" id="joinNet_download" value="立即安装" class="btn-org joinNetDownload" onclick="joinNetInstalled();window.open('http://<%=CONST_VIPABC_WEBHOST_NAME%>/download/<%=strDownloadSoftware%>')"  />
							<div name="joinNet_download_ok" id="joinNet_download_ok" class="sw_done" style="display:none"><img src="/images/environment_test/guide_icon_line_01.gif" /> 完成</div>
						<%end if%>
					</td>
					<%end if%>
				  </tr>
				</table>
				</div>
				<form name="formX" action="guide_EnvirDownload.asp?" method="post">
				<div class="t-mid">
					<input style="display:none" id="download_pass_btn" type="submit" name="button2" id="button2" value="下一步" class="btn-org"  />
					<input type="hidden" name="java_ver" id="java_ver" />
					<input type="hidden" name="flash_ver" id="flash_ver" />
					<input type="hidden" name="joinNet_ver" id="joinNet_ver" value="<%=installJoinNet%>" />
					<input type="hidden" name="browser" id="browser" value=""/>
					<input type="hidden" name="os" id="os" value=""/>
				</div>
				</form>
				<script>
					function joinNetInstalled(){
						$("#joinNet_download").hide();
						$("#joinNet_download_ok").show();
						$("#joinNet_ver").val("Y");
						//$("#is_download").show();
						//$("#download_pass_btn").show();
						//$("#not_download").hide();
					}
					
					var sh;
					var flashLogHaveSend = 0; //已寫flash的log
					var javaLogHaveSend = 0; //已寫java的log'
					function checkIsOk(){
						if ( deployJava.getJREs().length > 0 ){
							$("#java_id").addClass("ico_software-off");
							$("#java_download").hide();
							$("#java_download_ok").show();
							
							if( javaLogHaveSend === 0 ){
								ExeAjaxFunction ("EnvironmentTestLogAjaxService.asp","testLogKind=javaFinish");
								javaLogHaveSend = 1 ;
							}
						}else{
							$("#java_id").addClass("ico_software-on");
							$("#java_download_ok").hide();
							$("#java_download").show();
						}
						
						if ( window.isSupportFlash == true){
							$("#flash_id").addClass("ico_software-off");
							$("#flash_download").hide();
							$("#flash_download_ok").show();		

							if( flashLogHaveSend === 0 ){
								ExeAjaxFunction ("EnvironmentTestLogAjaxService.asp","testLogKind=flashFinish");
								flashLogHaveSend = 1;
							}
						}else{
							$("#flash_id").addClass("ico_software-on");
							$("#flash_download_ok").hide();	
							$("#flash_download").show();
						}
					
						if ( deployJava.getJREs().length > 0  && window.isSupportFlash == true && $("#joinNet_ver").val()=='Y' ){
							$("#not_download").hide();
							$("#is_download").show();
							$("#download_pass_btn").show();
							
							clearInterval(sh);
						}else{
							/*
							var str2="";
							if ( !deployJava.getJREs().length > 0 ){
								str2 = "Java ";
							}
							if ( window.isSupportFlash != true ){
								if (str2==""){
									str2 = "Flash ";
								}else{
									str2 = " / Flash ";
								}																					
							}
							if ( $("#joinNet_ver").val()!='Y' ){
								if (str2==""){
									str2 = "JoinNet ";
								}else{
									str2 = " / JoinNet ";
								}																					
							}
							$("#not_download").text("您好! 您的电脑尚未安装以下上课所需软件：【"+str2+"】");		
							*/
							$("#not_download").show();
						}
					}
										
					checkIsOk();					
					sh=setInterval(checkIsOk,3000);
					$("#flash_ver").val( playerVersion.major + "." + playerVersion.minor + "." + playerVersion.release );
					if ( deployJava.getJREs().length > 0 ) {
						$("#java_ver").val( deployJava.getJREs()[0].toString() );
					}
					//alert("You have Flash player " + playerVersion.major + "." + playerVersion.minor + "." + playerVersion.release + " installed");
					//alert(deployJava.getJREs()[1]);
				</script>
				<script>
				var BrowserDetect = {
					init: function () {
						this.browser = this.searchString(this.dataBrowser) || "An unknown browser";
						this.version = this.searchVersion(navigator.userAgent)
							|| this.searchVersion(navigator.appVersion)
							|| "an unknown version";
						this.OS = this.searchString(this.dataOS) || "an unknown OS";
					},
					searchString: function (data) {
						for (var i=0;i<data.length;i++)	{
							var dataString = data[i].string;
							var dataProp = data[i].prop;
							this.versionSearchString = data[i].versionSearch || data[i].identity;
							if (dataString) {
								if (dataString.indexOf(data[i].subString) != -1)
									return data[i].identity;
							}
							else if (dataProp)
								return data[i].identity;
						}
					},
					searchVersion: function (dataString) {
						var index = dataString.indexOf(this.versionSearchString);
						if (index == -1) return;
						return parseFloat(dataString.substring(index+this.versionSearchString.length+1));
					},
					dataBrowser: [
						{
							string: navigator.userAgent,
							subString: "Chrome",
							identity: "Chrome"
						},
						{ 	string: navigator.userAgent,
							subString: "OmniWeb",
							versionSearch: "OmniWeb/",
							identity: "OmniWeb"
						},
						{
							string: navigator.vendor,
							subString: "Apple",
							identity: "Safari",
							versionSearch: "Version"
						},
						{
							prop: window.opera,
							identity: "Opera",
							versionSearch: "Version"
						},
						{
							string: navigator.vendor,
							subString: "iCab",
							identity: "iCab"
						},
						{
							string: navigator.vendor,
							subString: "KDE",
							identity: "Konqueror"
						},
						{
							string: navigator.userAgent,
							subString: "Firefox",
							identity: "Firefox"
						},
						{
							string: navigator.vendor,
							subString: "Camino",
							identity: "Camino"
						},
						{		// for newer Netscapes (6+)
							string: navigator.userAgent,
							subString: "Netscape",
							identity: "Netscape"
						},
						{
							string: navigator.userAgent,
							subString: "MSIE",
							identity: "Explorer",
							versionSearch: "MSIE"
						},
						{
							string: navigator.userAgent,
							subString: "Gecko",
							identity: "Mozilla",
							versionSearch: "rv"
						},
						{ 		// for older Netscapes (4-)
							string: navigator.userAgent,
							subString: "Mozilla",
							identity: "Netscape",
							versionSearch: "Mozilla"
						}
					],
					dataOS : [
						{
							string: navigator.userAgent,
							subString: "Windows NT 6.1",
							identity: "Windows 7"
						},
						{
							string: navigator.userAgent,
							subString: "Windows NT 6.0",
							identity: "Windows Vista"
						},
						{
							string: navigator.userAgent,
							subString: "Windows NT 5.2",
							identity: "Windows Server 2003"
						},
						{
							string: navigator.userAgent,
							subString: "Windows NT 5.1",
							identity: "Windows XP"
						},
						{
							string: navigator.userAgent,
							subString: "Windows XP",
							identity: "Windows XP"
						},
						{
							string: navigator.userAgent,
							subString: "Windows NT 5.0",
							identity: "Windows 2000"
						},
						{
							string: navigator.userAgent,
							subString: "Windows 2000",
							identity: "Windows 2000"
						},
						{
							string: navigator.userAgent,
							subString: "Windows 98",
							identity: "Windows 98"
						},
						{
							string: navigator.userAgent,
							subString: "Win98",
							identity: "Windows 98"
						},
						{
							string: navigator.userAgent,
							subString: "Windows 95",
							identity: "Windows 95"
						},
						{
							string: navigator.platform,
							subString: "Win",
							identity: "Windows"
						},
						{
							string: navigator.platform,
							subString: "Mac",
							identity: "Mac"
						},
						{
							   string: navigator.userAgent,
							   subString: "iPhone",
							   identity: "iPhone/iPod"
						},
						{
							string: navigator.platform,
							subString: "Linux",
							identity: "Linux"
						}
						
					]

				};
				BrowserDetect.init();
				$("#browser").val( BrowserDetect.browser +" "+BrowserDetect.version );
				$("#os").val( BrowserDetect.OS );
				</script>
				</div>
				</div>
				<!-- -------------->
                                                                            
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->