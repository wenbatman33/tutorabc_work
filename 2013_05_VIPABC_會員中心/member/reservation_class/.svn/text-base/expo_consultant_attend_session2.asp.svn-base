<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/backoffice/inc/conn.asp"-->
<!--#include virtual="/lib/functions/functions_free_lobby_session.asp"-->
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>
<script src="/lib/javascript/AC_OETags.js" language="javascript"></script>
<div>

<script type="text/javascript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function showIntegratedTutorMeet(session_sn,client_sn,attend_room,room_rand_str,class_type) {
	popupIntegratedTutorMeet(session_sn, client_sn, attend_room, room_rand_str,class_type ,"", "abc" );
}
//-->
</script>
</head>
<div align="center">
</br>
</br>

<%
		Dim i, weekstr, aDate ,str_today, str_attend_consultant
        Dim dateAttendDate : dateAttendDate = date
        Dim strValidTime : strValidTime = "" '上課時間
        Dim strValidClassName : strValidClassName = "" '課程名稱
        Dim strValidDescription1 : strValidDescription1 = "" '課程描述
        Dim strValidDescription2 : strValidDescription2 = "" '課程描述2
        Dim strValidMaterial : strValidMaterial = "" '教材
        Dim strValidConsultant : strValidConsultant = "" '顧問
        Dim strValidRoom : strValidRoom = "" 'room
        Dim strSessionSn : strSessionSn = "" 'sessionSn
        Dim str_room_rand : str_room_rand = "" 'str_rand_str
        dateAttendDate = getFormatDateTime(dateAttendDate,5)
        arrFreeelobbySessionData = getFreeLobbySessionData1(dateAttendDate)
        strValidTime = arrFreeelobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_TIME)
        strValidClassName = arrFreeelobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_NAME)
        strValidDescription1 = arrFreeelobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION1)
        'strValidDescription2 = arrFreeelobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION2) 
        strValidMaterial = arrFreeelobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_MATERIAL)
        strValidConsultant = arrFreeelobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CONSULTANT)   
		strValidRoom = arrFreeelobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_ROOM_NUMBER)		
        '串出session_sn
		strSessionSn = getFormatDateTime(dateAttendDate,7)&left(strValidTime,2)&strValidRoom 
        strValidRoom = "session"&strValidRoom
		str_room_rand = GetTutorMeetVar(strValidRoom,strSessionSn,"abc")

%>
 <script language="javascript" >
	    var requiredMajorVersion = 10;
		// Minor version of Flash required
		var requiredMinorVersion = 0;
		// Minor version of Flash required
		var requiredRevision = 0;

		var airWin;
		function popupTutorMeetAirLocalConnWin() {
		   if (typeof(airWin) == "undefined" || airWin.closed == true)  {  
			  var browserName = navigator.appName;

			  if (browserName == "Netscape"||  navigator.userAgent.indexOf('safari')!= -1)  {
				  airWin = window.open('http://'+getHost()+'/tutormeetairloccon/tutormeetair_FF.html','','height=120,width=300,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,status=no');
			  } else {
				  airWin = window.open('http://'+getHost()+'/tutormeetairloccon/tutormeetair.html','','height=120,width=300,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,status=no');
			  }

			  airWin.moveTo(0,0);

			  window.opener = top;
			  airWin.document.title="TutorMeet Air";
		   }
		}
		
		function getSessionCode() {

		    var service_code = encodeTutorMeetInputParams("<%=strSessionSn%>", "<%=strValidConsultant%>", "<%=strValidRoom%>", "<%=str_room_rand%>", 2, "", "abc");

			return service_code;
		}
		</script>
		<script language="JavaScript" type="text/javascript">
		<!--
		// Version check for the Flash Player that has the ability to start Player Product Install (6.0r65)
		var hasProductInstall = DetectFlashVer(6, 0, 65);

		// Version check based upon the values defined in globals
		var hasRequestedVersion = DetectFlashVer(requiredMajorVersion, requiredMinorVersion, requiredRevision);

		// Check to see if a player with Flash Product Install is available and the version does not meet the requirements for playback
		if ( hasProductInstall && !hasRequestedVersion ) {
			// MMdoctitle is the stored document.title value used by the installation process to close the window that started the process
			// This is necessary in order to close browser windows that are still utilizing the older version of the player after installation has completed
			// DO NOT MODIFY THE FOLLOWING FOUR LINES
			// Location visited after installation is complete if installation is required
			var MMPlayerType = (isIE == true) ? "ActiveX" : "PlugIn";
			var MMredirectURL = window.location;
			document.title = document.title.slice(0, 47) + " - Flash Player Installation";
			var MMdoctitle = document.title;
			AC_FL_RunContent(
				"src", "playerProductInstall",
				"FlashVars", "MMredirectURL="+MMredirectURL+'&MMplayerType='+MMPlayerType+'&MMdoctitle='+MMdoctitle+"",
				"width", "140",
				"height", "25",
				"align", "middle",
				"id", "popupairconsultant",
				"quality", "high",
				"bgcolor", "#ffffff",
				"name", "popupairconsultant",
				"allowScriptAccess","sameDomain",
				"type", "application/x-shockwave-flash",
				"pluginspage", "http://www.adobe.com/go/getflashplayer"
			);

		} else if (hasRequestedVersion) {
			// if we've detected an acceptable version
			// embed the Flash Content SWF when all tests are passed
			AC_FL_RunContent(
					"src", "popupairconsultant",
					"width", "140",
					"height", "25",
					"align", "middle",
					"id", "popupairconsultant",
					"quality", "high",
					"bgcolor", "#ffffff",
					"name", "popupairconsultant",
		//			"flashvars",'historyUrl=history.htm%3F&lconid=' + lc_id + '',
					"allowScriptAccess","sameDomain",
					"type", "application/x-shockwave-flash",
					"pluginspage", "http://www.adobe.com/go/getflashplayer",
					"wmode", "opaque"			
			);
		  } else {  // flash is too old or we can't detect the plugin
			var alternateContent = 'This content requires the Adobe Flash Player. '
			+ '<a href=http://www.adobe.com/go/getflash/>Get Flash</a>';
			document.write(alternateContent);  // insert non-flash content
		  }
		// -->
		</script>
	

   <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
