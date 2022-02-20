<!--#include virtual="/lib/include/html/template/template_change.asp"-->
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
        arrFreeLobbySessionData = getFreeLobbySessionData1(dateAttendDate)
        strValidTime = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_TIME)
        strValidClassName = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_NAME)
        strValidDescription1 = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION1)
        'strValidDescription2 = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION2) 
        strValidMaterial = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_MATERIAL)
        strValidConsultant = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CONSULTANT)   
		strValidRoom = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_ROOM_NUMBER)		
        '串出session_sn
		strSessionSn = getFormatDateTime(dateAttendDate,7)&left(strValidTime,2)&strValidRoom 
        strValidRoom = "session"&strValidRoom
		str_room_rand = GetTutorMeetVar(strValidRoom,strSessionSn,"abc")

%>
	   <td width="178"><a href="expo_consultant_attend_session2.asp" />Open TutorMeet Air         </a></td>

	  <td width="178"><a href="javascript:void(0)" onclick="showIntegratedTutorMeet('<%=strSessionSn%>','<%=strValidConsultant%>','<%=strValidRoom%>','<%=str_room_rand%>',2); return false;" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image8','','/images/expo/p05ov.jpg',1)"><img src="/images/expo/p05.jpg" name="Image8" width="178" height="41" border="0" id="Image8" /></a></td>

   <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
