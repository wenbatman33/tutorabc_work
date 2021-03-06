<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type="text/javascript" src="/lib/javascript/AC_OETags.js"></script>
<%
'取得系統時間
dim strSysNowDateTime : strSysNowDateTime = getFormatDateTime(now(), 2)
dim intContractSn : intContractSn = getSession("contract_sn", CONST_DECODE_NO)
dim strGetSessionSn : strGetSessionSn = getSessionRoomTime() '欲找的session_sn
dim strTodaySessionHour : strTodaySessionHour = "" '今日上課 時
dim strTodaySessionDate : strTodaySessionDate = "" '今日上課 年月日
dim intSpecialSn : intSpecialSn = ""
dim intGoinClassMinute : intGoinClassMinute = "" '可進教室時間
dim intEndClassMinute : intEndClassMinute = CONST_DEFAULT_FINISH_CLASS_MINUTE '課程結束時間
dim strGoinClassDateTime : strGoinClassDateTime = ""
dim strFinishClassDateTime : strFinishClassDateTime = ""
dim intNowHour : intNowHour = hour(now)

dim arrParam : arrParam = null '傳入sql的array
dim strSql : strSql = "" 'sql
dim arrSqlResult : arrSqlResult = "" 'sql結果
dim intSqlRow : intSqlRow = 0
	
'------------------ 判斷合約未確認需導至確認頁 ------------- start -----------
'Session("contract_sn")-->有合約但尚未確認合約才會有值
'20101119 阿捨修正 續約客戶邏輯 當有未點擊同意合約時 且 合約服務日已到 才需導至合約確認頁
if ( not isEmptyOrNull(intContractSn)) then  
	arrParam = Array(g_var_client_sn, intContractSn)
    '修正續約客戶
	strSql = " SELECT contract_sn "
	strSql = strSql & " FROM client_purchase(nolock) "
	strSql = strSql & " WHERE (client_sn = @client_sn) AND (contract_sn = @contract_sn) AND (Valid = 1) "
	strSql = strSql & " AND datediff(day,product_sdate,GETDATE())>=0 AND datediff(day, GETDATE(),product_edate)>=0 "
	strSql = strSql & " ORDER BY contract_sn "
	arrSqlResult = excuteSqlStatementRead(strSql, arrParam, CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arrSqlResult)) then
		if (Ubound(arrSqlResult) >= 0) then
			'確認合約
			Call alertGo("", "contract_audit.asp", CONST_NOT_ALERT_THEN_REDIRECT)
			response.end
		end if
	end if
end if
'------------------ 判斷合約未確認需導至確認頁 ------------- end -------------

arrSqlResult = getClientAttendList(g_var_client_sn, strGetSessionSn)
if (isSelectQuerySuccess(arrSqlResult)) then
	if (Ubound(arrSqlResult) >= 0) then
		for intSqlRow = 0 to Ubound(arrSqlResult, 2) '只捉二筆
			if ( intNowHour = 0 ) then
                if ( intSqlRow = 0 and left(strGetSessionSn, 8) = left(arrSqlResult(2, intSqlRow), 8) ) then '捉第一筆
				    strTodaySessionHour = arrSqlResult(1, intSqlRow)
                    strTodaySessionDate = arrSqlResult(0, intSqlRow)
                    intSpecialSn = arrSqlResult(4, intSqlRow)
			    end if
            end if
            if ( intSqlRow = 0 and left(strGetSessionSn, 8) = left(arrSqlResult(2, intSqlRow), 8) ) then '捉第一筆
				strTodaySessionHour = arrSqlResult(1, intSqlRow)
                strTodaySessionDate = arrSqlResult(0, intSqlRow)
                intSpecialSn = arrSqlResult(4, intSqlRow)
			end if
		next
	else
        '導去上課頁面
		Call alertGo("", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT)
		response.end
	end if
end if

'大會堂20進教室，其他為26分以後才能進教室
if ( Not IsEmptyOrNull(intSpecialSn) ) then
    intGoinClassMinute = CONST_DEFAULT_FREELOBBY_GOIN_CLASS_MINUTE
else
    intGoinClassMinute = CONST_DEFAULT_GOIN_CLASS_MINUTE	
end if

if( Not isEmptyOrNull(strTodaySessionHour) ) then 
    if ( strTodaySessionHour = 23 ) then '臨界點, 要跨天
        strTodaySessionFinishHour = 0
        strTodaySessionFinishDate = DateAdd("d", 1, strTodaySessionDate)
    else
        strTodaySessionFinishHour = strTodaySessionHour + 1
        strTodaySessionFinishDate = strTodaySessionDate
    end if

    strGoinClassDateTime = strTodaySessionDate & " " & strTodaySessionHour & ":" & intGoinClassMinute & ":00"
    strFinishClassDateTime = strTodaySessionFinishDate & " " & strTodaySessionFinishHour & ":" & intEndClassMinute & ":00"

    if ( (DateDiff("s", strSysNowDateTime, strGoinClassDateTime) > 0) OR (DateDiff("s", strSysNowDateTime, strFinishClassDateTime) < 0) ) then
        '導去上課頁面
	    Call alertGo("", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT)
	    response.end
    end if
else
    Call alertGo("", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT)
	response.end
end if

dim TypeLib, NewGUID
Set TypeLib = CreateObject("Scriptlet.TypeLib")
NewGUID = TypeLib.Guid
Set TypeLib = Nothing
%>	
<div class="temp_contant">
    <div class="main_membox">
<%
		
Dim str_session_sn : str_session_sn= ""
Dim str_attend_level : str_attend_level = 0 '等級
Dim str_room : str_room = "" '教室
Dim str_get_session_sn : str_get_session_sn = getSessionRoomTime()	  '欲找的session_sn
Dim str_session_hour : str_session_hour = right(str_get_session_sn,2) '捉實際上session_sn.hour
Dim str_session_date 												  '捉實際上session_sn的日期
	
'table
Dim str_sql
Dim var_arr
Dim arr_result
'是否進教室
Dim var_class_arr  
Dim arr_class_result
		
'TutorMeet變數
Dim str_attend_room 
Dim str_room_rand 	'教室亂數
Dim str_chk_member 	'檢查TutorMeet是否有此人
Dim str_class_type	'教室種類
Dim str_status 		'教室是否開放
				
'先判斷現在時間(最後一堂課，會有跨日問題)
Dim str_search_date	'DB撈的日期
Dim str_search_time	'DB撈的時間
Dim str_settime 		'設定是否需要-1 

'新版課後評鑑卡時間點 Start ---
currentDateTime = year(now)&"/"&month(now)&"/"&day(now)&" "&right("0"&hour(now),2)&":"&right("0"&minute(now),2) '現在日期時間
newUIOnlineDateTime = "2011/12/28 23:20" '新版面開始使用日期時間
if currentDateTime > newUIOnlineDateTime then
	redirectFileName = "session_feedback_new.asp"
else
	redirectFileName = "session_feedback.asp"
end if
'新版課後評鑑卡時間點 End ---
		
str_session_date = left(str_get_session_sn,4)&"/"&right(left(str_get_session_sn,6),2)&"/"&right(left(str_get_session_sn,8),2)
var_arr = Array(g_var_client_sn,str_session_date,str_session_hour)
str_sql = "select Top 1 session_sn,isnull(attend_level,0) attend_level " 
str_sql = str_sql & " from client_attend_list(nolock) "   
str_sql = str_sql & " where client_sn=@client_sn and attend_date=@attend_date and attend_sestime = @attend_sestime and valid = 1"   
str_sql = str_sql & " order by sn desc "
				
arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)    

if (isSelectQuerySuccess(arr_result)) then
			   
	if (Ubound(arr_result) >= 0) then
		str_session_sn = arr_result(0,0)
		str_attend_level = arr_result(1,0)
		str_room = int(right(str_session_sn,3)) '教室
				
				
		'------- 更新進教室狀況 ------- start -------
		var_class_arr = Array(9,str_session_sn,g_var_client_sn)
		arr_class_result = excuteSqlStatementWrite("Update client_attend_list set Class_yesno=@Class_yesno where session_sn =@session_sn and client_sn =@client_sn" , var_class_arr, CONST_VIPABC_RW_CONN)
		'------- 更新進教室狀況 ------- end ---------
				
		'------tutormeet------ start -----
		str_attend_room = "session"&str_room
		'注意：要測試時，記得看Tutormeet 是否有資料
		'取教室亂數
		if (str_session_sn <> "") then 
			str_room_rand = GetTutorMeetVar(str_attend_room,str_session_sn,"abc")	
		end if 
		'檢查TutorMeet是否有此人
		str_chk_member = getTutorMeetMember(g_var_client_sn, "abc", "1")	
				
		if (str_attend_level >= 13 ) then
			str_class_type = 8 'For 大會堂
		else
			str_class_type = 1 'For regular , 1-1
		end if
	else
	'錯誤訊息
		Call alertGo("no room", "/program/member/reservation_class/normal_and_vip.asp", CONST_NOT_ALERT_THEN_REDIRECT )
		response.end
	end if 
end if 

if str_session_sn<>"" and str_attend_room<>"" and str_room_rand<>""  and str_chk_member = 1  and g_var_client_sn<>"" then
%>

<script language="javascript" type="text/javascript">
<!--    
// Major version of Flash required
var requiredMajorVersion = 10;
// Minor version of Flash required
var requiredMinorVersion = 0;
// Minor version of Flash required
var requiredRevision = 0;

var airWin;

function popupTutorMeetAirLocalConnWin() {
    if (typeof (airWin) == "undefined" || airWin.closed == true) {
        var browserName = navigator.appName;

        if (browserName == "Netscape" || navigator.userAgent.indexOf('safari') != -1) {
            airWin = window.open('http://' + getHost() + '/tutormeetairloccon/tutormeetair_FF.html', '', 'height=120,width=300,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,status=no');
        } else {
            airWin = window.open('http://' + getHost() + '/tutormeetairloccon/tutormeetair.html', '', 'height=120,width=300,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,status=no');
        }

        airWin.moveTo(0, 0);

        window.opener = top;
        airWin.document.title = "TutorMeet Air";
    }
}

function getSessionType() {
    return "1";
}

function getLanguage() {
    return "2";
}

function getSessionSn() {
    return "<%=str_session_sn%>";
}

function getUserSn() {
    return "<%=g_var_client_sn%>";
}

function getSessionRoomId() {
    return "<%=str_attend_room%>";
}

function getSessionRoomRandStr() {
    return "<%=str_room_rand%>";
}

function getClassType() {
    return "<%=str_class_type%>";
}

function getShowName() {
    return "";
}

function getCompStatus() {
    return "abc";
}

function getSessionCode() {
    <% if ( g_var_client_sn <> "" ) then %>
	var service_code = encodeTutorMeetInputParams("<%=str_session_sn%>","<%=g_var_client_sn%>","<%=str_attend_room%>","<%=str_room_rand%>",<%=str_class_type%>,"","<%=lcase("abc")%>");
	<% end if %>
	return service_code;
}

function showTutorMeet(session_sn, user_sn, session_room_id, session_room_rand_str, class_type, show_name, comp_status, airUpdate) {
    CountReload();
    popupIntegratedTutorMeetWithAirUpdate(session_sn, user_sn, session_room_id, session_room_rand_str, class_type, show_name, comp_status, airUpdate);
}

function showIntegratedTutorMeet() {
    CountReload();
	popupIntegratedTutorMeet(getSessionSn(), getUserSn(), getSessionRoomId(), getSessionRoomRandStr(), getClassType() ,"", getCompStatus());
}

function getHost() {
    return "cn.tutormeet.com";
}

function getAlternateHost() {
    return "www.tutormeet.com";
}

function getWebConnCheck() {
    return "1";
}

function setNewHost(newHost) {
    setTutorMeetWebHostName(newHost);
} 

function afterAirLaunched() {
    GotoURL();
}
//-->
</script>
<br />
<br />
<br />
<br />
<div align="center">
<div id="tutormeet_flash"></div>
<a href="javascript:showIntegratedTutorMeet();"><font size="2" color="#616D7E"><span language=zh-cn>若您无法观看上方图示，请按住Ctrl+鼠标左键点击此连结进入咨询室，谢谢。</span></font></a>
<br />
<br />
<!--<a href="/program/member/go_tutormeet_room_backup.asp"><font size="2" color="#616D7E">若您无法正常开启教室，请点选此连结，谢谢。</font></a>-->
<script language="javascript" type="text/javascript">
function getHost() {
    try{
        if(new Date() > new Date("2013/5/21 15:00:00")){
	        setTutorMeetWebHostName("cn.tutormeet.com");
        }else{
            setTutorMeetWebHostName("www.tutormeet.com");
        }
    }catch(e){setTutorMeetWebHostName("www.tutormeet.com");}
	return host;
}

<!--
var fo = new FlashObject("/flash/tutormeetpopup_new.swf", "tutormeetpopup", "320", "99", "7", "#ffffff");
fo.addParam("wmode", "opaque");
fo.addParam("allowFullscreen", "true");
fo.addParam("bgcolor", "#ffffff");
fo.addParam("quality", "high");
fo.addParam("play", "true");
fo.addParam("loop", "false");
fo.addParam("allowScriptAccess", "sameDomain");
fo.addParam("align", "middle");
fo.addParam("name", "tutormeetpopup");
fo.addParam("id", "tutormeetpopup");
fo.addParam("allowcrossDomainxhr", "true");
fo.write("tutormeet_flash");

/// <summary>
/// 20100825 阿捨新增 連結
/// 接收Flash選單值 改變連結
/// </summary>
/// <param name="text">按鈕值</param>
function GotoURL(){
    window.location = "<%=redirectFileName%>?mod=online&session_sn=<%=str_session_sn%>&language=<%=request("language")%>";
    //alert("<%=getWord("cue_session_1")%>,\n<%=getWord("cue_session_2")%>");
	//popupIntegratedTutorMeet("<%=str_session_sn%>", "<%=g_var_client_sn%>", "<%=str_attend_room%>", "<%=str_room_rand%>", <%=str_class_type%> ,"", "abc" );
}

/// <summary>
/// 20100805 阿捨新增 
///  呼叫ajax載入
/// </summary>
function CountReload(){
	//倒數呼叫
	set=setTimeout(function() {
        GotoURL();
	}, 10000);
}
//-->

</script>
<% else 
        Call alertGo("现在非您的咨询时间! ", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT )
        response.end
   end if
%>
    </div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
    </div>
</div>
<br />
<br />
<br />
<br />