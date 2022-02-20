 <!--#include virtual="/lib/include/global.inc"-->
<script type="text/javascript" src="/lib/javascript/swfobject.js"></script>
<script type="text/javascript" src="/lib/javascript/AC_RunActiveContent.js"></script>
<meta http-equiv="refresh" content="10" />
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>
<script type="text/javascript" src="/lib/javascript/jquery-1.4.2.min.js"></script>
<%
dim strTodayDate : strTodayDate = getFormatDateTime(now(), 5) '目前時間
if g_var_client_sn = "" then
	Response.end
end if
dim intNowHour : intNowHour = hour(now) '目前時
dim intNowMins : intNowMins = minute(now) '目前分
dim session_sn : session_sn = "" '上課教室序號
  
if ( intNowHour = 0 AND intNowMins <= CONST_DEFAULT_FINISH_CLASS_MINUTE ) then
	intNowSessionHour = 23
    '20101207 阿捨修正 跨天後 下課前 要用前一天尋找
    strTodayDate = getFormatDateTime(DateAdd("d", -1, strTodayDate), 5)
elseif ( intNowMins <= CONST_DEFAULT_FINISH_CLASS_MINUTE ) then
    intNowSessionHour = intNowHour - 1
else
    intNowSessionHour = intNowHour
end if

Dim status
Dim str_sql
Dim var_arr
Dim arr_result     

var_arr = Array(g_var_client_sn, intNowSessionHour)
if ( intNowHour = 0 AND intNowMins <= CONST_DEFAULT_FINISH_CLASS_MINUTE ) then
	str_sql = "SELECT client_attend_list.session_sn, session_room_status_1.status " 
	str_sql = str_sql & " FROM client_attend_list(nolock) "   
	str_sql = str_sql & " LEFT OUTER JOIN WEB_OFFICE.web_office.dbo.session_room_status AS session_room_status_1 "
    str_sql = str_sql & " ON client_attend_list.attend_room = session_room_status_1.room "   
	str_sql = str_sql & " WHERE (client_attend_list.client_sn = @g_var_client_sn) AND (client_attend_list.attend_sestime = @hstr) "
	str_sql = str_sql & " AND (client_attend_list.valid = 1) AND (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, '" & strTodayDate & "', 111)) "
else
	str_sql = "SELECT client_attend_list.session_sn, session_room_status_1.status " 
	str_sql = str_sql & " FROM client_attend_list(nolock) " 
	str_sql = str_sql & " LEFT OUTER JOIN WEB_OFFICE.web_office.dbo.session_room_status AS session_room_status_1 "   
	str_sql = str_sql & " ON client_attend_list.attend_room = session_room_status_1.room "  
	str_sql = str_sql & " WHERE (client_attend_list.client_sn = @g_var_client_sn) AND (client_attend_list.attend_sestime = @hstr) "   
	str_sql = str_sql & " AND (client_attend_list.valid = 1) AND (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, '" & strTodayDate & "', 111)) " 
end if
  
arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then
    'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
	if (Ubound(arr_result) >= 0) then
		session_sn = arr_result(0,0)
		status = arr_result(1,0)
	else
		Response.write "<script type='text/javascript'>top.location.href ='class.asp'</script>"
	end if
else
    'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
end if


'新版課後評鑑卡時間點 Start ---
currentDateTime = year(now)&"/"&month(now)&"/"&day(now)&" "&right("0"&hour(now),2)&":"&right("0"&minute(now),2) '現在日期時間
newUIOnlineDateTime = "2011/12/28 23:20" '新版面開始使用日期時間
if currentDateTime > newUIOnlineDateTime then
	redirectFileName = "session_feedback_new.asp"
else
	redirectFileName = "session_feedback.asp"
end if
'新版課後評鑑卡時間點 End ---

'-----join net start-----
dim bolOpenBtn : bolOpenBtn = false '載入flash btn判斷
Dim sql 'sql
Dim strJoinNetLink : strJoinNetLink = "" 'joinNet的連結
sql="SELECT url, status FROM session_room_status(nolock) WHERE (room = '"&right(session_sn,3)&"')"
Dim g_arr_result
Dim str_choise_joinet '網頁顯示的內容
g_arr_result = excuteSqlStatementReadQuick(sql , CONST_WEBOFFICE_CONN)
if (isSelectQuerySuccess (g_arr_result)) then
	if (Ubound(g_arr_result) >= 0) then
		if g_arr_result(1,0) <> 0 then
            bolOpenBtn = true
            'todo relay判斷
            strJoinNetLink = "/joinnet/org/" & g_arr_result(0,0)
			'strJoinNetLink = "http://www.tutorabc.com/joinnet/org/"&g_arr_result(0,0)
            str_choise_joinet = "<div align=""center""><div id=""url_button""></div><br />" &_
            "<a href='"&strJoinNetLink&"'>如果您无法进入教室，请点此连结</a></div>" 
		else
			'if(minute(now) >= 35) then
				'str_choise_joinet="<div align=""center""><font color=""#CC0000""><br><br>"&getWord("SYSTEM_ERROR_PLZ_CALL")&"</font>"&CONST_SERVICE_PHONE&"</div>"
			'else
				str_choise_joinet="<div align=""center""><br><br><img border='0' src='../../images/session/gray_enter_the_session_eng.gif' width='350' height='99'></div>"
			'end if
		end if
	end if 
else
	'錯誤訊息
	Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
end if
Response.write str_choise_joinet
%>
<script type="text/javascript">
<% if ( true = bolOpenBtn ) then %>
/*
var fo = new FlashObject("/flash/url_button_enter_session.swf", "url_button", "315", "93", "7", "#000000");
fo.addParam("wmode", "transparent");
fo.addParam("allowFullscreen", "true");
fo.addParam("bgcolor", "#ffffff");
fo.addParam("quality", "high");
fo.write("url_button");
*/
$("#url_button").html("<a href='javascript:;' onclick='GotoURL();'><img border='0' src='../../images/images_cn/003.gif'></a>");
<% end if %>
/// <summary>
/// 20100825 阿捨新增 連結
/// 接收Flash選單值 改變連結
/// </summary>
/// <param name="text">按鈕值</param>
function GotoURL() {
    parent.window.location = "<%=redirectFileName%>?mod=online&session_sn=<%=session_sn%>&language=<%=request("language")%>"
    window.open('<%=strJoinNetLink%>', '', 'height=100, width=100, top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no');
}
</script>