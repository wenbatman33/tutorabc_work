<!--#include virtual="/lib/include/global.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>無標題文件</title>
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>
</head>


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
		
		str_session_date = left(str_get_session_sn,4)&"/"&right(left(str_get_session_sn,6),2)&"/"&right(left(str_get_session_sn,8),2)
		var_arr = Array(g_var_client_sn,str_session_date,str_session_hour)
		str_sql = "select Top 1 session_sn,isnull(attend_level,0) attend_level " 
		str_sql = str_sql & " from client_attend_list(nolock) "   
		str_sql = str_sql & " where client_sn=@client_sn and attend_date=@attend_date and attend_sestime = @attend_sestime "   
		str_sql = str_sql & " order by attend_date,attend_sestime "
				
		arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)    
		'Response.Write("錯誤訊息sql...cccc.:" & g_str_sql_statement_for_debug & "<br>")
		'response.end 
		if (isSelectQuerySuccess(arr_result)) then
			   
			if (Ubound(arr_result) >= 0) then
				str_session_sn = arr_result(0,0)
				str_attend_level = arr_result(1,0)
				str_room = int(right(str_session_sn,3)) '教室
				
				
				'------- 更新進教室狀況 ------- start -------
				var_class_arr = Array(9,str_session_sn,g_var_client_sn)
			    arr_class_result = excuteSqlStatementWrite("Update client_attend_list set Class_yesno=@Class_yesno where session_sn =@session_sn and client_sn =@client_sn" , var_class_arr, CONST_VIPABC_RW_CONN)
	
				'if (arr_result > 0) then 
				    '認證成功
			    'else
				    '認證失敗
			    'end if 
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
					str_class_type = 1  'For regular , 1-1
				end if
				'------tutormeet------ end -----
				'response.write "若沒開教室的話，TutorMeet的資料check一下"
			else
			'錯誤訊息
				'Call alertGo("no room", "/program/member/reservation_class/normal_and_vip.asp", CONST_NOT_ALERT_THEN_REDIRECT )
                Call useASPtoJavascript("alert('现在非您的咨询时间!');window.close();")
				response.end
			end if 
		end if 
		
		'大會堂20進教室，其他為26分以後才能進教室
'		if(str_attend_level < 13) then
'			if (minute(now)>=15 and minute(now)<=26) then close="yes" else close="no" end if 
'		else
'			if (minute(now)>=15 and minute(now)<=20) then close="yes" else close="no" end if 
'		end if
 
'response.write "g_str_session_sn...."&str_session_sn&"<br>"
'response.write "g_str_attend_room...."&str_attend_room&"<br>"
'response.write "g_str_room_rand...."&str_room_rand&"<br>"
'response.write "g_str_chk_member...."&str_chk_member&"<br>"
'response.write "g_var_client_sn...."&g_var_client_sn&"<br>"
'response.end


if str_session_sn<>"" and str_attend_room<>"" and str_room_rand<>""  and str_chk_member=1  and g_var_client_sn<>"" then

%>
	<body Onload="javascript:showIntegratedTutorMeet();window.open('session_feedback.asp?session_sn=<%=str_session_sn%>','Session_feedback')">
<script type="text/javascript">
//顯示教室
function showIntegratedTutorMeet() {
	alert("<%=getWord("cue_session_1")%>,\n<%=getWord("cue_session_2")%>");
	popupIntegratedTutorMeet("<%=str_session_sn%>", "<%=g_var_client_sn%>", "<%=str_attend_room%>", "<%=str_room_rand%>", <%=str_class_type %> ,"", "abc" );
}
</script>
<% else %>
	<body>
<%
Call useASPtoJavascript("alert('现在非您的咨询时间!');window.close();")
'Call alertGo("现在非您的咨询时间! ", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT )
response.end
%>
<% end if%>
</body>
</html>
