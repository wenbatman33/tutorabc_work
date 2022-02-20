<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>

<%
'Dim now_hour : now_hour=hour(now)
'dim hc
'dim hstr
'dim session_sn
  
'if now_hour =0 and int(minute(now))<15 then
'	now_hour =24
'end if

'if minute(now)<15 then hc=-1 else hc=0
'hstr=now_hour+hc
'session_sn=""


'Dim str_sql
'Dim var_arr
'Dim arr_result     

'		var_arr = Array(g_var_client_sn,hstr)
'		if hour(now)=0 and int(minute(now))<15 then
'		    str_sql = "SELECT client_attend_list.session_sn, session_room_status_1.status " 
'		    str_sql = str_sql & " FROM client_attend_list WITH (nolock) LEFT OUTER  "   
'		    str_sql = str_sql & " JOIN WEB_OFFICE.web_office.dbo.session_room_status AS session_room_status_1 "   
'		    str_sql = str_sql & " WHERE (client_sn = @g_var_client_sn) AND (attend_sestime = @hstr) "
'		    str_sql = str_sql & " AND (valid = 1) AND (CONVERT(varchar, attend_date, 111) = CONVERT(varchar, DATEADD(dd, - 1, GETDATE()), 111)) "
'	    else
'	         str_sql = "SELECT client_attend_list.session_sn, session_room_status_1.status " 
'	         str_sql = str_sql & " FROM client_attend_list WITH (nolock)  " 
''	         str_sql = str_sql & " LEFT OUTER JOIN WEB_OFFICE.web_office.dbo.session_room_status AS session_room_status_1 "   
'	         str_sql = str_sql & " ON client_attend_list.attend_room = session_room_status_1.room WHERE "  
'	         str_sql = str_sql & " client_attend_list.client_sn = @g_var_client_sn) AND (client_attend_list.attend_sestime = @hstr) "   
'	         str_sql = str_sql & "  ( AND (client_attend_list.valid = 1) AND (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE(), 111)) " 
'	    end if
'	    
''		arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)    
'		'Response.Write("錯誤訊息sql...cccc.:" & g_str_sql_statement_for_debug & "<br>")
'		'response.end 
'		if (isSelectQuerySuccess(arr_result)) then
'			if (Ubound(arr_result) >= 0) then
'				session_sn = arr_result(0,0)
'				status = arr_result(1,0)
'			else
'				Call alertGo("", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT )
'			end if
'		end if	
'		
'		
'		'-----join net start-----
'		Dim sql 'sql
'		sql="SELECT weboffice_fullname FROM cfg_weboffice_name(nolock) WHERE (rno = '"&right(session_sn,3)&"')"
'	    Dim g_arr_result
'		Dim str_choise_joinet '網頁顯示的內容
'		g_arr_result = excuteSqlStatementReadQuick(sql , CONST_VIPABC_RW_CONN)
'		if (isSelectQuerySuccess (g_arr_result)) then
'			if (Ubound(g_arr_result) >= 0) then
'				str_choise_joinet="<a href=""session_feedback.asp?mod=online&session_sn="&session_sn&"&language="&request("language")&""" target=""_parent"" onclick=""window.location.href('http://" & CONST_ABC_WEBHOST_NAME & "/jnj_file/new/"&g_arr_result(0,0)&"')""><img border='0' src='../../images/session/enter-session-button-gb.gif' width='350' height='99'></a>" 
'			else
''				str_choise_joinet="<div align=""center""><font color=""#CC0000""><br><br>"&getWord("SYSTEM_ERROR_PLZ_CALL")&"</font>"&CONST_SERVICE_PHONE&"</div>"
'			end if 
'		else
'			'錯誤訊息
'			Response.Write("錯誤訊息" & g_arr_result & "<br>")
'		end if
		%>
		<div class="temp_contant">
			<!--內容start-->
			<div class="main_membox">
				111	
			</div>
			<!--內容end-->
			<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
		</div>   