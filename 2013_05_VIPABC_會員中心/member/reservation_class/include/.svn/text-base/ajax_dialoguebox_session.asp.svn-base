<!--#include virtual="/lib/include/global.inc"-->
<%
'strOrderSessionData 例:一對多 042010120821 /一對一072010120821
Dim strOrderSessionData : strOrderSessionData = getRequest("strOrderSessionData", CONST_DECODE_NO) 
Dim strClientEmail : strClientEmail = getRequest("strClientEmail", CONST_DECODE_NO)    'client_basic email
Dim strDebug : strDebug = getRequest("strDebug", CONST_DECODE_NO) 

if ( Not isEmptyOrNull(strOrderSessionData) ) then
	strLobbySessionDateTime = right(strOrderSessionData,10) 'date+hour 例:2010120821
	strLobbySessionDate = left(strLobbySessionDateTime,8) 'date，例:20101208
	strLobbySessionTime = right(strLobbySessionDateTime,2) 'hour，例:21
end if

if ( "yes" = strDebug ) then
	response.write "strOrderSessionData=" & strOrderSessionData & "<br>"
	response.write "strLobbySessionTime=" & strLobbySessionTime & "<br>"
	response.write "strLobbySessionDate=" & strLobbySessionDate & "<br>"
	response.write "strClientEmail=" & strClientEmail & "<br>"
end if

'***************判斷是否執行燈箱效果**START**jeanne************
bolOpenBox = true
if ( Not isEmptyOrNull(strClientEmail) AND Not isEmptyOrNull(strLobbySessionTime) ) then
	arrParam = Array(strClientEmail, strLobbySessionTime, 30)
	strSql = " SELECT client_email FROM lobby_dialoguebox WHERE client_email=@client_email AND close_hour=@close_hour AND close_min=@close_min AND valid=1 "
	arrQueryResult = excuteSqlStatementRead(strSql, arrParam, CONST_TUTORABC_R_CONN)
	if ( isSelectQuerySuccess(arrQueryResult) ) then
		if ( Ubound(arrQueryResult)>=0 ) then
			bolOpenBox = false
		end if
	end if
end if
'***************判斷是否執行燈箱效果**END**jeanne************

if ( "yes" = strDebug ) then
	response.write "bolOpenBox=" & bolOpenBox & "<br>"
end if

if ( bolOpenBox ) then
	'convert(varchar(12), session_date ,112) 格式轉換為 20101208  12/9 21
	if ( Not isEmptyOrNull(strLobbySessionTime) AND Not isEmptyOrNull(strLobbySessionDate) ) then
		arrParam = Array(strLobbySessionDate, strLobbySessionTime)
		strSql = " SELECT topic_cn, introduction_cn, sn, lev FROM lobby_session WHERE CONVERT(VARCHAR(12), session_date ,112)=@session_date AND session_time=@session_time AND valid=1 AND ( SessionPeriod = 45 OR SessionPeriod IS NULL ) ORDER BY sn DESC "
		arrQueryResult = excuteSqlStatementRead(strSql, arrParam, CONST_TUTORABC_R_CONN)
		
		if ( "yes" = strDebug ) then
			response.write "g_str_sql_statement_for_debug =" & g_str_sql_statement_for_debug& "<br>"
		end if
		
		if ( isSelectQuerySuccess(arrQueryResult) ) then
			if ( Ubound(arrQueryResult)>=0 ) then
				for i = 0 to Ubound(arrQueryResult,2)
					strTopic = arrQueryResult(0,i)
					if ( Not isEmptyOrNull(strTopic) ) then
						if ( InStr(strTopic,"<br>") <=0 ) then
							strTop0 = strTopic
							strTop1 = ""
						else
							arrTopic = Split(strTopic,"<br>")
							strTop0 = arrTopic(0)
							strTop1 = arrTopic(1)
						end if
					end if
					
					strIntroduction = arrQueryResult(1,i)
					strLobbySn = arrQueryResult(2,i)
					
					if ( "yes" = strDebug ) then
						response.write "strTopic="&strTopic& "<br>"
						response.write "strTop0="&strTopic& "<br>"
						response.write "strTop1="&strTop1& "<br>"
						response.write "strIntroduction="&strIntroduction& "<br>"
						response.write "strLobbySn="&strLobbySn& "<br>"
					end if
					%>
					<div class="ss_desccc_left">
                        <input type="radio" name="input_name_session" class="lobbySession" id="chk_lobby_class_<%=strLobbySn%>" value="99<%=strLobbySessionDateTime%><%=strLobbySn%>" />
                    </div>
					<div class="ss_desccc">
						<h1><span class="font_big font_bold"><%=strTop0%></span><br><%=strTop1%></h1>
					<%
					'處理LV
					strLev = arrQueryResult(3, i)
					arrLev = Split(strLev,",")
					if ( Ubound(arrLev)>=0 ) then
					%>
						<span style="color:#aaaaaa;">适用程度LV<%=arrLev(Lbound(arrLev)+1)%>-<%=arrLev(Ubound(arrLev)-1)%></span><br />
					<%
					end if
					%>
					<%=strIntroduction%>
					</div>
					<div class="f-clear"></div>
                    <div>
                        <span id="span_lobby_class_date_<%=strLobbySn%>" style="display:none"><%=strLobbySessionDate%></span><!--開課日期-->
                        <span id="span_lobby_class_time_<%=strLobbySn%>" style="display:none"><%=strLobbySessionTime%></span><!--開課時間-->
                        <span id="span_lobby_class_topic_<%=strLobbySn%>" style="display:none"><%=strTopic%></span><!--大會堂主題-->
                        <input type="hidden" id="hdn_reservation_class_datetime" value="<%=strLobbySessionDateTime%>" />
                        <input type="hidden" id="hdn_class_type_<%=strLobbySessionDateTime%>" value="99" />
                        <input type="hidden" id="hdn_lobby_session_sn_<%=strLobbySessionDateTime%>" value="<%=strLobbySn%>" />   
                        <input type="hidden" id="hdn_prev_select_class_date" value="<%=strLobbySessionDate%>" />               
                    </div>
					<%
				next
                %>
                <div style="display:none">
                    <div id="ul_reservation_class_info"></div>
                    <form id="form_reservation_class_module" method="post" action="/program/member/reservation_class/reservation_class_exe.asp">
                        <%'hdn_reservation_class_count 隱藏值 是為了要紀錄已新增了幾筆課程%>
                        <input type="hidden" id="hdn_reservation_class_count" name="hdn_reservation_class_count" value="0" />    
                        <%'hdn_total_reservation_class_dt 隱藏值 客戶已預訂的課程日期和時間 YYYYMMDDHH%>
                        <input type="hidden" id="hdn_total_reservation_class_dt" name="hdn_total_reservation_class_dt" value="" />
                        <%'hdn_order_or_reorder_type: 1=預定 2=重新預定%>
                        <input type="hidden" id="hdn_order_or_reorder_type" name="hdn_order_or_reorder_type" value="1" />
                        <span id="span_reservation_class_total_subtract_session"></span>
                    </form>
                </div>
                <%
			end if
		end if
	else
		'todo
		'response.write "沒有訂課日期/時間"
	end if 'if ( Not isEmptyOrNull(strLobbySessionTime) AND Not isEmptyOrNull(strLobbySessionDate) ) then
end if  'if  ( bolOpenBox ) then
%>