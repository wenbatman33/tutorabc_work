<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<%
'停用頁面的快取暫存
Call setPageNoCache()

Dim int_run_result
Dim int_insert_result
dim int_update_result
'client email
Dim strClientEmail : strClientEmail = getSession("client_email", CONST_DECODE_NO)
dim strWhatToDo : strWhatToDo = getRequest("do", CONST_DECODE_NO)

Dim obj_sql_opt_read : Set obj_sql_opt_read = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")

if ( strWhatToDo = "remind" ) then
	'客戶要提醒大會堂
	for i = 0 to 23
		int_run_result = obj_sql_opt_read.excuteSqlStatementEach(" SELECT count(*) AS sum FROM lobby_dialoguebox WHERE close_hour = " & i & " AND client_email= '" & strClientEmail & "'",  null, CONST_TUTORABC_R_CONN)
		if ( int_run_result <> CONST_FUNC_EXE_SUCCESS ) then
'			Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
'			Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
            response.write "false"
		end if
		
		if ( obj_sql_opt_read("sum") = 0 ) then
			strSqlWrite = " INSERT INTO lobby_dialoguebox (client_email, close_hour, close_min, create_at, valid, mdf_date, modifier) VALUES ('" & strClientEmail & "', " & i & ", 30, getdate(), 0, getdate(), 3177) "
			int_insert_result = excuteSqlStatementWrite(strSqlWrite, null, CONST_TUTORABC_RW_CONN)
			if ( int_insert_result <> CONST_FUNC_EXE_SUCCESS ) then
				'Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
				'Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
				response.write "false"
			end if
		else
			strSqlUpdate = " UPDATE lobby_dialoguebox SET valid = 0 WHERE client_email = '" & strClientEmail & "' AND close_hour = " & i
			int_update_result = excuteSqlStatementWrite(strSqlUpdate, null, CONST_TUTORABC_RW_CONN)
			if ( int_update_result <> CONST_FUNC_EXE_SUCCESS ) then
				'Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
				'Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
				response.write "false"
			end if
		end if
	next
end if

if ( strWhatToDo = "noremind" ) then
	'客戶設定不提醒大會堂      
	for i = 0 to 23
		int_run_result = obj_sql_opt_read.excuteSqlStatementEach(" SELECT count(*) AS sum FROM lobby_dialoguebox WHERE close_hour = " & i & " AND client_email = '" & strClientEmail & "'",  null, CONST_TUTORABC_R_CONN)
		if ( int_run_result <> CONST_FUNC_EXE_SUCCESS ) then
'			Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
'			Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
            response.write "false"
		end if
		if ( obj_sql_opt_read("sum") = 0 ) then
			strSqlWrite = " INSERT INTO lobby_dialoguebox (client_email, close_hour, close_min, create_at, valid, mdf_date, modifier) VALUES ('" & strClientEmail & "', " & i & ", 30, getdate(), 1, getdate(), 3177)"
			int_insert_result = excuteSqlStatementWrite(strSqlWrite, null, CONST_TUTORABC_RW_CONN)
			if ( int_insert_result <> CONST_FUNC_EXE_SUCCESS ) then
'				Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
'				Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
				response.write "false"
			else
				response.write "true"
			end if
		else
			strSqlUpdate = " UPDATE lobby_dialoguebox SET valid = 1 WHERE client_email = '" & strClientEmail & "' AND close_hour = " & i
			int_update_result = excuteSqlStatementWrite(strSqlUpdate, null, CONST_TUTORABC_RW_CONN)
			if ( int_update_result <> CONST_FUNC_EXE_SUCCESS ) then
'				Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
'				Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
				response.write "false"
			else
				response.write "true"
			end if
		end if
	next
end if
%>
