<%
Dim contract_sn : contract_sn = ""
Dim str_lead_sn
Dim idno		'身分證
Dim lidno
Dim idno1
'Dim idno2
'Dim family_num
Dim str_content		'合約內容
'sql
Dim str_sqlstr, arr_sql_query_input, arr_sql_query_result, arr_sql_query_result_2, int_sql_result_index
	
	response.write request("csn")&"cccccccccccccccc"
If request("csn") <> "" then
	contract_sn = Request("csn")
	
	str_sqlstr = "Select "
	str_sqlstr = str_sqlstr & " cdata, "
	str_sqlstr = str_sqlstr & " lead_sn "
	str_sqlstr = str_sqlstr & "	From client_temporal_contract "
	str_sqlstr = str_sqlstr & "	Where sn = @sn "
	str_sqlstr = str_sqlstr & "		and client_sn = @client_sn "
	arr_sql_query_input = Array(contract_sn, session("client_Sn"))
	arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
	
	response.write g_str_sql_statement_for_debug
	If Not(isSelectQuerySuccess(arr_sql_query_result) and  Ubound(arr_sql_query_result)>0) Then
		'查詢失敗或沒有值
		'Call alertGo("", "/index.asp", CONST_ALERT_THEN_REDIRECT)
		Response.Write("alertGo")
        Response.End
	Else
		str_content = arr_sql_query_result(0, 0)
		str_lead_sn = arr_sql_query_result(1, 0)
	End If
	
	'取得使用者已設定為同意或不同意
	If instr(str_content,"value=""同意""")>0 Then
	    str_content = replace(str_content,"value=""同意""","value=""同意"" disabled ")
	ElseIf instr(str_content,"value=""不同意""")>0 Then
	    str_content = replace(str_content,"value=""不同意""","value=""不同意"" disabled ")
	End If
	
	If Not(isnull(str_lead_sn) and str_lead_sn<>"") Then
		'找出該客戶lead_basic的身分證
		str_sqlstr = "Select "
		str_sqlstr = str_sqlstr & " isnull(idno,'') as idno, "
		str_sqlstr = str_sqlstr & " isnull(legal_idno,'') as lidno "
		str_sqlstr = str_sqlstr & "	From lead_basic "
		str_sqlstr = str_sqlstr & "	Where sn = @sn "
		
		arr_sql_query_input = Array(str_lead_sn)
		arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
		If (isSelectQuerySuccess(arr_sql_query_result) and  Ubound(arr_sql_query_result)>0) Then
			idno = arr_sql_query_result(0, 0)
			lidno = arr_sql_query_result(1, 0)
		End If
		
		'找出家庭號合約的身分證
		'family_num = 1
        '"select isnull(idno,'') as idno From client_contract_family where contract_sn="&contract_sn    
		str_sqlstr = "Select "
		str_sqlstr = str_sqlstr & "	isnull(idno,'') as idno "
		str_sqlstr = str_sqlstr & "	From client_contract_family "
		str_sqlstr = str_sqlstr & "	Where contract_sn = @contract_sn "
		arr_sql_query_input = Array(contract_sn)
		If (isSelectQuerySuccess(arr_sql_query_result) and  Ubound(arr_sql_query_result)>0) Then
			idno1 = arr_sql_query_result(0, 0)
			'if (family_num = 1) then idno1 = arr_sql_query_result(0, 0)
            'if (family_num = 2) then idno2 = arr_sql_query_result(0, 0)
            'family_num = family_num + 1 	
		End If
	End If
	
	If (idno<>"") Then
	    str_content = Replace(str_content, idno, "******" & right(idno, 4) )
	    str_content = Replace(str_content, left(idno,1) & "*****" & right(idno, 4), "******" & right(idno, 4) )
	end if
	If (lidno<>"") Then str_content = Replace(str_content, lidno, "******" & right(lidno, 4) )
    If (idno1<>"") Then str_content = Replace(str_content, idno1, "******" & right(idno1, 4) )
    'If (idno2<>"") Then str_content = Replace(str_content, idno2, "******" & right(idno2, 4) )	
	
	Response.Write "Test"
	Response.Write "<div align=""center""><font color=""red""><strong>提醒您，堂數及帳號將於款項繳清後一併開通，若您的款項已繳清請忽略此訊息</strong></font><BR><BR></div>"
	Response.Write str_content
	
Else
	Response.Write("no sn")
End If
%>
