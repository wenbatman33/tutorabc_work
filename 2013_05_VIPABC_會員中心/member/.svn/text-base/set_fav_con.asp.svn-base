<!--#include virtual="/lib/include/global.inc" -->

<%
Dim var_arr						'小幫手提供的conn開法所需變數
Dim arr_result					'小幫手提供的conn開法所需變數
Dim str_sql						'小幫手提供的conn開法所需變數
Dim str_tablecol

	
Dim str_con_sn		: str_con_sn 		= request("con_sn")
Dim str_fov_valid	: str_fov_valid 	= request("fov") ' 1:有效 0:無效

Dim str_ccr_require_object : str_ccr_require_object = "T"
Dim str_ccr_require_code : str_ccr_require_code = "F"
Dim str_ccr_sdate : str_ccr_sdate = now()
Dim str_ccr_edate : str_ccr_edate = dateadd("yyyy",20,now)
Dim str_ccr_inputer : str_ccr_inputer = 2 '客戶

Dim str_ccr_valid 
If (str_fov_valid = 1) then 
	str_ccr_valid = 0
else
	str_ccr_valid = 1
end if

Dim str_ccr_edtdate : str_ccr_edtdate = now()
Dim str_ccr_editor : str_ccr_editor = 2 '客戶


	'-------- update dcgs_client_class_member_require_bas ------------- start ---------
	var_arr = Array(str_con_sn, g_var_client_sn)
	arr_result = excuteSqlStatementRead("SELECT * FROM dcgs_client_class_member_require_bas WHERE (ccr_require_object = 'T') AND (ccr_require_code = 'F') AND (ccr_object_sn = @ccr_object_sn ) and client_sn = @client_sn  ", var_arr, CONST_VIPABC_RW_CONN)

	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) > 0) then  
			'有資料 (做update處理)
			str_tablecol = "ccr_edtdate=@ccr_edtdate, ccr_editor=@ccr_editor, ccr_editor_sn=@ccr_editor_sn ,ccr_valid=@ccr_valid"
			var_arr = Array(str_ccr_edtdate, str_ccr_editor, g_var_client_sn, str_ccr_valid, g_var_client_sn, str_con_sn )
			arr_result = excuteSqlStatementWrite("Update dcgs_client_class_member_require_bas set "&str_tablecol&" where client_sn=@client_sn AND ccr_object_sn = @ccr_object_sn" , var_arr, CONST_VIPABC_RW_CONN)
	
		else
			'無資料 (做insert處理)
			var_arr = Array(g_var_client_sn, str_ccr_require_object, str_ccr_require_code, str_ccr_sdate, str_ccr_edate, str_con_sn, str_ccr_inputer, g_var_client_sn, str_ccr_valid )
			arr_result = excuteSqlStatementWrite("Insert into dcgs_client_class_member_require_bas (client_sn, ccr_require_object, ccr_require_code, ccr_sdate, ccr_edate, ccr_object_sn, ccr_inputer, ccr_inputer_sn, ccr_valid ) values (@client_sn, @ccr_require_object, @ccr_require_code, @ccr_sdate, @ccr_edate, @ccr_object_sn, @ccr_inputer, @ccr_inputer_sn, @ccr_valid ) " , var_arr, CONST_VIPABC_RW_CONN)
			
		end if 
	end if 
		'-------- update dcgs_client_class_member_require_bas ------------- end ---------


'		sql="SELECT * FROM dcgs_client_class_member_require_bas  "
'		sql = sql & " WHERE (ccr_require_object = 'T') AND (ccr_require_code = 'F') "
'		sql = sql & " AND (ccr_object_sn = "&con_sn&") and client_sn ="&client_sn&" "
'		'sql = sql & " AND ccr_valid ="&fov_valid&" "
'
'		set rs=server.createobject("adodb.recordset")
'		rs.open sql,conn,3,3
'		if rs.eof then 
'			rs.addnew
'			rs("client_sn")=client_sn
'			rs("ccr_require_object")="T"
'			rs("ccr_require_code")="F"
'			rs("ccr_sdate")=now
'			rs("ccr_edate")=dateadd("yyyy",20,now)
'			rs("ccr_object_sn")=con_sn
'			rs("ccr_inputer")=2 '客戶
'			rs("ccr_inputer_sn")=client_sn
'		else
'			rs("ccr_edtdate")=now
'			rs("ccr_editor")=2
'			rs("ccr_editor_sn")=client_sn
'		end if 
'		
'		if fov_valid = 0 then rs("ccr_valid")=1
'		if fov_valid = 1 then rs("ccr_valid")=0
'
'				
'		rs.update
'		rs.close
'		set rs=nothing	
		
Call alertGo(getWord("set_fav_con_success"), "learning_record.asp", CONST_ALERT_THEN_REDIRECT )

%>
