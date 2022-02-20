<!--#include virtual="/lib/include/global.inc" -->
<%
'p_int_score --> 來源 1:industry / 2: role / 3:interest
'p_str_get_value -->page上面的checkbox值
'p_client_sn -->client_basic.sn

sub updateClientDcgsValue (Byval p_int_score, Byval p_str_get_value, Byval p_client_sn)
	
	Dim i
	
	Dim arr_get_value 		'page上的值存為陣列
	
	Dim str_org_value 		'db中原始設定值
	Dim str_addnew_value 	'需要新增的值
	 
	Dim arr_org_value 		'db中原始設定值(轉成陣列)
	Dim arr_addnew_value 	'需要新增的值(轉成陣列)
	
	'塞資料的相關設定值
	Dim str_update_sql 		'回傳的update SQL
	Dim str_sql 
	Dim var_arr
	Dim arr_result
	Dim transaction_opt
	Dim int_add_sql_state 
	Dim int_result_id

	
	
	'先找出原始存在client_profile_log中的資料
	str_sql = " select client_sn,profile_type,profile_sn from client_profile_log "
	str_sql = str_sql & " where profile_type=@profile_type and client_sn=@client_sn and edate is null "

	var_arr = Array(p_int_score,g_var_client_sn)
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)

	if (isSelectQuerySuccess(arr_result)) then
		
		if (Ubound(arr_result) > 0) then  
			'有資料
			for i = 0 to Ubound(arr_result,2) 
				str_org_value = str_org_value & "-" & arr_result(2,i) &"-"
			next
		else
			'無資料
			str_org_value = ""
		end if 
		
	end if 


	'-------- 整理需要異動的值 -------------- start ---------
	if (p_str_get_value <> "") then 
		arr_get_value = split(p_str_get_value,",")
	
		for i = 0 to Ubound(arr_get_value)
		'比對，判斷是否存在
			if (InstrRev(str_org_value,"-"&trim(arr_get_value(i))&"-") >0) then
				'把保留的減掉，留下需要刪除的
				str_org_value = replace(str_org_value,"-"&trim(arr_get_value(i))&"-","") 
			else 
				'addnew (需要新增的)
				str_addnew_value = str_addnew_value &"-"& arr_get_value(i) & "-" 'addnew
			end if 
		next
	
	'刪除部份
	if left(str_org_value,1)="-" then str_org_value=right(str_org_value,len(str_org_value)-1)
	if right(str_org_value,1)="-" then str_org_value=left(str_org_value,len(str_org_value)-1)
	'新增部份
	if left(str_addnew_value,1)="-" then str_addnew_value=right(str_addnew_value,len(str_addnew_value)-1)
	if right(str_addnew_value,1)="-" then str_addnew_value=left(str_addnew_value,len(str_addnew_value)-1)
	
	str_org_value = str_org_value & "--"
	str_addnew_value = str_addnew_value & "--"
	
	arr_org_value = split(str_org_value,"--")
	arr_addnew_value = split(str_addnew_value,"--")
	
	'--------------------------- 整拼寫入 client_profile_log -------------------- start -----------------
	Set transaction_opt = beginTranscation(CONST_VIPABC_RW_CONN)
		'------------------- 處理刪除部份 ----------------------- start ----------		
		For i = 0 To ubound(arr_org_value)
			str_update_sql = ""
			
			if (arr_org_value(i) <> "") then 
				var_arr = Array(p_client_sn,p_int_score,trim(arr_org_value(i)))
				str_update_sql = " update client_profile_log set edate=getdate() "
				str_update_sql = str_update_sql &" where client_sn =@client_sn and profile_type=@profile_type and profile_sn=@profile_sn "
			
				int_add_sql_state = addTranscationSQL(str_update_sql,var_arr,transaction_opt)

			end if 
			'可看單筆是否成功
			if(int_add_sql_state=-1) then 
				'Response.write("Insert Failed_delete<br/>") 
				'Response.Write("錯誤訊息UPDATE_sql....:" & str_update_sql & "<br>")
			end if 
			
		Next
		'------------------- 處理刪除部份 ----------------------- end ----------
		
		'------------------- 處理新增部份 ----------------------- start ----------
		For i = 0 To ubound(arr_addnew_value)
			str_update_sql = "" 
			if (arr_addnew_value(i) <> "") then 
			
				var_arr = Array(p_client_sn,p_int_score,replace(trim(arr_addnew_value(i)),"-",""))
				
				str_update_sql = " insert client_profile_log(client_sn,profile_type,profile_sn,sdate) "
				str_update_sql = str_update_sql &" values(@client_sn,@profile_type,@profile_sn,getdate()) "

				int_add_sql_state = addTranscationSQL(str_update_sql,var_arr,transaction_opt)
			end if 
			'可看單筆是否成功
			if(int_add_sql_state=-1) then 
				'Response.write("Insert Failed_insert<br/>") 
				'Response.Write("錯誤訊INSERT_息sql....:" & str_update_sql & "<br>")
			end if 
			
		Next
		'------------------- 處理新增部份 ----------------------- end ----------
		
		int_result_id = endTranscation(transaction_opt)
	
		if(int_result_id<0) then
			'失敗
			'Response.write g_str_run_time_err_msg
		else
			'成功
			'Response.write "Transcation Insert Success" & "<br>"
		end if
	Set transaction_opt =Nothing
	
	'--------------------------- 整拼寫入 client_profile_log-------------------- start -----------------
end if 
'-------- 整理需要異動的值 --------------end ---------

end sub

'將儲存至DB的資料整理
function getDcgsValue(Byval p_str_value)
	Dim i 
	Dim arr_dcgs		'轉陣列
	dim str_dcgs_value : str_dcgs_value = "" '欲回傳的值
	
	arr_dcgs = split(p_str_value,",")
	For i = 0 to ubound(arr_dcgs)
		str_dcgs_value = str_dcgs_value & "-" & Trim(arr_dcgs(i))& "-"
	next
	
	getDcgsValue = str_dcgs_value
end function 

'-------------- 程式開始處------------------------------

Dim str_industry : str_industry = Trim(Request("industry")) '職業
Dim str_role : str_role = Trim(Request("role"))				'職位
Dim str_interest : str_interest = Trim(Request("interest"))	'興趣
Dim str_focus : str_focus = Trim(Request("focus"))			'學習重點

'塞資料的相關設定值
Dim var_arr 													'傳excuteSqlStatementRead的陣列
Dim arr_result												'接回來的陣列值
Dim str_tablecol 												'table欄位
Dim str_sql

'儲至DB中的型式為：-1--4--7--10-
	str_industry = getDcgsValue(str_industry) 
	str_role = getDcgsValue(str_role) 
	str_interest = getDcgsValue(str_interest) 

'----------------- 職業別、職位別、興趣、focus update client_profile_log ------------- start ----------------

	'call sub
	updateClientDcgsValue 1,str_industry,g_var_client_sn  'industry
	updateClientDcgsValue 2,str_role,g_var_client_sn 		'role
	updateClientDcgsValue 3,str_interest,g_var_client_sn  'interest
	updateClientDcgsValue 4,str_focus,g_var_client_sn  'interest

'----------------- 職業別、職位別、興趣、focus update client_profile_log ------------- end ----------------


	str_tablecol = "industry=@industry,role=@role,interest=@interest,focus=@focus"
	var_arr = Array(str_industry,str_role,str_interest,str_focus,g_var_client_sn)
	arr_result = excuteSqlStatementWrite("Update client_basic set "&str_tablecol&" where sn=@sn" , var_arr, CONST_VIPABC_RW_CONN)

if (arr_result > 0) then
    '完成session 
	session("FinishDCGS") = 1
    '成功
	Call alertGo( getWord("update"), "reservation_class/normal_and_vip.asp", CONST_ALERT_THEN_REDIRECT )
	Response.end
else
	'失敗
end if 
%>
