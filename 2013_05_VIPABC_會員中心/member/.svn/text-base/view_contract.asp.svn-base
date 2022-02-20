<!--#include virtual="/lib/include/html/template/template_change.asp"-->


<div class="temp_contant">
	<!--內容start-->
   
    <%
	
	'response.write int_contract_sn
	
	Dim str_lead_sn : str_lead_sn = "" 'lead_basic.sn
	Dim str_cdata : str_cdata = "" '同意合約內容
	Dim str_idno : str_idno = ""	'身份證字號
	Dim str_lidno : str_lidno = ""
	
	
	'取得同意後的合約內容
	Dim arr_contract_result : arr_contract_result = getClientCdate(g_var_client_sn)
	
	Dim str_sqlstr 
	Dim var_arr 		'傳excuteSqlStatementRead的陣列
	Dim arr_result	
	
	if (isSelectQuerySuccess(arr_contract_result)) then
		if (Ubound(arr_contract_result) >= 0) then
			
			str_cdata = arr_contract_result(0, 0)
			str_lead_sn = arr_contract_result(1, 0)
			
			
			'取得使用者已設定為同意或不同意
			if ( instr(str_cdata,"name=""agree_or_yet""")>0 ) then 
				str_cdata = replace(str_cdata,"name=""agree_or_yet""","name=""agree_or_yet"" disabled ")
			end if 
			
			'是否列印
			if ( instr(str_cdata,"name=""agree_or_yet""")>0 ) then 
				str_cdata = replace(str_cdata,"name=""print_or_prepage"" style=""display:none""","name=""print_or_prepage"" ")
			end if 
			
			if ( instr(str_cdata,"type=""checkbox""")>0 ) then 
				str_cdata = replace(str_cdata,"type=""checkbox""","type=""checkbox"" disabled ")
			end if 
			
			
			if  ( not isEmptyOrNull(str_lead_sn) ) then
				'找出該客戶lead_basic的身分證
				str_sqlstr="select isnull(idno,'') as idno, isnull(legal_idno,'') as lidno From lead_basic where sn=@sn"    
				var_arr = Array(str_lead_sn)
				arr_result = excuteSqlStatementRead(str_sqlstr, var_arr, CONST_VIPABC_RW_CONN)
				
				if (isSelectQuerySuccess(arr_result)) then
					if (Ubound(arr_result) >= 0) then
						str_idno = arr_result(0, 0) 	
						str_lidno = arr_result(1, 0)
					else
					end if
				end if 
				
			end if
			
			if ( not isEmptyOrNull(str_idno) ) then 
				str_cdata = replace(str_cdata, str_idno, "******" & right(str_idno,4) )
				str_cdata = replace(str_cdata, left(str_idno,1) & "*****" & right(str_idno,4), "******" & right(str_idno,4) )
			end if
			
			if ( not isEmptyOrNull(str_lidno) ) then 
				str_cdata = replace(str_cdata,str_idno,"******"&right(str_idno,4))
			end if 
			
			'Response.Write str_cdata	
			
		else
			'查詢失敗或沒有值
			Call alertGo("", "/index.asp", CONST_ALERT_THEN_REDIRECT)
			Response.End
		
		end if 
	end if 
		
	%>
    
                      
    
    <!--<div id="div_contact_info"></div>-->
	<!--內容end-->
<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>">
	<%=str_cdata%>
</font>     
</div>
