<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/functions_hash1way.asp" -->
<!--#include virtual="/lib/functions/functions_hexvalue.asp" -->
<!--#include virtual="/lib/functions/functions_chgkeyword.asp" -->
<%
'檢查客戶的e-mail 和 是否認證
'g_str_score：來源  
'目前使用有四個地方：登入、忘記密碼信、認證、檢核是否有相同email 、chkpassword (舊密碼)


Dim str_score : str_score = Trim(Request("get_score"))
Dim str_email_addr : str_email_addr = Trim(Request("get_email_addr"))
Dim str_login_password : str_login_password =Trim(Request("get_login_password")) '是否驗證密碼 有值:是 無值:否


Dim str_passwd : str_passwd=""
Dim str_chk_passwd : str_chk_passwd =""
Dim str_salt : str_salt = ""
Dim int_valid : int_valid = 0 

'塞資料的相關設定值
Dim var_arr 					'傳excuteSqlStatementRead的陣列
Dim arr_result				'接回來的陣列值
Dim str_sql


	var_arr = Array(str_email_addr)
'g_arr_result = excuteSqlStatementRead("select valid from client_basic where email=@email and (email <> '')", var_arr, CONST_VIPABC_RW_CONN)
	str_sql = "select client_basic.sn," 
	str_sql = str_sql & "client_basic.email,client_basic.valid, "
	str_sql = str_sql & "client_passwd.passwd, client_passwd.salt "
	str_sql = str_sql & "from client_basic(nolock) "
	str_sql = str_sql & "inner join  client_passwd(nolock) on client_basic.sn = client_passwd.client_sn "
	str_sql = str_sql & "where email=@email and (email <> '') and web_site='c' "

	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)

if (isSelectQuerySuccess(arr_result)) then
    '有資料
	
    if (Ubound(arr_result) >= 0) then  
		
		int_valid = arr_result(2,0)

		'-----------判斷是否帳號已認證過 ----------- start --------------
		if (int_valid = 0 ) then 
		
			
			if (str_score = "authenticate" ) then
				'success
				Response.Write getWord("AJAX_MEMBER_LOGIN_5")
				
			else
				'非常抱歉，您尚未完成E-mail认证步骤！请先至您的电子邮箱完成认证，或是选择系统
				'重发认证信
				Response.Write getWord("ajax_member_login_1")&"<a onclick='showAndFocuse();' href='javascript:void(0);'>"&getWord("ajax_member_login_2")&"</a>"
			end if  
			
		else
			'已認證過
			
			if (str_score = "authenticate" ) then
				'您已认证过电子邮箱，不需重覆认证！
				Response.Write getWord("ajax_member_login_4")
			
			else
							
				if (str_login_password <> "" ) then 
				
					'------------- 驗認密碼 ------------------------ start --------
					str_passwd = arr_result(3,0)
					str_salt = arr_result(4,0)
			
					if (str_passwd <> "" and str_salt <> "") then
						'密碼驗證
						str_chk_passwd = Replace(Replace(Trim(str_login_password),"'",""),"""","") & str_salt
						str_chk_passwd = hashencode(str_chk_passwd)
					
						if (str_passwd = str_chk_passwd) then
							'成功
							'success
							Response.write getWord("ajax_member_login_5")
						else
							'失敗
							'抱歉！您输入的密码错误！请重新输入！
							Response.write getWord("ajax_member_login_6")
						end if 
					else
						'錯誤訊息
						'没有密码，请确认！
						Response.write getWord("ajax_member_login_7")
					end if 
					'------------- 驗認密碼 ------------------------ end --------
					
				else
					Response.Write getWord("ajax_member_login_5")	'success
				end if
				 			
			end if 
		end if 
		'-----------判斷是否帳號已認證過 ----------- end --------------
	else
		'無此帳號
		'抱歉！系统找不到您输入的电子邮箱！请重新输入
		Response.Write getWord("ajax_member_login_8")
	end if 
	
end if 


%>
