<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/functions_hash1way.asp" -->
<!--#include virtual="/lib/functions/functions_hexvalue.asp" -->
<!--#include virtual="/lib/functions/functions_chgkeyword.asp" -->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"--><!--TREE VIP系統信-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>無標題文件</title>
</head>

<body>
<%

Dim str_client_email : str_client_email = session("client_email")
Dim str_fname : str_fname = session("client_fname")

Dim str_password : str_password = Trim(Request("pwd_passwd"))			'User 輸入的舊密碼
Dim str_new_password : str_new_password = Trim(Request("pwd_passwd_retry"))	'新密碼

Dim str_chk_password : str_chk_password =""			'原密碼
Dim str_rand : str_rand = ""						'解碼字串
Dim str_strencyptedpassword : str_strencyptedpassword = "" '密碼編碼
Dim str_joindate : str_joindate = date & " " & formatdatetime(now,4) & ":" & second(now)
Dim client_ip : client_ip = getIpAddress

'塞資料的相關設定值
Dim var_arr 																							'傳excuteSqlStatementRead的陣列
Dim arr_result																						'接回來的陣列值
Dim str_tablecol 																						'table欄位
Dim str_tablevalue 																					'table值
Dim exec_result

'mail的部份
Dim arr_email_from_paras
Dim arr_email_to_paras
Dim arr_set_paras
Dim arr_opt_paras
Dim str_mail_subjcet : str_mail_subjcet = getWord("MODIFY_PASSWORD")
Dim str_mail_url : str_mail_url=""

'---------------- 先判斷舊密碼是否正確 ------------------------- start ----------------
	var_arr = Array(str_client_email)
	arr_result = excuteSqlStatementRead("SELECT passwd,salt FROM client_passwd WHERE (client_user = @client_user)", var_arr, CONST_VIPABC_RW_CONN)

if (isSelectQuerySuccess(arr_result)) then
	
	if (Ubound(arr_result) > 0) then  
		'有資料
		str_chk_password = arr_result(0,0) 'password
		
		str_password = Replace(Replace(str_password,"'",""),"""","") & arr_result(1,0)
		str_password = HashEncode(str_password)
	
		if (str_chk_password = str_password) then
			'---------------- 變更密碼 ------------------------- start ----------------
			
			
			str_password=replace(str_new_password,"'","''")  '新密碼
			str_rand = getSalt(Len(str_password))
			str_strencyptedpassword = str_password & str_rand
			str_strencyptedpassword = hashencode(str_strencyptedpassword)
			
			str_tablecol = "passwd=@passwd,salt=@salt,lastchgdate=@lastchgdate,lastchgip=@lastchgip"
			var_arr = Array(str_strencyptedpassword,str_rand,str_joindate,client_ip,g_var_client_sn)
			arr_result = excuteSqlStatementWrite("Update client_passwd set "&str_tablecol&" where client_sn=@client_sn" , var_arr, CONST_VIPABC_RW_CONN)
		
			if (arr_result > 0) then 
				'Response.Write("影響筆數 " & arr_result & "<br>")
				'----------------------發送密碼變更信-------------- start -----------
				''str_mail_url = "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/mail_template/forget_password/forget_password.asp?mail_template_name="&str_fname&"&mail_template_password="&str_new_password
			
				''arr_email_from_paras 	= Array(CONST_MAIL_ADDRESS_SERVICES_ABC,"",str_mail_subjcet,"","",str_mail_url)
				'arr_email_to_paras 	= Array("pipiyao@tutorabc.com",str_fname)
				''arr_email_to_paras 	= Array(str_client_email,str_fname)
				
				''arr_set_paras 		= Array(CONST_MAIL_SERVER_CLIENT, CONST_EMAIL_CHARSET)
				''arr_opt_paras			= Array(true)
								
				'call function sendEmail
				''sendEmail arr_email_from_paras, arr_email_to_paras, arr_set_paras, arr_opt_paras, CONST_TUTORABC_WEBSITE

				'20110214 tree VB10110205 VIPABC系統信件 --開始--
				set SendMailvip = new VIPSystemSendMail											'建立class物件
				SendMailvip.str_sendlettermail_sn 			= "0017"							'第0017封信 	主旨：密碼修改						
				SendMailvip.strEmailToAddr 					= str_client_email					'客戶的電子郵件
				SendMailvip.client_cname 					= str_fname							'客戶姓名
				SendMailvip.client_modify_data_date 		= str_joindate						'客戶修改資料時間
				SendMailvip.client_password 				= str_new_password					'客戶新密碼				
				SendMailvip.client_new_password				= str_password&str_cancel_success_msg_for_email	'客戶新密碼
				Call SendMailvip.MainSendMail()													'呼叫class物件的函式來發信
				set SendMailvip = nothing														'釋放CLASS物件
				'20110214 tree VB10110205 VIPABC系統信件 --結束--  
					
				if (exec_result <> CONST_FUNC_EXE_SUCCESS) then
					' E-mail 錯誤訊息
					Response.Write(g_str_run_time_err_msg)
				end if
				
				%>
                <form name="frm_profile_pwd_exe" id="frm_profile_pwd_exe" action="profile.asp" method="post">
                	<input name="hdn_get_password" type="hidden" value="true" />
                    <script language="javascript" type="text/javascript">
						document.frm_profile_pwd_exe.submit();
					</script>
					
                </form>
                
                <%
				'----------------------發送密碼變更信-------------- end -----------
			else
				'錯誤訊息(沒有變更到資料)
			end if 
			'---------------- 變更密碼 ------------------------- end ----------------
		 
		else
			'比對密碼失敗
			Call alertGo( getWord("please_confirm_old_password_error"), "", CONST_ALERT_THEN_GO_BACK)
			Response.End
			
		end if 
		
	else
	
	
	end if 
end if 
'---------------- 先判斷舊密碼是否正確 ------------------------- end ----------------	
%>
</body>
</html>
