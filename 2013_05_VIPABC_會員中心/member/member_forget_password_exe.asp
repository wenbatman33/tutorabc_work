<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/functions_hash1way.asp" -->
<!--#include virtual="/lib/functions/functions_hexvalue.asp" -->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"-->
<%
dim bolDebugMode : bolDebugMode = false
Dim str_member_forget_password_email

Dim str_fname : str_fname = "" '客戶名字
Dim str_joindate : str_joindate = date & " " & formatdatetime(now,4) & ":" & second(now)
Dim client_ip : client_ip = getIpAddress
Dim str_email_address	: str_email_address="" '客戶 email
Dim str_tmp_sql				
			
'塞資料的相關設定值
Dim var_arr 			'傳excuteSqlStatementRead的陣列
Dim arr_result			'接回來的陣列值
Dim str_tablecol 	    'table欄位
Dim str_tablevalue 'table值

'hex的部份
Dim str_password
Dim str_rand
Dim str_strencyptedpassword

'mail的部份
Dim arr_email_from_paras
Dim arr_email_to_paras
Dim arr_set_paras
Dim arr_opt_paras
Dim str_mail_subjcet : str_mail_subjcet = getWord("get_new_password")
Dim str_mail_url : str_mail_url=""

dim strPasswdMailCookie : strPasswdMailCookie = Request.Cookies("passwd_mail")

if ( bolDebugMode ) then
    response.Write "strPasswdMailCookie:" & strPasswdMailCookie & "<br />"
    response.Write "sCStr(strPasswdMailCookie):" & sCStr(strPasswdMailCookie) & "<br />"
end if

'20120909 阿捨新增 新增cookie避免重覆發信10分鐘
if ( "true" = sCStr(strPasswdMailCookie) ) then
    Response.Write("请勿于10分钟内重复索取密码信, 以免造成无法登入的状况!")
else
	str_member_forget_password_email	 = Trim(Request("get_email_addr"))
	str_email_address = Trim(Request("get_email_addr"))

	var_arr = Array(str_member_forget_password_email)
	str_tmp_sql = "SELECT TOP 1 email, "
	str_tmp_sql = str_tmp_sql & "CASE "
	str_tmp_sql = str_tmp_sql & "WHEN (fname IS NOT NULL AND lname IS NOT NULL) THEN fname+' '+lname "
	str_tmp_sql = str_tmp_sql & "WHEN (cname is not null) THEN cname "
	str_tmp_sql = str_tmp_sql & "WHEN (fname is not null) THEN fname "
	str_tmp_sql = str_tmp_sql & "WHEN (lname is not null) THEN lname "
	str_tmp_sql = str_tmp_sql & "ELSE 'Dear member' "
	str_tmp_sql = str_tmp_sql & "END AS member_name, "
	str_tmp_sql = str_tmp_sql & "sn AS client_sn "
	str_tmp_sql = str_tmp_sql & "FROM client_basic "
	str_tmp_sql = str_tmp_sql & "WHERE email = @email "
	arr_result = excuteSqlStatementRead(str_tmp_sql, var_arr, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_result)) then
         '有資料
        if (Ubound(arr_result) > 0) then  
		    str_email_address = arr_result(0,0)
		    str_fname = arr_result(1,0)
		    g_var_client_sn = arr_result(2,0)
		    '---------- 寫入 client_passwd ----------- start ----------------------
		    str_password = hexValue(8)
		    'Create a salt value For the new password
		    str_rand = getSalt(8)
		
		    str_strencyptedpassword = str_password & str_rand
		    str_strencyptedpassword = hashencode(str_strencyptedpassword)
            			
		    str_tablecol = "passwd = @passwd, salt = @salt, lastchgdate = @lastchgdate, lastchgip = @lastchgip"
		    var_arr = Array(str_strencyptedpassword,str_rand,str_joindate,client_ip,g_var_client_sn)
		    arr_result = excuteSqlStatementWrite("UPDATE client_passwd SET "&str_tablecol&" WHERE client_sn = @client_sn", var_arr, CONST_VIPABC_RW_CONN)
		    '---------- 寫入 client_passwd ----------- end ----------------------
		
		    if (arr_result > 0) then 
			    '---------- 寄忘記密碼信------------------- start --------------------
    '			str_mail_url = "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/mail_template/forget_password/forget_password.asp?mail_template_name="&Escape(str_fname)&"&mail_template_password="&str_password
    '			
    '			arr_email_from_paras 	= Array(CONST_MAIL_ADDRESS_SERVICES_ABC,"",str_mail_subjcet,"","",str_mail_url)
    '			'arr_email_to_paras 	= Array("pipiyao@tutorabc.com",str_fname)
    '			arr_email_to_paras 	= Array(str_member_forget_password_email,str_fname)
    '			arr_set_paras 		= Array(CONST_MAIL_SERVER_CLIENT, CONST_EMAIL_CHARSET)
    '			arr_opt_paras			= Array(true)
    '								
    '			'call function sendEmail
    '			sendEmail arr_email_from_paras, arr_email_to_paras, arr_set_paras, arr_opt_paras, CONST_TUTORABC_WEBSITE

			    '20110214 tree VB10110205 VIPABC系統信件 --開始--
			    set SendMailvip = new VIPSystemSendMail                                                                           '建立class物件
			    SendMailvip.str_sendlettermail_sn = "0012"							                                                '第0012封信 	主旨：索取新密码						
			    SendMailvip.strEmailToAddr = str_member_forget_password_email                                   '客戶的電子郵件
			    SendMailvip.client_cname = str_fname							                                                        '客戶姓名
			    SendMailvip.client_modify_data_date = str_joindate						                                    '客戶修改資料時間
			    SendMailvip.client_new_password = str_password&str_cancel_success_msg_for_email	'客戶新密碼
			    Call SendMailvip.MainSendMail()													                                        '呼叫class物件的函式來發信
			    set SendMailvip = nothing														                                                '釋放CLASS物件
			    '20110214 tree VB10110205 VIPABC系統信件 --結束--
            
                '20120909 阿捨新增 新增cookie避免重覆發信10分鐘
                Response.Cookies("passwd_mail") = "true"
	            Response.Cookies("passwd_mail").Expires = DateAdd("n", 10, Now())
		    else 'if (arr_result > 0) then
			    '錯誤訊息
		    end if 'if (arr_result > 0) then 
		    'call sub alertGo
		    'Call alertGo("密碼信件已送出", "/program/member/member_login.asp", CONST_ALERT_THEN_REDIRECT)
		    Response.Write("success")
        else 'if (Ubound(arr_result) > 0) then
		    '沒有資料
            '錯誤訊息
            'Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
            'Response.End()
	        'Call alertGo("資料錯誤 請輸入您註冊時的電子郵件帳號", "", CONST_ALERT_THEN_GO_BACK)
	        Response.Write getWord("data_error_please_input_email_account")
        end if ''if (Ubound(arr_result) > 0) then
    end if 'if (isSelectQuerySuccess(arr_result)) then
end if 'if ( sCStr(strPasswdMailCookie) = "true" ) then
%>
