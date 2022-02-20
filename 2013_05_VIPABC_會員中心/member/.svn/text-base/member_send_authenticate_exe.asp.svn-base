<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"-->
<%
Dim str_sql : str_sql="" 
Dim str_email_addr : str_email_addr = Trim(Request("get_email_addr"))

Dim str_cname : str_cname =""
Dim str_rand : str_rand =""
Dim int_valid : int_valid = 0
Dim j

'塞資料的相關設定值
Dim var_arr 						'傳excuteSqlStatementRead的陣列
Dim arr_result					'接回來的陣列值
'Dim str_tablecol 					'table欄位
'Dim str_tablevalue 				'table值
									
'亂數產生密碼(目前已10個字來處理)
Dim str_randpas : str_randpas = "0123456789abcdefghijklmnopqrstuvwxyzabcdefghijklmnopqrstuvwxyz"

'mail的部份
Dim arr_email_from_paras
Dim arr_email_to_paras
Dim arr_set_paras
Dim arr_opt_paras
Dim str_mail_subjcet : str_mail_subjcet = getWord("apply_authenticate_latter") 
Dim exec_result 
Dim str_mail_url '信件的url

dim strAuthorityMailCookie : strAuthorityMailCookie = Request.Cookies("authority_mail")

'20120909 阿捨新增 新增cookie避免重覆發信10分鐘
if ( "true" = sCStr(strAuthorityMailCookie) ) then
    Response.Write("请勿于10分钟内重复索取认证信, 以免造成无法成功认证的状况!")
else
	var_arr = Array(str_email_addr)
	arr_result = excuteSqlStatementRead("SELECT cname, rand_str, valid FROM client_basic WHERE email = @email AND (email <> '')", var_arr, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_result)) then
        '有資料
        if (Ubound(arr_result) >= 0) then  
		    str_cname = arr_result(0,0)
		    int_valid = arr_result(2,0)
 
		    '-----------帳號已認證過，無法再寄發認證信-------------- start----------
		    if (int_valid = 1 ) then 
			    'Call alertGo("此電子郵件帳號已開啟，無法再寄認證信", "", CONST_ALERT_THEN_GO_BACK)
	    	    Response.Write getWord("apply_authenticate_again") 
		    else 'if (int_valid = 1 ) then
			    '產生亂數字串
			    randomize
			    for j = 1 to 10
			       str_rand = str_rand & mid(str_randpas,int(62*rnd+1),1)
			    next
			
			    var_arr = Array(str_rand,str_email_addr)	
			    arr_result = excuteSqlStatementWrite("UPDATE client_basic SET rand_str = @rand_str WHERE email = @email", var_arr, CONST_VIPABC_RW_CONN)
	
			    if (arr_result > 0) then 
    '				'有更改到資料
    '				'----------------------發送註冊認證信-------------- start -----------
    '				str_mail_url = "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/mail_template/repeat_authenticate/repeat_authenticate.asp?mail_template_name="&Escape(str_cname)&"&mail_template_randstr="&str_rand
    '				arr_email_from_paras 	= Array(CONST_MAIL_ADDRESS_SERVICES_ABC,"",str_mail_subjcet,"","",str_mail_url)
    '				'arr_email_to_paras 	= Array("pipiyao@tutorabc.com",str_cname)
    '				arr_email_to_paras 	= Array(str_email_addr,str_cname)
    '				arr_set_paras 		= Array(CONST_MAIL_SERVER_CLIENT, CONST_EMAIL_CHARSET)
    '				arr_opt_paras			= Array(true)
    '				exec_result = sendEmail(arr_email_from_paras, arr_email_to_paras, arr_set_paras, arr_opt_paras, CONST_TUTORABC_WEBSITE)
    '				if (exec_result <> CONST_FUNC_EXE_SUCCESS) then
    '				' E-mail 錯誤訊息
    '					Response.Write(g_str_run_time_err_msg)
    '				end if
    '				'----------------------發送註冊認證信-------------- end ----------- 
    '				'Call alertGo("", "member_authenticate.asp", CONST_NOT_ALERT_THEN_REDIRECT ) 
                    Response.Write("success")
				    '20110214 tree VB10110205 VIPABC系統信件 --開始--
				    set SendMailvip = new VIPSystemSendMail			   '建立class物件
				    SendMailvip.str_sendlettermail_sn = "0007"			   '第0007封信 	主旨：申请邮箱认证信						
				    SendMailvip.strEmailToAddr = str_email_addr		   '客戶的電子郵件
				    SendMailvip.client_cname = str_cname					   '客戶姓名
				    SendMailvip.client_Account_new = str_email_addr '客戶帳號
				    SendMailvip.client_mail_Parameters = str_rand      '客戶電子郵件認證參數
				    Call SendMailvip.MainSendMail()                             '呼叫class物件的函式來發信
				    SendMailvip.successFailure
				    set SendMailvip = nothing                                         '釋放CLASS物件
				    '20110214 tree VB10110205 VIPABC系統信件 --結束--
                    
                    '20120909 阿捨新增 新增cookie避免重覆發信10分鐘
                    Response.Cookies("authority_mail") = "true"
	                Response.Cookies("authority_mail").Expires = DateAdd("n", 10, Now())   
			    else 'if (arr_result > 0) then
				    '無資料
			    end if 'if (arr_result > 0) then
		    end if 'if (int_valid = 1 ) then
	        '-----------帳號已認證過，無法再寄發認證信-------------- end ----------	
	    else 'if (Ubound(arr_result) >= 0) then
		    '無資料 (錯誤訊息)
		    'Call alertGo("資料錯誤 請輸入您註冊時的電子郵件帳號", "", CONST_ALERT_THEN_GO_BACK)
	       Response.Write getWord("DATA_ERROR_PLEASE_INPUT_EMAIL_ACCOUNT")
	    end if 'if (Ubound(arr_result) >= 0) then
    end if 'if (isSelectQuerySuccess(arr_result)) then
end if 'if ( "true" = sCStr(strAuthorityMailCookie) ) then
%>

