<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/functions_hash1way.asp" -->
<!--#include virtual="/lib/functions/functions_hexvalue.asp" -->
<!--#include virtual="/lib/functions/functions_chgkeyword.asp" -->
<!--#include virtual="/lib/include/testLog.inc"-->
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<%
dim bolDeBugMode : bolDeBugMode = false '除錯模式
dim bolTestCase : bolTestCase = false '測試模式

Dim var_arr '小幫手提供的conn開法所需變數
Dim arr_result '小幫手提供的conn開法所需變數
Dim int_row	 '小幫手提供的conn開法所需變數
Dim int_col '小幫手提供的conn開法所需變數
Dim str_tablecol '小幫手提供的conn開法所需變數
Dim str_tablevalue '小幫手提供的conn開法所需變數
Dim str_tmp_sql

Dim txt_login_account
Dim pwd_login_password
Dim chk_remember_account
Dim chk_remember_password

Dim chk_remember_account_Y : chk_remember_account_Y = "Y" '判斷從前端post過來是否要記憶帳號 寫入cookies  Y:是  (沒有勾選則會傳空值)
Dim chk_remember_password_Y : chk_remember_password_Y = "Y"	'判斷從前端post過來是否要記憶密碼 寫入cookies  Y:是  (沒有勾選則會傳空值)

Dim president : president = "n" '針對董事長開的後門
Dim str_password

Dim str_client_ip
Dim str_gobackurl
Dim str_tmp_go_page
Dim str_client_valid : str_client_valid = 0  '是否認證過


'VM12082820 好友推薦名單獲禮活動 - Constant接收該活動篇名，以利login後可直接導入活動頁 by Lily Lin 2012/10/08
dim strConstant : strConstant = getRequest("Constant", CONST_DECODE_NO)

'20120319 阿捨新增 香港測試模式
'dim intTestMode : intTestMode = getRequest("test_mode", CONST_DECODE_NO)
'session("test_mode") = sCInt(intTestMode)

'response.write strConstant

session("client_password") = Request("pwd_login_password")

'接POST參數
if (not isEmptyOrNull(Trim(Request("txt_login_account")))) then 
	txt_login_account = Trim(Request("txt_login_account"))
else
	'call sub alertGo
	Call alertGo(getWord("please_input_account"), "", CONST_ALERT_THEN_GO_BACK)
	Response.End
end if 

if (not isEmptyOrNull(Trim(Request("pwd_login_password")))) then 
	pwd_login_password = getDecode(Trim(Request("pwd_login_password")), CONST_DECODE_URL)
else
	'call sub alertGo
	Call alertGo(getWord("PWS_INPUT"), "", CONST_ALERT_THEN_GO_BACK)
	Response.End
end if 

chk_remember_account	= Trim(Request("chk_remember_account"))
chk_remember_password = Trim(Request("chk_remember_password"))

'判斷 记住我的Email账号 寫Cookies 
if (chk_remember_account = chk_remember_account_Y) then 
	Response.Cookies("remember_account") = txt_login_account
	Response.Cookies("remember_account").Expires	= DateAdd("m",1,Now())
else
	Response.Cookies("remember_account") = ""
	Response.Cookies("remember_account").Expires	= DateAdd("d",-1,Now())
end if

'判斷记住我的密码 寫Cookies 
if (chk_remember_password = chk_remember_password_Y) then 
	Response.Cookies("remember_password") = pwd_login_password
	Response.Cookies("remember_password").Expires = DateAdd("m",1,Now())
else
	Response.Cookies("remember_password") = ""
	Response.Cookies("remember_password").Expires = DateAdd("d",-1,Now())
end if

'針對董事長開的後門
if (Request("txt_username")="mingyang" and Request("txt_password")="president") then
	president="y"
	Session("client_fname")="Sliver"
	Session("client_fullname")="Sliver Chiu"
	Session("client_sn")=100291
	Session("client_email")="sliverchiu@smartabc.com"
	Session("client_nlevel")=6
	Session("client_type")=1
	Session("company_id")=237

    Response.Cookies("everlogin") = 100291
    Response.Cookies("everlogin").Expires = dateadd("yyyy",15,date)
    Response.Cookies("everlogin").domain = "vipabc.com"

	'call sub alertGo
	Call alertGo("", "/asp/member/member_services.asp", CONST_NOT_ALERT_THEN_REDIRECT)
	Response.End
end if

if (Request("go")="1") then 
	'call sub alertGo
	Call alertGo("", "/asp/member/fisat_session_member.asp", CONST_NOT_ALERT_THEN_REDIRECT)
	Response.End
end if
   
var_arr = Array(txt_login_account)
str_tmp_sql = "select client_basic.sn, client_basic.fname,"  '0 1
str_tmp_sql = str_tmp_sql & "client_basic.nlevel, client_basic.lname, "  '2 3
str_tmp_sql = str_tmp_sql & "client_basic.sex,client_basic.olevel, "  '4 5
str_tmp_sql = str_tmp_sql & "client_basic.email,client_basic.valid, "  ' 6 7
str_tmp_sql = str_tmp_sql & "client_basic.company_id, client_basic.industry, " ' 8 9
str_tmp_sql = str_tmp_sql & "client_basic.role, client_basic.status, client_basic.client_type, " ' 10 11 12
str_tmp_sql = str_tmp_sql & "client_passwd.passwd, client_passwd.salt, "  '13 14
str_tmp_sql = str_tmp_sql & "client_passwd.client_user, vip_session_valid, client_basic.web_site, client_basic.cname, " '15 16 17 18
str_tmp_sql = str_tmp_sql & "client_basic.industry , client_basic.role , client_basic.interest , client_basic.Focus "  '19 20 21 22 判斷是否做過DCGS
str_tmp_sql = str_tmp_sql & "from client_basic(nolock) "
str_tmp_sql = str_tmp_sql & "inner join  client_passwd(nolock) on client_basic.sn = client_passwd.client_sn "
str_tmp_sql = str_tmp_sql & "where email=@email "
 
arr_result = excuteSqlStatementRead(str_tmp_sql, var_arr, CONST_VIPABC_RW_CONN)

if (isSelectQuerySuccess(arr_result)) then
    '有資料
    if (Ubound(arr_result) > 0) then
        'TODO: 只允許大陸人和RD的人 登入
        'if (Trim(arr_result(17,0)) <> "C" _
        '    and txt_login_account <> "pipiyao@tutorabc.com" _
        '    and txt_login_account <> "denilechen@tutorabc.com" _
        '    and txt_login_account <> "raychien@tutorabc.com" _
        '    and txt_login_account <> "deneyhsiao@tutorabc.com" _
        '    and txt_login_account <> "shannonlin@tutorabc.com" _
        '    and txt_login_account <> "aboochen@tutorabc.com" _
        '    and txt_login_account <> "summerchang@tutorabc.com" _
        '    ) then
	    '        Call alertGo(getWord("data_error_please_input_email_account_and_password"), "", CONST_ALERT_THEN_GO_BACK)
	    '        Response.End
        'end if
        if ( Instr(txt_login_account, "@vipabc.com") > 0 OR Instr(txt_login_account, "@tutorabc.com") > 0 ) then
            Session("test_staff") = true	
        else
            Session("test_staff") = false	
        end if
    else
	    Call alertGo(getWord("data_error_please_input_email_account_and_password"), "", CONST_ALERT_THEN_GO_BACK)
	    Response.End
    end if
end if

'g_arr_result(0,0) 	'client_basic.sn
'g_arr_result(1,0) 	'client_basic.fname
'g_arr_result(2,0) 	'client_basic.nlevel
'g_arr_result(3,0) 	'client_basic.lname
'g_arr_result(4,0) 	'client_basic.sex
'g_arr_result(5,0) 	'client_basic.olevel
'g_arr_result(6,0) 	'client_basic.email
'g_arr_result(7,0) 	'client_basic.valid
'g_arr_result(8,0) 	'client_basic.company_id
'g_arr_result(9,0) 	'client_basic.industry
'g_arr_result(10,0) 'client_basic.role
'g_arr_result(11,0) 'client_basic.status
'g_arr_result(12,0) 'client_basic.client_type
'g_arr_result(13,0) 'client_passwd.passwd
'g_arr_result(14,0) 'client_passwd.salt
'g_arr_result(15,0) 'client_passwd.client_user
'g_arr_result(16,0) 'vip_session_valid

'進行資料庫的密碼比對開始
if (arr_result(13,0)<>"" and arr_result(14,0)<>"") then
	str_password = Replace(Replace(Trim(pwd_login_password),"'",""),"""","") & arr_result(14,0)
	str_password = hashencode(str_password)
	if (str_password = arr_result(13,0)) then					                    
		'產生session
		Session("client_cname")	 = gkeyword(arr_result(18,0))
		Session("client_fname") = gkeyword(arr_result(1,0))
		Session("client_fullname") = gkeyword(arr_result(1,0)) & " "  & gkeyword(arr_result(3,0))
		Session("client_sn") = arr_result(0,0) 
        int_client_sn = arr_result(0,0) 
		Session("client_email") = arr_result(6,0)
		Session("client_nlevel") = arr_result(2,0)
		Session("client_type")	 = arr_result(12,0)
		'新增客戶狀態  1:已買  0:免費
		Session("client_status") = arr_result(11,0)
		'20051004新增將公司id紀錄到session值.
		Session("company_id") = arr_result(8,0)
		Session("search_vip_session_str") = arr_result(16,0)
		'''新增語言程式分析 及 DCGS 的 session Johnny 2012-08-16
		if(isEmptyOrNull(arr_result(2,0)) and isEmptyOrNull(arr_result(5,0))) then
			session("bolComplateLevelTest") = "0"
		else
			session("bolComplateLevelTest") = "1"
		end if
        
        Dim str_client_sn : str_client_sn = session("client_sn")
		g_var_client_sn = str_client_sn
		%>
        <!--#include virtual="/program/member/getJRClient.asp"-->
        <%
        '20120320 阿捨新增 VIPABCJR客戶判斷 bolVIPABCJR
        if ( true = bolDebugMode ) then
            response.Write "bolVIPABCJR: " & bolVIPABCJR & "<br />"
        end if
		
        if ( true = bolVIPABCJR ) then
            bolkompletDcgs = isDCGSClient(str_client_sn, 0, CONST_VIPABC_RW_CONN)
            if ( true = bolDeBugMode ) then
                response.Write "bolkompletDcgs: " & bolkompletDcgs & "<br/>"
                'response.end
            end if
        else
		    if ( (not isEmptyOrNull(arr_result(19,0))) AND (not isEmptyOrNull(arr_result(20,0))) AND (not isEmptyOrNull(arr_result(21,0))) AND (not isEmptyOrNull(arr_result(22,0))) ) then
			    bolkompletDcgs = true 'DCGS第1~4個頁面 有值表示全部設定完畢過
		    end if
        end if
		if ( true = bolkompletDcgs ) then
			session("bolkompletDcgs") = "1" '全部設定完畢過
		else
			session("bolkompletDcgs") = "2" 'DCGS資料有缺
		end if
		
        Response.Cookies("everlogin") = arr_result(0,0)
        Response.Cookies("everlogin").Expires = dateadd("yyyy",15,date)
        Response.Cookies("everlogin").domain = "vipabc.com"

		
		'*******************2013/02/27 Fox 判斷是否為有過有效合約的客戶，提供官網登入後右方Banner判斷使用********
		'str_sql = " select account_sn from client_purchase (nolock) where client_sn = @client_sn and valid = 1 and product_sdate <= GETDATE() and purchase_status in ('短期方案客戶','正常方案客戶')"
		'Dim arr_sql_result_f 
		'arr_sql_result_f = excuteSqlStatementRead(str_sql,Array(session("client_sn")),CONST_TUTORABC_R_CONN)
		'if (isSelectQuerySuccess(arr_sql_result_f)) then
		'	if (Ubound(arr_sql_result_f) >= 0) then
		'		session("bol_was_had_valid_contract") = "1" '有過有效合約
		'	end if
		'end if
		'********************************************************************************************************

		'*******************2013/04/01 Fox 判斷是否3/16-5/15续约活动banner EDM 短信客戶，提供官網登入後右方Banner判斷使用********
		'此TABLE由發信用的TABLE轉移到官網資料庫提供名單查詢，inp_date = "2013/04/15"
		str_sql = " select client_sn from temp_edm_fetch_roster (nolock) where client_sn = @client_sn and inp_date = @inp_date"
		Dim arr_sql_result_P 
		arr_sql_result_P = excuteSqlStatementRead(str_sql, Array(session("client_sn"), "2013/04/01"), CONST_TUTORABC_R_CONN)
		if (isSelectQuerySuccess(arr_sql_result_P)) then
			if (Ubound(arr_sql_result_P) >= 0) then
				session("PromotionsforU") = "1" '续约活动
			end if
		end if
		'********************************************************************************************************
		
		'*******************2013/04/15 Fox 判斷是否4/7-5/15banner EDM 客戶，提供官網登入後右方Banner判斷使用********
		'此TABLE由發信用的TABLE轉移到官網資料庫提供名單查詢，inp_date = "2013/04/15"
		str_sql = " select client_sn from temp_edm_fetch_roster (nolock) where client_sn = @client_sn and inp_date = @inp_date"
		Dim arr_sql_result_J 
		arr_sql_result_J = excuteSqlStatementRead(str_sql, Array(session("client_sn"), "2013/04/15"), CONST_TUTORABC_R_CONN)
		if (isSelectQuerySuccess(arr_sql_result_J)) then
			if (Ubound(arr_sql_result_J) >= 0) then
				session("JuniorMGM") = "1" 'JuniorMGM活动
			end if
		end if
		'********************************************************************************************************

		'*******************2013/05/15 Fox 判斷是否至5/31banner EDM 客戶，提供官網登入後右方Banner判斷使用********
		'此TABLE由發信用的TABLE轉移到官網資料庫提供名單查詢，inp_date = "2013/04/15"
		str_sql = " select client_sn from temp_edm_fetch_roster (nolock) where client_sn = @client_sn and inp_date = @inp_date"
		Dim arr_sql_result_S 
		arr_sql_result_S = excuteSqlStatementRead(str_sql, Array(session("client_sn"), "2013/05/15"), CONST_TUTORABC_R_CONN)
		if (isSelectQuerySuccess(arr_sql_result_S)) then
			if (Ubound(arr_sql_result_S) >= 0) then
				session("ShareMGM") = "1" '与好友分享英语变好的秘密 共享五月咖啡时光 MGM活动
			end if
		end if
		'********************************************************************************************************
		
		'-------------------2012/08/21 Johnny 從ABC官網移植過來 首測重測登入頁面判斷導入 start-------------------
		Dim str_is_firstsession
		'sql所需要的參數
		Dim arr_parameter_data 
		'查詢出客戶課互有沒有合約，的結果
		Dim arr_sql_result 
		'是否有有效合約
		Dim bol_is_had_valid_contract : bol_is_had_valid_contract = false 
		
		dim bolFirstTestQuestionaryPassed : bolFirstTestQuestionaryPassed = false
		'**** 客戶開通後才能預約 (Sabrina 例外) START ****
		str_sql = " select account_sn " &_
				  " from client_purchase (nolock) " &_
				  " where client_sn = @client_sn and valid = 1" &_
				   " and (CONVERT(varchar, GETDATE(), 120) BETWEEN appdate AND product_edate) " '這句補加在合約時間內  Johnny 2012-10-11 等待確認是否會有其他影響

		arr_parameter_data = Array(str_client_sn)
		arr_sql_result = excuteSqlStatementRead(str_sql,arr_parameter_data,CONST_TUTORABC_R_CONN)
		if (isSelectQuerySuccess(arr_sql_result)) then
			if (Ubound(arr_sql_result) >= 0) then
				bol_is_had_valid_contract = true 
				session("bol_is_had_valid_contract") = "1" '是否有有效合約
			end if
		end if
		
		if (true = bol_is_had_valid_contract) then
			Dim arr_chk_data 'function getFirstTestQuestionaryPassedStatus 傳入參數
			Redim arr_chk_data(2)
			arr_chk_data(0)=1 '首測問卷
			arr_chk_data(1)="T" 'tutorABC
			arr_chk_data(2)="tw" '語系

			'1: 通過 2: 沒通過
			if (not 2 = getFirstTestQuestionaryPassedStatus(session("client_sn"),arr_chk_data,CONST_TUTORABC_R_CONN)) then
				bolFirstTestQuestionaryPassed = true
			end if
		end if
		
		Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")  
		set envirService = new EnvironmentTestService
		Dim haveTestOk : haveTestOk = envirService.checkEnvironmentHaveTestOk(objPCInfoInsert)
		set envirService = nothing
		
		'session("bolFirstTestQuestionaryhaveTestOk") 電腦測試是否成功 有 Y E N幾種狀態	by Johnny
		'session("bolFirstTestQuestionaryhaveTestNOTOkOrder") 是否預約 有Y N	by Johnny
		session("bolFirstTestQuestionaryhaveTestOk") = haveTestOk  '有 Y E N幾種狀態	
		'它如果沒有做過問卷 就當做電腦測試不通過，讓它不能訂課
		'my_sql = "select * from add_client_datetime where client_sn=@g_var_client_sn and flag=4 and testflag=1 " 
		my_sql = " select * from  add_client_datetime ac WITH (nolock) LEFT JOIN dbo.questionary_fill_record qf WITH (NOLOCK) ON ac.sn = qf.add_client_datetime_sn   where ac.client_sn = @intclient_sn and ac.status not in (5)  and ac.description <> N'客戶於VIP官網新手上路自行填寫問卷' AND (qf.qfr_passed<>'2' OR qf.qfr_passed IS NULL) "

        call objPCInfoInsert.excuteSqlStatementEach(my_sql, Array( str_client_sn ) , CONST_VIPABC_RW_CONN)   
		if objPCInfoInsert.RecordCount > 0 then
			session("bolFirstTestQuestionaryhaveTestNOTOkOrder") = "Y"
		else
			session("bolFirstTestQuestionaryhaveTestNOTOkOrder") = "N"
		end if

		if ( bolFirstTestQuestionaryPassed = true ) then
			session("bolFirstTestQuestionaryPassed") = "1" '是否通過首冊1通過  2不通過
		else
			session("bolFirstTestQuestionaryPassed") = "2" '是否通過首冊1通過  2不通過
		end if
		
		'' 增加鬆綁條件，如果客戶有上過課，則視為首測重測作過 2012-11-23
		strSql =    getSqlNotes() &_  
				"   SELECT		TOP (1) sn                                      "&_
				"   FROM		client_attend_list                              "&_
				"   WHERE		(client_sn = @intclient_sn)                     "&_
				"               and session_sn is not null                      "&_
				"               and valid = '1'                                 "&_
				"               and attend_date < getDate()    "
		intSqlWriteResult = objPCInfoInsert.excuteSqlStatementEach(strSql, Array(str_client_sn), CONST_VIPABC_RW_CONN)     '有傳入參數用
		if ( objPCInfoInsert.RecordCount > 0 ) then
			session("boleverclass") = "1" '有客戶曾經上過課
		else
			session("boleverclass") = "2" '無客戶曾經上過課
		end if
	
		objPCInfoInsert.close()
		set objPCInfoInsert =nothing
			
		str_client_valid = arr_result(7,0) '是否認證過
		str_client_ip = getIpAddress()  
		str_tablecol = "client_sn,email,password,client_ip"
		str_tablevalue = "@client_sn,@email,@password,@client_ip"
		var_arr = Array(arr_result(0,0),arr_result(6,0),Replace(Replace(Trim(pwd_login_password),"'",""),"""",""),str_client_ip)
		arr_result = excuteSqlStatementWrite("Insert into client_password_input ("&str_tablecol&") values("&str_tablevalue&")" , var_arr, CONST_VIPABC_RW_CONN)

        Response.Cookies("client_sn") = Session("client_sn")
	    Response.Cookies("client_sn").Expires = DateAdd("d", 1, Now())
        '20110614 阿捨新增 特殊客戶關懷
        '20120926 阿捨新增 特殊客戶關懷include檔
        dim intNoifiyType : intNoifiyType = ""
        intNoifiyType = 1
%>
<!--#include virtual="/program/member/CareClientNotification.asp"-->
<%
		'----- 判斷此客戶是否認證過email -------- start -------
		if ( str_client_valid = 0 or cint(str_client_valid = 0) ) then
			'客戶尚未透過email啟動帳號,則系統不開放使用功能
			Call alertGo("", "member_login.asp", CONST_NOT_ALERT_THEN_REDIRECT )
			'close_resource()
			response.end
		end if
		'----- 判斷此客戶是否認證過email -------- end -------		
		
		'------------------找出客戶的productno ------- start -------------					
		'取出合約
		Dim int_purchase_got_flag : int_purchase_got_flag = 0	'是否已設定合約編號的session 1:yes 0:no
		Dim str_product_sql : str_product_sql = ""
		Dim int_i : int_i = ""
		str_product_sql = " SELECT product_sn, res_type, product_sdate, product_edate, purchase_status "
		str_product_sql = str_product_sql & " FROM client_purchase "
		str_product_sql = str_product_sql & " WHERE (client_sn =@client_sn) AND (Valid = @Valid)  "
		str_product_sql = str_product_sql & " order by account_sn ASC "
		var_arr = Array(Session("client_sn"),1)
		arr_result = excuteSqlStatementRead(str_product_sql, var_arr, CONST_VIPABC_RW_CONN)

		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then
				for int_i = 0 to Ubound(arr_result,2)
					If ( int_purchase_got_flag = 0 And IsDate(arr_result(2,int_i)) And IsDate(arr_result(3,int_i)) ) Then
						If ((DateDiff("d",DateValue(arr_result(2,int_i)),DateValue(now())) >= 0) And (DateDiff("d",DateValue(arr_result(3,int_i)),DateValue(now())) <= 0)) then 
							session("product_no")=arr_result(0,int_i) 'product_sn
							int_purchase_got_flag = 1
						end if 
					end if 
					If ( int_purchase_got_flag = 0 ) Then
						session("product_no")=arr_result(0,int_i) 'product_sn
					End If
					session("purchase_status")=arr_result(4,int_i) 'purchase_status
				next
			else
			'無資料
			end if 
		end if 
		'------------------找出客戶的productno ------- end -------------	
		
		'------------------------------------2012/12/11   VM12112904當會員已成為客戶後，官網的在線直播講堂屏蔽不要顯示  jessica ----------------------------------------
		Dim str_purchase_sql : str_purchase_sql = ""
		Dim str_purchase_sql_j: str_purchase_sql_j=""
		Dim int_j : int_j = ""
		DIM type_j: type_j=""
		str_purchase_sql = " SELECT product_sn, res_type, product_sdate, product_edate, purchase_status "
		str_purchase_sql = str_purchase_sql & " FROM client_purchase "
		str_purchase_sql = str_purchase_sql & " WHERE (client_sn =@client_sn) AND (Valid = @Valid)  "
		str_purchase_sql = str_purchase_sql & " order by account_sn ASC "
		var_arr = Array(Session("client_sn"),1)
		arr_result = excuteSqlStatementRead(str_purchase_sql, var_arr, CONST_VIPABC_RW_CONN)
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then
				for int_j = 0 to Ubound(arr_result,2)
					If ( int_purchase_got_flag = 0 And IsDate(arr_result(2,int_j)) And IsDate(arr_result(3,int_j)) ) Then
						If ((DateDiff("d",DateValue(arr_result(2,int_j)),DateValue(now())) >= 0) And (DateDiff("d",DateValue(arr_result(3,int_j)),DateValue(now())) <= 0)) then 
							session("product_no")=arr_result(0,int_j) 'product_sn
							int_purchase_got_flag = 1
						end if 
					end if 
					If ( int_purchase_got_flag = 0 ) Then
						session("product_no")=arr_result(0,int_j) 'product_sn
					End If
						session("purchase_status")=arr_result(4,int_j) 'purchase_status
				next
			else
			'無資料
			end if 
		end if 
		if(session("purchase_status")="企業用戶客戶" and session("purchase_status")="員工") then
			session("menu_expo_attend_session")="N"
		else
			str_purchase_sql_j=" select a.sn,a.cname,a.renew,b.typename "
			str_purchase_sql_j=str_purchase_sql_j & " from cfg_product a with (nolock) inner join cfgproducttype b with (nolock) "
			str_purchase_sql_j=str_purchase_sql_j & " on a.product_status=b.sn "
			str_purchase_sql_j=str_purchase_sql_j & " where a.sn=@sn" 
			var_arr = Array(Session("product_no"),1)
			arr_result = excuteSqlStatementRead(str_purchase_sql_j, var_arr, CONST_VIPABC_RW_CONN)
			if (isSelectQuerySuccess(arr_result)) then
				if(Ubound(arr_result) >= 0) then
					type_j=arr_result(3,0)
				end if
			end if
			if(type_j="定時定額" or type_j="無限卡" or type_j="終身無限卡" or type_j="一般年約")then
				session("menu_expo_attend_session")="N"
			end if
		end if
		'-----------------------------end 2012/12/11  VM12112904當會員已成為客戶後，官網的在線直播講堂屏蔽不要顯示    jessica ----------------------
		
if ( true = bolDeBugMode ) then
else
		'================================================================
		'判斷是否要先導頁面至點選合約同意頁面 Begin
		Dim str_contract_sql
		dim clientsn_contract_sn : clientsn_contract_sn = ""
		dim contract_cagree : contract_cagree = ""
		dim contract_ctype : contract_ctype = ""
		dim contract_ptype : contract_ptype = ""
		dim contract_pmethod : contract_pmethod = ""
		
		Session("contract_sn") = ""
		'20130429 Beson 修改SQL, 續約才會取得無確認合約
		str_contract_sql="SELECT top 1 sn, cagree, ctype, ptype, pmethod FROM client_temporal_contract WHERE (client_sn = @client_sn) AND (valid=1) order by sn desc"
		
		var_arr = Array(Session("client_sn"))
		arr_result = excuteSqlStatementRead(str_contract_sql,var_arr,CONST_VIPABC_RW_CONN)
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then
				'Response.write("欄:" & Ubound(arr_result) & "<br>")
				'Response.write("列:" & Ubound(arr_result,2) & "<br>")
				clientsn_contract_sn = arr_result(0, 0)	
				contract_cagree = arr_result(1, 0)
				contract_ctype = arr_result(2, 0)
				contract_ptype = arr_result(3, 0)
				contract_pmethod = arr_result(4, 0)
				if (contract_ctype = "1") then
				    Session("direct_contract_sn") = "Y" '一次性付款
				    'Response.write("Session(contract_sn) = " & Session("contract_sn") & "<br>")
				    if (contract_cagree = "0") then   '尚未同意電子合約 
				    	Session("contract_sn") = arr_result(0, 0)
				    end if
				end if
			else
				'Response.write "no data!!"
			end if
		end if

		set str_contract_sql = nothing
		set var_arr = nothing
		set arr_result = nothing
		'判斷是否要先導頁面至點選合約同意頁面 End
		'================================================================
		
		'******如果客户有预约demo且业务已为其开教室，登录后跳转至进入试读教室页面 VM13040315 2013/4/10 paul begin
		Dim bol_classroom : bol_classroom = false
		Dim str_sql : str_sql = ""
		Dim arr_client_result : arr_client_result = ""
		str_sql = "SELECT  a.room "
		str_sql = str_sql & "FROM  dbo.client_installation a WITH (NOLOCK) "
		str_sql = str_sql & "LEFT JOIN dbo.lead_basic b WITH (NOLOCK) ON a.client_sn = b.sn "
		str_sql = str_sql & "LEFT JOIN dbo.client_record c WITH(NOLOCK) ON a.client_sn=c.client_sn "
		str_sql = str_sql & "WHERE  b.client_sn = @sn "
		str_sql = str_sql & "AND c.inp_time BETWEEN CONVERT(VARCHAR(10), GETDATE(), 120) + ' 00:00:00' "
		str_sql = str_sql & "AND CONVERT(VARCHAR(10), GETDATE(), 120) + ' 23:59:59' "  
		str_sql = str_sql & "AND room > '' "
		str_sql = str_sql & "AND c.valid = 1 "
		str_sql = str_sql & "AND c.status IN (6,7) "
		arr_client_result = excuteSqlStatementRead(str_sql,Array(session("client_sn")),CONST_VIPABC_RW_CONN)
		if (isSelectQuerySuccess(arr_client_result)) then
			if (Ubound(arr_client_result) >= 0) then
				bol_classroom = true
			end if
		end if
		'******如果客户有预约demo且业务已为其开教室，登录后跳转至进入试读教室页面 VM13040315 2013/4/10 paul end
		
		'若需要線上刷卡導向刷卡頁面 ---------------- start --------------
		Dim str_pay_contract_sn : str_pay_contract_sn = "" '線上刷卡的合約編號 
		str_pay_contract_sn = getClientBrushCardSn(Session("client_sn"),"client")
		'若需要線上刷卡導向刷卡頁面 ---------------- end --------------

		'20130412 Fox ACH導頁至刷入卡號，client_temporal_contract欄位ptype為3則是ACH，要合約已確認才會成立***********************************
        'response.write "client_sn = " & Session("client_sn") & "<br />"
		'response.write "int_contract_sn = " & clientsn_contract_sn & "<br />"
		if (contract_ptype = "3" and contract_pmethod = "48" and contract_cagree = "1") then
		    Dim sqlcardtype : sqlcardtype = ""
            Dim arrParameterforcardtype
            
			'首先判斷合約是否已經開通，開通則不執行，未開通則設定session("ach_cvs_input") = "Y"
			arrParameterforcardtype = Array(Session("client_sn"), clientsn_contract_sn)
		    sqlcardtype = " SELECT account_sn FROM client_purchase WITH (NOLOCK) "
            sqlcardtype = sqlcardtype & " WHERE client_sn = @client_sn and contract_sn = @contract_sn "
            
			session("ach_cvs_input") = ""
		    'response.write "int_contract_sn = " & clientsn_contract_sn & "<br />"
		    arr_result = excuteSqlStatementRead(sqlcardtype, arrParameterforcardtype, CONST_VIPABC_RW_CONN)
		    if (isSelectQuerySuccess(arr_result)) then
		    	if (Ubound(arr_result) >= 0) then
				    if (Session("client_sn") = "462488") then
					    session("ach_cvs_input") = "Y"	
					    session("ach_contract_sn") = clientsn_contract_sn
					end if
				else '合約未開通
		    	    session("ach_cvs_input") = "Y"	
					session("ach_contract_sn") = clientsn_contract_sn
					arrParameterforcardtype = Array(clientsn_contract_sn, 1)
		            sqlcardtype = " SELECT Cwthorizationcode FROM VIPABC_ACH_CARD_INFO_FIRST_PART WITH (NOLOCK) "
                    sqlcardtype = sqlcardtype & " WHERE ContractSn = @ContractSn and valid = @valid "
            
		            'response.write "int_contract_sn = " & clientsn_contract_sn & "<br />"
		            arr_result = excuteSqlStatementRead(sqlcardtype, arrParameterforcardtype, CONST_VIPABC_RW_CONN)
		            if (isSelectQuerySuccess(arr_result)) then
		    	        if (Ubound(arr_result) >= 0) then
		    	            'response.write "Cwthorizationcode = " & arr_result(0, 0) & "<br />"
                            if (isEmptyOrNull(arr_result(0, 0))) then '空的未填CVS要轉頁
                                response.redirect ("CardExpireDateCheck.asp?contract_sn=" & clientsn_contract_sn)
		    	            end if 
                        else
                            response.redirect ("CardExpireDateCheck.asp?contract_sn=" & clientsn_contract_sn)			
		    	        end if
		            end if
		    	end if
		    end if
		end if
        '**************************************************************************************************************

		'判斷轉址頁面開始=====================================================================================================
					
	    '**************推薦好友活動，判斷登入後導致活動頁面，由參數mgm決定篇名 -  by fox 2013/01/06**********
		if (request("mgm")<>"") then
			  if (request("mgm")="20130307") then '倒数10天 呼朋引伴限时抢现金
				  response.redirect ("/program/mail_template/MgmEDM/Wanted/index.asp")
		      elseif (request("mgm")="20130109") then '推荐好友上VIPABC 惊喜现金翻倍送
				  response.redirect ("/program/mail_template/MgmEDM/DoubleGain/index.asp")
		      elseif (request("mgm")="20130319") then '3月尊礼特权 职场英语好礼免费送亲友
				  response.redirect ("/program/mail_template/MgmEDM/WhiteCard/index.asp")
		      elseif (request("mgm")="20130423") then '妈妈最放心品牌- VIPABC让孩子轻松开口说英语
				  response.redirect ("/program/mail_template/MgmEDM/JuniorMGM/index.asp")
		      elseif (request("mgm")="20130513") then '与好友分享英语变好的秘密 共享五月咖啡时光
				  response.redirect ("/program/mail_template/MgmEDM/ShareToFriend/index.asp")				  
			  else '推荐亲友一起VIPABC 领取最高3000元现金回馈
				  response.redirect ("/program/mail_template/MgmEDM/cashfeedback/Recommend.asp")
              end if			  
		end if		
		'****************************************************************************************************************
						
	    '**************廣告跳轉訂課頁面  by fox  2013/04/01**********
		if (request("edm")<>"") then
            if (request("edm")="edm20130104" or request("edm")="edm20130401") then '連接至大會堂訂課頁面
                response.redirect ("/program/member/reservation_class/lobby_near_order.asp")
            else 
  			    response.redirect ("/program/member/reservation_class/class_choose.asp")
            end if
		end if		
		'****************************************************************************************************************
	
        if ( true = bolSpecialSpyCare ) then
            Session.abandon	
            Response.Cookies("stupid_spy") = "faking"
	        Response.Cookies("stupid_spy").Expires = DateAdd("Y", 1, Now())
		    call delayFrame(90, "/program/fix_incident.asp")
            response.end
		end if

	   'eddieliu 201000513	判斷今天是否有訂世博大會堂，有則在20分鐘前登入後導至進入教室頁面
		Dim expobooking_sql, str_attend_date
		'ryanlin 20100519 新增世博大會堂判斷-----start-----
		'---客戶資料 COM 物件
		Dim bol_is_customer '是否為客戶
		Dim obj_member_opt_exe : Set obj_member_opt_exe = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
		Dim int_member_result_exe : int_member_result_exe = obj_member_opt_exe.prepareData(int_client_sn, CONST_VIPABC_RW_CONN)

		if (int_member_result_exe = CONST_FUNC_EXE_SUCCESS) then        
			bol_is_customer = obj_member_opt_exe.getData(CONST_CUSTOMER_IS_HAVE_CONTRACT)
		else
			'錯誤訊息
			Response.Write "member prepareData error = " & obj_member_opt_exe.PROCESS_STATE_RESULT & "<br>"
		end if
		expobooking_sql = "select attend_date from free_member_attend_list where valid =1 and client_sn ="&int_client_sn

		arr_result = excuteSqlStatementReadQuick(expobooking_sql,CONST_VIPABC_R_CONN)
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then
				str_attend_date = arr_result(0,0)
				'如果是客戶
				if ( true = bol_is_customer ) then
					'且剛好那天訂了免費大會堂，導至訂課頁面
					if DateDiff("d", now(), str_attend_date) = 0 then
						response.redirect ("/program/member/reservation_class/expo_attend_session.asp")
						response.end
                    else
                        '20120319 阿捨新增 為VIPABCJR客戶 導頁至.net
                        if ( true = bolVIPABCJR ) then
                            'response.redirect ("/program/aspx/AspxLogin.asp?url_id=1")
                            response.redirect ("/program/member/class.asp")
                        else
                            response.redirect ("/program/member/member.asp")
						end if
                        response.end
					end if	
				else
				'如果不是客戶
					'如果還沒到訂課得時間
					if DateDiff("d", now(), str_attend_date) >= 0 then
						response.redirect ("/program/member/reservation_class/expo_attend_session.asp")
						response.end
					end if	
				end if
			end if 
		end if		
		'ryanlin 20100519 新增世博大會堂判斷-----end-----
		
		if ( Session("contract_sn") <> "" ) then 	'確認合約
			str_tmp_go_page =  "contract_audit.asp"
		elseif ( bol_classroom ) then		'进入试读教室
			str_tmp_go_page = "/program/classroom/demo_online.asp"
		elseif ( str_pay_contract_sn <> 0 ) then 
			str_tmp_go_page =  "buy_1.asp"
		elseif ( Request("lobbyID") <> "" ) then 	'pipi 20090423 未登入會員訂位大會堂
			str_tmp_go_page =  "/asp/member/reservation_lobby_old_exe.asp?loglob="& Request("lobbyID")
		elseif ( Request("gobackurl")<>"" or Request("backURL")<>"" or Request("education")<>"" ) then '判斷客戶原本有指定要轉址到其中的一個頁面
			if ( Request("gobackurl")<>"" ) then
				If( Request("gobackurl")="back_search" ) Then
					Response.Write ("<script language='javascript'>history.go(-2);</script>")
					Response.End
				Else
					str_gobackurl = Request("gobackurl")
				End If
			elseif ( Request("backURL")<>"" ) then 
				str_gobackurl = Request("backURL")
			end if
		
			if str_gobackurl="testquestions/start.asp" then 
				'加開 多益測試頁
				response.Write "<Script>window.open('http://" & abc_webhost_name  & "/" & str_gobackurl & "');</script>"							
			else
				str_tmp_go_page =  "/" & str_gobackurl 
			end if

            if ( session("exporURL") = "1" ) then 'eddieliu 20100504 加入判斷如果是從世博頁面來的登入後即導入訂課頁面
			    response.redirect ("/program/member/reservation_class/class_choose.asp")
		    end if
		'VM12082820 好友推薦名單獲禮活動 - Constant接收該活動篇名，以利login後可直接導入活動頁 by Lily Lin 2012/10/08
		elseif ( strConstant = "friendAwardProj") then
			response.redirect ("/program/mail_template/MgmEDM/friendAwardProj/tj_friend3.asp")
		else	'什麼都沒有 登入完導回首頁
            '20120319 阿捨新增 為VIPABCJR客戶 導頁至.net
            if ( true = bolVIPABCJR ) then
                'str_tmp_go_page = "/program/aspx/AspxLogin.asp?url_id=1"
				str_tmp_go_page = "/program/member/class.asp"
            else
				str_tmp_go_page = "/program/member/member.asp"
            end if
		end if	
		
        '登入頁面
        if (isEmptyOrNull(Trim(Request("popup_login_member")))) then
		    'call sub alertGo
		    Call alertGo("", str_tmp_go_page, CONST_NOT_ALERT_THEN_REDIRECT)
        '從popup login member 視窗做登入的動作 不轉址
        else
            Response.Clear
            Response.Write "success"
        end if
		Response.End
		'判斷轉址頁面結束=====================================================================================================
end if 'if ( bolDeBugMode ) then
	else '密碼錯誤
	    Call alertGo(getWord("data_error_please_input_email_account_and_password"), "", CONST_ALERT_THEN_GO_BACK)
	    Response.End
	end if
end if
%>