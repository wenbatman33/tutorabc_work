<%
Dim int_client_purchase_product_type : int_client_purchase_product_type = 1         '客戶購買合約的產品類型 1:一般產品 2:定時定額
Dim dat_now_week_account_start_date : dat_now_week_account_start_date = ""          '本週的週期初開始日
Dim dat_now_week_account_end_date : dat_now_week_account_end_date = ""              '本週的週期末結束日 (結算日)
Dim flt_client_total_session : flt_client_total_session = null                      '客戶總剩餘堂數
Dim flt_client_week_session : flt_client_week_session = 0                           '客戶當週剩餘的堂數
Dim flt_client_bonus_session : flt_client_bonus_session = 0                         '客戶贈送堂數(紅利包)
Dim str_service_start_date : str_service_start_date = ""                            '客戶目前合約的服務開始日
Dim str_service_end_date : str_service_end_date = ""                                '客戶目前合約的服務結束日
Dim bol_client_contract_edate_valid : bol_client_contract_edate_valid = false       '客戶合約是否已到期
Dim str_now_use_contract_sn : str_now_use_contract_sn = null                        '目前電子合約序號
Dim bol_client_have_valid_contract : bol_client_have_valid_contract = false         '客戶是否有有效合約
Dim str_last_first_session_datetime : str_last_first_session_datetime = ""          '客戶最近一次預約First Session Datetime (首次上課客戶的)
Dim str_now_use_product_sn : str_now_use_product_sn = null                          '目前使用的產品編號
Dim str_purchase_account_sn : str_purchase_account_sn = ""                          '已經開始還未結束的於PURCHASE的序號
Dim arr_msg_no_order_in_hour(2)                                                     '(0)一對一和(1)一對三和(2)大會堂幾小時內不能訂課 訊息
Dim arr_msg_no_cancel_in_hour(2)                                                    '(0)一對一和(1)一對三和(2)大會堂幾小時內不能取消 訊息
Dim arr_no_cancel_in_minute(2)                                                      '(0)一對一和(1)一對三和(2)大會堂幾分鐘內不能取消 minute
Dim str_client_full_ename : str_client_full_ename = ""                              '客戶英文名字
Dim str_client_login_email_addr : str_client_login_email_addr = ""                  '客戶登入的帳號    
Dim int_week_deduct_session : int_week_deduct_session = 0                           '每週期(期末)應扣除堂數
Dim str_client_website : str_client_website = "C"                                   'T=TutorABC, C=VIPABC
Dim int_product_point_type : int_product_point_type = 0                             'cfg_product.point_type
Dim int_account_cycle_day : int_account_cycle_day = 0                               '結算週期
Dim arr_contract_session_point_info : arr_contract_session_point_info = null
Dim arr_cfg_product_tmp : arr_cfg_product_tmp = null
Dim arr_session_cycle_data_info_tmp : arr_session_cycle_data_info_tmp = null
Dim int_chk_client_type, int_member_client_sn, int_client_attend_level
Dim bolFinishDCGS : bolFinishDCGS = getSession("FinishDCGS" , CONST_DECODE_NO) '是否完成DCGS

int_member_client_sn = getSession("client_sn", CONST_DECODE_NO)

intVIPABCJR = session("VIPABCJR")
if ( 1 = sCInt(intVIPABCJR) ) then
	bolVIPABCJR = true
end if

if ( isEmptyOrNull(bolVIPABCJR) ) then
	bolVIPABCJR = isVIPABCJRClient(int_member_client_sn, CONST_VIPABC_RW_CONN)
end if

if ( true = bolVIPABCJR ) then
    strVIPABCJR = "_JR"
end if

'/****************************************************************************************
'描述： 發信給客戶相關
'參數描述 / 建立人員 / 建立日期：
'*****************************************************************************************/
'新增/取消 訂位的 信件url
Dim str_class_email_url : str_class_email_url = "http://" & CONST_VIPABC_WEBHOST_NAME & "/program/mail_template/member/reservation_class/order_cancel_mail.asp"

'信件主旨
Dim str_class_email_subject : str_class_email_subject = getWord("CLASS_EMAIL_SUBJECT")  '"使用課程預約新增/取消系統"

'客戶訂課要寄信給下列的人
'Dim str_mail_to_management : str_mail_to_management = "kevinlai@tutorabc.com,sharonwang@tutorabc.com,kathywu@tutorabc.com,sabrinatsai@tutorabc.com,johnchiang@tutorabc.com,joankuo@tutorabc.com"
Dim str_mail_to_management : str_mail_to_management = CONST_MAIL_TO_VIPABC_MANAGERS_ACCOUNT & CONST_MAIL_TO_TUTOR_CS_MANAGERS_ACCOUNT & CONST_MAIL_TO_TUTOR_TS_MANAGERS_ACCOUNT

'/****************************************************************************************
'程式/元件簡述(目的) ： 所有訂課頁面 要用到的 include 檔 和載入頁面前的檢查
'備註：'被 /program/member/reservation_clas/normal_and_vip.asp include
       '被 /program/member/reservation_class/normal_and_vip_near_order.asp include
'歷程(作者/日期)：[Ray Chien] 2009/12/22 Created
'*****************************************************************************************/

'檢查客戶是否有登入
if (isEmptyOrNull(int_member_client_sn)) then
    Call alertGo(getWord("PLEASE_LOGIN"), "/program/member/member_login.asp", CONST_ALERT_THEN_REDIRECT)
    Response.End
end if

'檢查是否在退費中 若在退費中且沒有堂數的話 不能訂課
if (getIsClientMoneyBack(int_member_client_sn) = 1 and isCheckSpecialRule() and 401572 <> int_member_client_sn) then
    Call alertGo("", "/index.asp", CONST_NOT_ALERT_THEN_REDIRECT)
    Response.End
end if

'檢查是否有DCGS判斷
if ( isEmptyOrNull(bolFinishDCGS) ) then
   bolFinishDCGS = isFinishDCGS_session(int_member_client_sn, CONST_VIPABC_RW_CONN)
end if

'/****************************************************************************************
'描述： 客戶相關的資料
'參數描述 / 建立人員 / 建立日期：
'CONST_CLASS_TYPE_NORMAL / RayChien / 20091129
'*****************************************************************************************/
'TODO: DEBUG MODE 參數 若有設定的話 會輸出比較多的錯誤訊息給工程師看 正式機的話為空值
Dim str_debug_mode : str_debug_mode = ""

'客戶資料 COM 物件
Dim obj_member_opt : Set obj_member_opt = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
Dim int_member_result : int_member_result = obj_member_opt.prepareData(int_member_client_sn, CONST_VIPABC_RW_CONN)

if (int_member_result = CONST_FUNC_EXE_SUCCESS) then
    str_client_full_ename = obj_member_opt.getData("fname") & " " & obj_member_opt.getData("lname") '取得客戶英文名字
    str_client_login_email_addr = obj_member_opt.getData("email")                                   '取得客戶登入的帳號
    int_chk_client_type = obj_member_opt.getData("client_type")

	int_client_attend_level = session("int_client_attend_level_cache_data")
	if isEmptyOrNull(int_client_attend_level) then
		int_client_attend_level = obj_member_opt.getData(CONST_CUSTOMER_NOW_LEVEL)
		session("int_client_attend_level_cache_data") = int_client_attend_level
	end if
    'int_client_attend_level = obj_member_opt.getData(CONST_CUSTOMER_NOW_LEVEL)                  '取得客戶的等級

	str_service_start_date = session("str_service_start_date_cache_data")
	if isEmptyOrNull(str_service_start_date) then
		str_service_start_date = obj_member_opt.getData(CONST_CONTRACT_SERVICE_BEGIN_DATE)
		session("str_service_start_date_cache_data") = str_service_start_date
	end if
    'str_service_start_date = obj_member_opt.getData(CONST_CONTRACT_SERVICE_BEGIN_DATE)          '取得目前服務開始日

	str_service_end_date = session("str_service_end_date_cache_data")
	if isEmptyOrNull(str_service_end_date) then
		str_service_end_date = obj_member_opt.getData(CONST_CONTRACT_SERVICE_END_DATE)
		session("str_service_end_date_cache_data") = str_service_end_date
	end if
    'str_service_end_date = obj_member_opt.getData(CONST_CONTRACT_SERVICE_END_DATE)              '取得目前服務結束日

	str_now_use_contract_sn = session("str_now_use_contract_sn_cache_data")
	if isEmptyOrNull(str_now_use_contract_sn) then
		str_now_use_contract_sn = obj_member_opt.getData(CONST_CONTRACT_SN)
		session("str_now_use_contract_sn_cache_data") = str_now_use_contract_sn
	end if
    'str_now_use_contract_sn = obj_member_opt.getData(CONST_CONTRACT_SN)                         '目前電子合約序號

	bol_client_have_valid_contract = session("bol_client_have_valid_contract_cache_data")
	if isEmptyOrNull(bol_client_have_valid_contract) then
		bol_client_have_valid_contract = obj_member_opt.getData(CONST_CUSTOMER_IS_HAVE_CONTRACT)
		session("bol_client_have_valid_contract_cache_data") = bol_client_have_valid_contract
	end if
    'bol_client_have_valid_contract = obj_member_opt.getData(CONST_CUSTOMER_IS_HAVE_CONTRACT)    '取得客戶是否有合約

	str_now_use_product_sn = session("str_now_use_product_sn_cache_data")
	if isEmptyOrNull(str_now_use_product_sn) then
		str_now_use_product_sn = obj_member_opt.getData(CONST_CONTRACT_PRODUCT_SN)
		session("str_now_use_product_sn_cache_data") = str_now_use_product_sn
	end if
    'str_now_use_product_sn = obj_member_opt.getData(CONST_CONTRACT_PRODUCT_SN)                  '離目前時間最近的一筆 客戶合約的產品序號

	str_purchase_account_sn = session("str_purchase_account_sn_cache_data")
	if isEmptyOrNull(str_purchase_account_sn) then
		str_purchase_account_sn = obj_member_opt.getData(CONST_LAST_CONTRACT_ACCOUNT_SN)
		session("str_purchase_account_sn_cache_data") = str_purchase_account_sn
	end if
    'str_purchase_account_sn = obj_member_opt.getData(CONST_LAST_CONTRACT_ACCOUNT_SN)            '離目前時間最近的一筆的於PURCHASE的序號
    
    '設定 client website 和 信件來源 和 主旨
	Dim tmp_use_str_client_website
	tmp_use_str_client_website = session("tmp_use_str_client_website_cache_data")
	if(isEmptyOrNull(tmp_use_str_client_website)) then
		tmp_use_str_client_website = obj_member_opt.getData(CONST_CLIENT_WEB_SITE)
		session("tmp_use_str_client_website_cache_data") = tmp_use_str_client_website
	end if
	
    if (Not isEmptyOrNull(tmp_use_str_client_website)) then
        str_client_website = tmp_use_str_client_website
        'VIPABC
        'if (UCase(str_client_website) = "C") then
            'str_class_email_url = "http://" & CONST_VIPABC_WEBHOST_NAME & "/program/mail_template/member/reservation_class/order_cancel_mail.asp"
            'str_class_email_subject = getWord("CLASS_EMAIL_SUBJECT")
        'TutorABC
        'elseif (UCase(str_client_website) = "T") then
            'str_class_email_url = "http://" & CONST_ABC_WEBHOST_NAME & "/asp/mail/FirstClassNote/index.asp"
            'str_class_email_subject = "使用課程預約新增/取消系統"
        'end if
    end if

    if (Not isEmptyOrNull(str_now_use_product_sn)) then
        '設定 一對一和一對三和大會堂 幾小時內不能訂課 和 不能取消 要顯示的訊息
        Call setNoOrderAndCancelMessage(str_now_use_product_sn, arr_msg_no_order_in_hour, arr_msg_no_cancel_in_hour, arr_no_cancel_in_minute, CONST_VIPABC_RW_CONN)        
    end if

    '取得客戶總剩餘堂數 和 客戶當週剩餘的堂數
    if (Not isEmptyOrNull(str_purchase_account_sn)) then
		if ( true = bolVIPABCJR ) then
			arr_contract_session_point_info = getCustomerContractSessionPointInfoJR(int_member_client_sn, "", "", CONST_VIPABC_RW_CONN)
		else
			arr_contract_session_point_info = getCustomerContractSessionPointInfo("", "", str_purchase_account_sn, CONST_VIPABC_RW_CONN)
		end if
        if (IsArray(arr_contract_session_point_info)) then
            '客戶總剩餘堂數
            flt_client_total_session = CDbl(arr_contract_session_point_info(20))
            '客戶當週剩餘的堂數
            flt_client_week_session = CDbl(arr_contract_session_point_info(21))
			'客戶贈送堂數(紅利包)
			if ( Not isEmptyOrNull(arr_contract_session_point_info(30)) ) then
				flt_client_bonus_session = CDbl(arr_contract_session_point_info(30))
			else
				flt_client_bonus_session = 0
			end if
            '取得產品的資訊
            arr_cfg_product_tmp = getProductInfo(arr_contract_session_point_info(19), CONST_VIPABC_RW_CONN)
            if (IsArray(arr_cfg_product_tmp)) then
                int_product_point_type = arr_cfg_product_tmp(13, 0)
                '取得 客戶購買合約的產品類型 1:一般產品 2:定時定額
                if (Not isEmptyOrNull(arr_cfg_product_tmp(10, 0))) then
                    int_client_purchase_product_type = CInt(arr_cfg_product_tmp(10, 0))
                    '若是定時定額產品，取得客戶本週的週期初開始日 和 本週的週期末結束日 (結算日)
                    if (CONST_PRODUCT_QUOTA = int_client_purchase_product_type) then
                        '每週期(期末)應扣除堂數
                        int_week_deduct_session = CInt(arr_cfg_product_tmp(3, 0))
						'結算週期
                        int_account_cycle_day = CInt(arr_cfg_product_tmp(2, 0))
                        arr_session_cycle_data_info_tmp = getSessionCycleDateInfo(Array(int_account_cycle_day, str_service_start_date), now)
                        if (IsArray(arr_session_cycle_data_info_tmp)) then
                            dat_now_week_account_start_date = arr_session_cycle_data_info_tmp(2)
                            dat_now_week_account_end_date = arr_session_cycle_data_info_tmp(3)
                        end if
                    end if
                end if 
            end if 'end of if (IsArray(arr_cfg_product_tmp)) then
        end if
    end if

    if (Not isEmptyOrNull(str_service_end_date)) then
        str_service_end_date = getFormatDateTime(str_service_end_date, 5)
        '合約未到期
        if (DateDiff("d",DateValue(now), DateValue(str_service_end_date)) >= 0) then
            bol_client_contract_edate_valid = true
            bol_client_have_valid_contract = true   '有 有效合約

            '若可用堂數為0 可能是合約尚未開始 則用總堂數替代
            'if (flt_client_total_session <= 0) then
                'flt_client_total_session = getCustomerContractTotalSessionPoint(int_member_client_sn)
            'end if
        end if
    end if
else
    '錯誤訊息
    if (isEmptyOrNull(str_debug_mode)) then
        '"訂位失敗，請與客服人員連絡：02-3365-9999"
        Call alertGo(str_msg_reservation_class_fail_contact_cs, "/index.asp", CONST_ALERT_THEN_REDIRECT)
    else
        Call alertGo(str_msg_reservation_class_fail_contact_cs & "(無法正確取得客戶的資料 錯誤訊息:" & obj_member_opt.PROCESS_STATE_RESULT & ")", "/index.asp", CONST_ALERT_THEN_REDIRECT)
    end if
    Response.End
end if
Set obj_member_opt = Nothing

'檢查客戶的電子合約是否未確認(如果有電子合約存在的話)
if (isTemporalContractNoConfirm(str_now_use_contract_sn)) then
    '導回電子合約確認畫面
    Call alertGo("", "/program/member/contract_audit.asp", CONST_NOT_ALERT_THEN_REDIRECT)
    Response.End
end if

if (   (isPayMember() and  "1" <> session("boleverclass") ) and ( Request.ServerVariables("PATH_INFO") <> "/program/member/class.asp" )  ) then '開通客戶
	Dim newMemberTestPass : newMemberTestPass = 1 '新手上路是否測試全過？
	'DCGS及語言程度分析及 首測重測第十三題要過 ，這三個條件都必需符合才能使用訂課功能。
	'客戶還未做過 首測/重測  2012-08-22 Johnny
	'2012-10-26 修改規則
	'* 後台: 只要有程度分析 + 預約測試就可以訂課
	'* 前台:  測試失敗(EnvironmentTestRecord.finishTest=E)需要預約時間(含2情境：測試通過、未測試，不分前後台預約)+DCGS+程度分析 --> 可訂課
	'* 前台:  測試成功(EnvironmentTestRecord.finishTest=Y)+DCGS+程度分析 --> 可訂課
	'session("bolFirstTestQuestionaryhaveTestOk") 電腦測試是否成功 有 Y E N幾種狀態	by Johnny
	'session("bolFirstTestQuestionaryhaveTestNOTOkOrder") 是否預約 有Y N	by Johnny
	
	
	'''重抓session
	if "1" <> getSession("bolkompletDcgs", CONST_DECODE_NO) or "1" <> getSession("bolComplateLevelTest", CONST_DECODE_NO)  or  "N" = getSession("bolFirstTestQuestionaryhaveTestNOTOkOrder", CONST_DECODE_NO)  or  "Y"<>getSession("bolFirstTestQuestionaryhaveTestOk", CONST_DECODE_NO) then
		Set objPCInfoInsert3 = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")  
		Dim  client_sn_tmp2 : client_sn_tmp2 = getSession("client_sn", CONST_DECODE_NO)
		
		if "1" <> getSession("bolkompletDcgs", CONST_DECODE_NO) or "1" <> getSession("bolComplateLevelTest", CONST_DECODE_NO) then
			strSql = " select  industry AS '1industry', role AS '2role',interest AS '3interest', Focus AS '4Focus' , olevel , nlevel   from client_basic WITH (NOLOCK) where  sn=@intclient_sn "		
			
			intSqlWriteResult = objPCInfoInsert3.excuteSqlStatementEach(strSql, Array( client_sn_tmp2 ), CONST_TUTORABC_RW_CONN)     '有傳入參數用
			if not objPCInfoInsert3.EOF then
				str1industry = objPCInfoInsert3("1industry")
				str2role = objPCInfoInsert3("2role")
				str3interest = objPCInfoInsert3("3interest")
				str4Focus = objPCInfoInsert3("4Focus")
				olevel_ = objPCInfoInsert3("olevel")
				nlevel_ = objPCInfoInsert3("nlevel")
			end if
			
			if ( true = bolVIPABCJR ) then
				bolkompletDcgs = isDCGSClient(int_member_client_sn, 0, CONST_VIPABC_RW_CONN)
				if ( true = bolkompletDcgs ) then
					session("bolkompletDcgs") = "1" '全部設定完畢過
				else
					session("bolkompletDcgs") = "2" 'DCGS資料有缺
				end if
			else
				if ( (not isEmptyOrNull(str1industry)) AND (not isEmptyOrNull(str2role)) AND (not isEmptyOrNull(str3interest)) AND (not isEmptyOrNull(str4Focus)) ) then
					session("bolkompletDcgs") = "1" '全部設定完畢過
				else
					session("bolkompletDcgs") = "2" 'DCGS資料有缺
				end if
			end if
			
			if ( isEmptyOrNull(olevel_) and isEmptyOrNull(nlevel_) ) then
				session("bolComplateLevelTest") = "0"
			else
				session("bolComplateLevelTest") = "1"
			end if
		end if
		
		if "N" = getSession("bolFirstTestQuestionaryhaveTestNOTOkOrder", CONST_DECODE_NO)  then
			my_sql = " select * from  add_client_datetime ac WITH (nolock) LEFT JOIN dbo.questionary_fill_record qf WITH (NOLOCK) ON ac.sn = qf.add_client_datetime_sn   where ac.client_sn = @intclient_sn and ac.status not in (5) and ac.description <> N'客戶於VIP官網新手上路自行填寫問卷' AND (qf.qfr_passed<>'2' OR qf.qfr_passed IS NULL) "
			call objPCInfoInsert3.excuteSqlStatementEach(my_sql, Array( client_sn_tmp2 ) , CONST_VIPABC_RW_CONN)   
			if objPCInfoInsert3.RecordCount > 0 then
				session("bolFirstTestQuestionaryhaveTestNOTOkOrder") = "Y"
			end if		
		end if
		
		'修正VIP 某些客不知為何 總是會跳新手上路alert訊息 ， 看加這段是否可以解決
		if "Y"<>getSession("bolFirstTestQuestionaryhaveTestOk", CONST_DECODE_NO) then
			my_sql = "/* EnvironmentTestService.checkEnvironmentHaveTestOk */ SELECT TOP 1 finishTest FROM EnvironmentTestRecord WITH(NOLOCK)  WHERE clientSN = @client_sn  order by inpdate desc "
			Dim finishTestStr : finishTestStr="N"
			intResult = objPCInfoInsert3.excuteSqlStatementEach( my_sql , Array( client_sn_tmp2 ) , CONST_TUTORABC_RW_CONN)
			if not objPCInfoInsert3.EOF then
				finishTestStr = objPCInfoInsert3("finishTest")
			end if
			if not "Y"=finishTestStr then
				my_sql = "/* EnvironmentTestService.checkEnvironmentHaveTestOk */ SELECT qfr_passed FROM questionary_fill_record WITH(NOLOCK)  WHERE client_sn= @client_sn and  qfr_passed = 1 "
				intResult = objPCInfoInsert3.excuteSqlStatementEach( my_sql , Array( client_sn_tmp2 ) , CONST_TUTORABC_RW_CONN)
				if not objPCInfoInsert3.EOF then
					finishTestStr = "Y"
				end if
			end if
			session("bolFirstTestQuestionaryhaveTestOk") = finishTestStr
		end if
		objPCInfoInsert3.close()
		set objPCInfoInsert3 =nothing
	end if
	'''重抓session END

	Dim isCanClass : isCanClass = "N"
	if (  "E" =  getSession("bolFirstTestQuestionaryhaveTestOk", CONST_DECODE_NO) and ("1" = getSession("bolkompletDcgs", CONST_DECODE_NO) ) and ("1" = getSession("bolComplateLevelTest", CONST_DECODE_NO) ) and ("Y" = getSession("bolFirstTestQuestionaryhaveTestNOTOkOrder", CONST_DECODE_NO) ) ) then
		isCanClass = "Y"
	end if
	if  ( "Y" =  getSession("bolFirstTestQuestionaryhaveTestOk", CONST_DECODE_NO) and ("1" = getSession("bolkompletDcgs", CONST_DECODE_NO) ) and ("1" = getSession("bolComplateLevelTest", CONST_DECODE_NO)) ) then
		isCanClass = "Y"
	end if
	
	if ( "N" = isCanClass and isCheckSpecialRule() ) then
		if  "N" = getSession("bolFirstTestQuestionaryhaveTestOk", CONST_DECODE_NO) and isCheckSpecialRule() then
			'newMemberTestPass = 0
			Call alertGo("亲爱的客户您好\n欢迎您加入VIPABC学习行列，为确保您的上课学习质量，上课前请您先至「新手上路」专区完成下列步骤：\nStep 1.测试上课电脑设备\nStep 2.DCGS学习偏好设定\nStep 3.语言程度分析\nStep 4.观看教学影片","/program/member/environment_test/guide_index.asp",CONST_ALERT_THEN_REDIRECT)  '2012-08-28 改成要彈
			Response.End
		end if

		'判斷客戶是否還未做DCGS
		dim bolAlertGo : bolAlertGo = true
		if ( "1" <> session("bolkompletDcgs") and isCheckSpecialRule() ) then
			newMemberTestPass = 0
			if ( true = bolVIPABCJR ) then
				dim strRequestHtml : strRequestHtml = Request.ServerVariables("script_name")
				strDCGSUrl = "/program/aspx/AspxLogin.asp?url_id=5"
				if ( sCstr(strRequestHtml) = sCstr(strDCGSUrl) ) then
					bolNotAlertGo = false
				end if
			else
				strDCGSUrl = "/program/member/learning_dcgs.asp"
			end if
			if ( true = bolNotAlertGo ) then
				Call alertGo("亲爱的会员您好！！由於您尚未完成[DCGS学习偏好设置]，VIPABC暂时无法为您安排符合您需求的教材及顾问，请完成[DCGS学习偏好设置]後再进行预约。", strDCGSUrl, CONST_ALERT_THEN_REDIRECT) '亲爱的会员您好！！由於您尚未完成[DCGS学习偏好设置]，VIPABC暂时无法为您安排符合您需求的教材及顾问，请完成[DCGS学习偏好设置]後再进行预约。
				'Call alertGo("亲爱的客户您好\n欢迎您加入VIPABC学习行列，为确保您的上课学习质量，上课前请您先至「新手上路」专区完成下列步骤：\nStep 1.测试上课电脑设备\nStep 2.DCGS学习偏好设定\nStep 3.语言程度分析\nStep 4.观看教学影片","/program/member/environment_test/guide_index.asp",CONST_ALERT_THEN_REDIRECT)
				Response.End
			end if
		end if

		'客戶還未做程度分析
		if ((isEmptyOrNull( session("bolComplateLevelTest") ) or CStr( session("bolComplateLevelTest") ) = "0") and ( true = bolVIPABCJR OR CStr(int_chk_client_type) <> "4" )  and isCheckSpecialRule()) then
			newMemberTestPass = 0
			Call alertGo(getWord("NO_COMPLETE_LEVEL_ANALYZE_1") & "\n\n" & getWord("NO_COMPLETE_LEVEL_ANALYZE_2") & "\n\n" & getWord("WEB_NAME") & getWord("NO_COMPLETE_LEVEL_ANALYZE_3") & "\n\n" & getWord("NO_COMPLETE_LEVEL_ANALYZE_4"), "/program/member/learning_level.asp", CONST_ALERT_THEN_REDIRECT) '亲爱的会员您好!!\n\n由於您尚未完成程度分析\n\nVIPABC暂时无法为您安排符合您需求的教材与顾问\n\n请完成程度分析後再进行预约!
			'Call alertGo("亲爱的客户您好\n欢迎您加入VIPABC学习行列，为确保您的上课学习质量，上课前请您先至「新手上路」专区完成下列步骤：\nStep 1.测试上课电脑设备\nStep 2.DCGS学习偏好设定\nStep 3.语言程度分析\nStep 4.观看教学影片","/program/member/environment_test/guide_index.asp",CONST_ALERT_THEN_REDIRECT)
			Response.End
		end if 

		'Call alertGo("亲爱的客户您好\n欢迎您加入VIPABC学习行列，为确保您的上课学习质量，上课前请您先至「新手上路」专区完成下列步骤：\nStep 1.测试上课电脑设备\nStep 2.DCGS学习偏好设定\nStep 3.语言程度分析\nStep 4.观看教学影片","/program/member/environment_test/guide_index.asp",CONST_ALERT_THEN_REDIRECT)  '2012-08-28 改成要彈
		'Response.End
	end if
	
	
	'判斷客戶是否還未做 First Session
	'str_last_first_session_datetime = getLastReservationFirstSessionDatetime(int_member_client_sn)
	'if (str_last_first_session_datetime = "-2" and isCheckSpecialRule()) then
	'	Call alertGo(getWord("DO_FIRST_SESSION_FIRST"), "/program/member/first_session.asp", CONST_ALERT_THEN_REDIRECT) '亲爱的客户，您好。第一次上课前，请先预约「First Session」测试，以确保您的上课品质。
	'	Response.End
	'end if
end if
%>
<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->
<!--#include file="include_all_message.inc"-->