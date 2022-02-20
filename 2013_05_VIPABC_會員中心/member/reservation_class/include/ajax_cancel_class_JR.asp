<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<!--#include virtual="/lib/functions/functions_ABCSystemSendMail/functions_ABCSystemSendMail.asp"--><%'ABC系統信%>
<!--#include virtual="/lib/functions/functions_tool.asp"--><%'asp呼叫javascript%>
<%    
Dim int_client_sn, str_client_ip '客戶編號、客戶的IP
Dim flt_refund_session : flt_refund_session = 0.0 '要歸還的堂數
Dim str_attend_list_sn : str_attend_list_sn = Trim(Request("attend_list_sn")) 'client_attend_list.sn
Dim str_attend_list_date : str_attend_list_date = Trim(Request("attend_list_date")) 'client_attend_list.attend_date
Dim str_attend_list_time : str_attend_list_time = Trim(Request("attend_list_time")) 'client_attend_list.attend_sestime 
Dim str_cancel_success_msg_for_email : str_cancel_success_msg_for_email = "" '訂課成功後 寄送給客戶的新增課程訊息
Dim int_class_type : int_class_type = null '課程類型(一對一或一對多或大會堂)
Dim arr_success_reservation_class : arr_success_reservation_class = null '成功預訂課程的日期和扣堂 (0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn
Dim str_sql, arr_param_value, arr_attend_list_data, int_update_result
Dim dat_attend_date, int_attend_sestime, str_class_type_description, str_cancel_success_msg, str_cancel_fault_msg, str_error_code
Dim arr_client_attend_list, arr_session_cycle_info, arr_tmp
dim str_class_datetime, arr_cancel 
dim intProfileId, intContractId
dim intPage : intPage = Trim(Request("intPage"))
dim intPageType : intPageType = Trim(Request("intPageType"))

'本週已使用堂數(排除return)
Dim int_week_use_session_count : int_week_use_session_count = 0

str_cancel_success_msg = ""
str_cancel_fault_msg = ""
str_error_code = ""

'取得客戶編號
if ( g_var_client_sn <> "" ) then
    int_client_sn = g_var_client_sn
end if

'檢查是否有傳入要取消課程的資訊
if (isEmptyOrNull(str_attend_list_sn) Or isEmptyOrNull(str_attend_list_date) Or isEmptyOrNull(str_attend_list_time)) then
    Response.Clear
    Response.Write "-10"
    Response.End
end if

'檢查要取消的課程是否在四小時前
'取得相差的分
Dim str_cancel_class_datetime : str_cancel_class_datetime = str_attend_list_date & " " & Right("0"&str_attend_list_time, 2) & ":30"
if ( DateDiff("n", now(), str_cancel_class_datetime) <= 240 ) then
    Response.Clear
    Response.Write "-2"
    Response.End
end if

'取得客戶IP位置
str_client_ip = sCStr(getIpAddress())

'先確定客戶真的有訂這堂課 且 訂課日期不小於今日 在取消
'str_sql = " SELECT sn, svalue, CONVERT(VARCHAR, attend_date, 111) AS sdate, attend_sestime, attend_livesession_types, attend_level FROM client_attend_list WHERE " &_
'                " (client_sn = @client_sn) AND (sn = @sn) AND (valid = 1) AND (CONVERT(VARCHAR, attend_date, 111) = CONVERT(VARCHAR, @str_attend_list_date, 111)) AND (attend_sestime = @stime) "
str_sql = " SELECT ClassRecordId, ClassCostPoints, CONVERT(VARCHAR, ClassStartDate, 111) AS sdate, ClassStartTime, ClassTypeId, ProfileLearnLevel, LobbySessionId FROM member.ClassRecord "
str_sql = str_sql & " WHERE (client_sn = @client_sn) AND (ClassRecordId = @ClassRecordId) AND (ClassStatusId = 2) AND (CONVERT(VARCHAR, ClassStartDate, 111) = CONVERT(VARCHAR, @ClassStartDate, 111)) AND (ClassStartTime = @ClassStartTime) "
arr_param_value = Array(int_client_sn, str_attend_list_sn, str_attend_list_date, str_attend_list_time)
arr_attend_list_data = excuteSqlStatementRead(str_sql, arr_param_value, CONST_VIPABC_RW_CONN)
'response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")

if ( isSelectQuerySuccess(arr_attend_list_data) ) then
    if ( Ubound(arr_attend_list_data) >= 0 ) then

        '取得課程在結算週期內的資訊，只有大陸定時定額產品才需要判斷
        'arr_session_cycle_info = getSessionCycleDateInfo(Array(int_account_cycle_day, str_service_start_date), str_attend_list_date)
        'if ( IsArray(arr_session_cycle_info) ) then
        '    '預訂或取消當週課程
        '    if ( arr_session_cycle_info(0) ) then
        '        '取出這個週期內的上課筆數(排除return)
        '        arr_client_attend_list = getCustomerBookingCount(int_client_sn, arr_session_cycle_info(2), arr_session_cycle_info(3), CONST_VIPABC_RW_CONN)
        '        int_week_use_session_count = sCInt(arr_client_attend_list(4))
        '    end if
        'end if

        '重新定義 成功預訂課程的日期和扣堂 陣列
        Redim arr_success_reservation_class(2, 0)

        '要歸還的堂數
        flt_refund_session = sCDbl(arr_attend_list_data(1, 0))
        flt_refund_session = sCDbl(flt_refund_session) / CONST_SESSION_TRANSFORM_POINT
        int_client_attend_level = arr_attend_list_data(5, 0)
        dat_attend_date = arr_attend_list_data(2, 0) 
        int_attend_sestime = arr_attend_list_data(3, 0) 
        str_cancel_class_type = arr_attend_list_data(4, 0)
        dim arrClassTypeVIPJR : arrClassTypeVIPJR = ""
        dim arrClassTypeVIPABC : arrClassTypeVIPABC = ""

        arrClassTypeVIPJR = split(CONST_CLASS_TYPE_VIPJR,",") '1, 1, 4, 5, 6, 2, 3
	    arrClassTypeVIPABC = split(CONST_CLASS_TYPE_VIPABC,",") '1, 7, 2, 3, 14, 4, 99
        '轉成新Table的值
	    for intI = 0 to Ubound(arrClassTypeVIPJR) 
	        if ( sCInt(str_cancel_class_type) = sCInt(arrClassTypeVIPJR(intI)) ) then
                str_cancel_class_type = arrClassTypeVIPABC(intI)
                exit for
            end if
	    next

        dim intLobbySessionSn : intLobbySessionSn = arr_attend_list_data(6, 0)
        'str_sql = " UPDATE client_attend_list SET attend_consultant = null, attend_room = null, " &_
        '                " attend_mtl_1 = null, attend_mtl_2 = null, session_sn = null, " &_
        '                " attend_level = null, valid = 0, view_ip = @view_ip, " &_
        '                " special_sn = null, attend_datetime = @attend_datetime " &_
        '                " WHERE (client_sn = @client_sn) AND (sn = @sn) AND (valid = 1) "
        'arr_param_value = Array(str_client_ip, getFormatDateTime(now(), 2), int_client_sn, str_attend_list_sn)
        '
        'int_update_result = excuteSqlStatementWrite(str_sql, arr_param_value, CONST_VIPABC_RW_CONN)
        '課程日期/時間

        '**** 參數設定 START ****
        Dim int_db_connect : int_db_connect = CONST_VIPABC_RW_CONN                          'DB Connection Code
        Dim int_website_code : int_website_code = CONST_VIPABC_WEBSITE                      'Website Code
        Dim str_class_minute : str_class_minute = "30"
        '**** 參數設定 END   ****
        'function cancelNewClassRecord(ByVal p_intClassRecordId, ByVal p_intClassType, ByVal p_strClassCostPoints, ByVal p_strClassStartDate, ByVal p_intClassStartTime, _
        '                             ByVal p_intClient_sn, ByVal p_intContract_sn, ByVal p_intTestMode)
        '/****************************************************************************************
        '描述      ：取消課程
        '牽涉數據表：
        '備註      ：
        '歷程(作者/日期/原因) ：[ArTHurWu] 2013/04/06 Created
        '*****************************************************************************************/
        'todo 上線要改為0
        'arr_cancel = cancelNewClassRecord(str_attend_list_sn, str_cancel_class_type, flt_refund_session, str_attend_list_date, str_attend_list_time, int_client_sn, str_now_use_contract_sn, 0)
        arr_cancel = cancelNewClassRecord(str_attend_list_sn, str_cancel_class_type, flt_refund_session, str_attend_list_date, str_attend_list_time, int_client_sn, str_now_use_contract_sn, 1)
                        
        'if ( int_update_result > 0 ) then
            '紀錄已成功訂課的資訊
            '(0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn
            arr_success_reservation_class(0, 0) = str_attend_list_date
            arr_success_reservation_class(1, 0) = flt_refund_session
            arr_success_reservation_class(2, 0) = str_attend_list_sn

            '更新要顯示的取消成功訊息
            'if (arr_attend_list_data(5, 0) > 12) then
            if ( 0 < sCInt(intLobbySessionSn) ) then
                str_class_type_description = getWord("SPECIAL_SESSION") '"大會堂"
                int_class_type = CONST_CLASS_TYPE_LOBBY
            elseif (Int(arr_attend_list_data(4, 0)) = CONST_CLASS_TYPE_ONE_ON_THREE) then
                'str_class_type_description = getWord("NORMAL_CLASS_NOTE_1") '"小班(1-3人)"
                str_class_type_description = "小班制" '"小班(1-3人)"
                int_class_type = CONST_CLASS_TYPE_ONE_ON_THREE
            elseif (Int(arr_attend_list_data(4, 0)) = CONST_CLASS_TYPE_ONE_ON_FOUR) then
                str_class_type_description = "小班制"
                int_class_type = CONST_CLASS_TYPE_ONE_ON_FOUR
            elseif (Int(arr_attend_list_data(4, 0)) = CONST_CLASS_TYPE_ONE_ON_ONE) then
                str_class_type_description = getWord("VIP_CLASS_NOTE_1") '"1对1"
                int_class_type = CONST_CLASS_TYPE_ONE_ON_ONE
            end if

            '更新取消課程成功的訊息
            str_cancel_success_msg = dat_attend_date & " " & int_attend_sestime & ":30 " & str_class_type_description

            '更新要寄送email給客戶的取消課程成功 課程訊息
            str_cancel_success_msg_for_email = str_cancel_success_msg_for_email & getClassMsgForEmail(int_class_type, dat_attend_date, int_attend_sestime, CONST_CANCEL_CLASS) & "$br$"

            'if ( DateDiff("n",now(), dat_attend_date & " " & int_attend_sestime & ":30:00") <= 1440 ) then
            '    '更新各時段，各身份可上課狀態
            '    Call updateBookingStatus(dat_attend_date, int_attend_sestime, CONST_VIPABC_RW_CONN)
            'end if
				
            'Special 1 on 1 取消課程
            'Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
            'sql = " SELECT sn FROM Special1on1Session WHERE Valid = 1 AND ClientPlatform = 2 AND ClientBookingSn = "&str_attend_list_sn
            'intQueryResult = rs.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_RW_CONN)
            'if ( not rs.eof ) then '是Special 1 on 1課程
            '    Set objInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
            '    sql = " SELECT  a.sn , "
            '    sql = sql & "        CONVERT(VARCHAR(10),SessionDate,111)+' '+RIGHT('0'+CONVERT(VARCHAR(2),SessionTime),2)+':30' SessionDateTime, "
            '    sql = sql & "        b.userid+'@tutorabc.com' submitterEmail, "
            '    sql = sql & "        c.basic_fname+' '+c.basic_lname consultantName, "
            '    sql = sql & "        a.ClientSn, "
            '    sql = sql & "        d.lead_sn, "
            '    sql = sql & "        d.cname "
            '    sql = sql & " FROM    Special1on1Session a WITH ( NOLOCK ) "
            '    sql = sql & " INNER JOIN staff_basic b WITH (NOLOCK) ON a.Creator = b.sn "
            '    sql = sql & " LEFT JOIN con_basic c WITH (NOLOCK) ON a.Consultant = c.con_sn "
            '    sql = sql & " INNER JOIN client_basic d WITH (NOLOCK) ON a.ClientSn = d.sn "
            '    sql = sql & " WHERE ClientBookingSn = "&str_attend_list_sn&" "
            '    sql = sql & "        AND a.Valid = 1 " '有效, 未被顧問取消或客戶取消
            '    sql = sql & "        AND ClientPlatform = 2 " 'VIPABC
            '    sql = sql & "        AND a.Status NOT IN ( 7, 8 ) "
            '    intQueryResult = objInfo.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_RW_CONN)
            '
            '    if ( not objInfo.eof ) then
            '        '更新狀態
            '        Set rs1 = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
            '        sql = " UPDATE Special1on1Session SET Status=7, isProcess=0, Valid=0 WHERE sn=@sn "
            '        arrParam=Array(objInfo("sn"))
            '        intQueryResult = rs1.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_RW_CONN)
            '        rs1.close
            '        Set rs1 = nothing
			'			
            '        Dim SendMailabc :set SendMailabc = new ABCSystemSendMail
            '        strEmailTo = objInfo("submitterEmail")&",CSOP@tutorabc.com"
            '        strSessionDateTime = objInfo("SessionDateTime")
            '        strConsultantName = objInfo("consultantName")
            '        SendMailabc.str_sendlettermail_sn = "special1on1_0001"
            '        SendMailabc.strEmailTitle = "Special 1 on 1 客戶 "&objInfo("cname")&" 取消訂位"
            '        SendMailabc.client_cname = objInfo("cname") '客戶姓名
            '        SendMailabc.str_para2 = "http://"&CONST_VIPABC_IMSHOST_NAME&"/asp/cs/client_3.asp?lead="&objInfo("lead_sn")&"&client="&objInfo("ClientSn")&"&search=yes"
            '        SendMailabc.str_para4 = strSessionDateTime
            '        SendMailabc.str_para5 = strConsultantName
            '        if ( CFG_SPECIAL1ON1_DEBUG_MODEL = 1 ) then			
            '            SendMailabc.strEmailToAddr = CFG_SPECIAL1ON1_MAIL_TO '客戶的電子郵件
            '            testMailTo = strEmailTo
            '            SendMailabc.str_para1 = "<div style='background:#99FF88'>{內為正式站時會發的人mail 正式機時不會看到此訊息: To: "&testMailTo &" / CC:}</div>"
            '            result =  SendMailabc.MainSendMail()
            '        else
            '            SendMailabc.strEmailToAddr 	=  strEmailTo
            '            result = SendMailabc.MainSendMail()
            '        end if
            '        Set SendMailabc = nothing
            '    end if 'if ( not objInfo.eof ) then
            '    objInfo.close
            '    Set objInfo = nothing
            'end if 'if ( not rs.eof ) then
            'rs.close
            'Set rs = nothing
        'else 'if ( int_update_result > 0 ) then
        if ( isArray(arr_cancel) ) then
            'Response.Clear
        else
            'Response.Write("錯誤訊息" & g_str_run_time_err_msg & "$br$")
            '更新取消課程失敗的訊息
            str_error_code = "-11"
            'str_cancel_fault_msg = "<span class=""bold"">" & str_attend_list_date & " " & str_attend_list_time & ":30 " & "</span> (" & getWord("CANCEL_FAIL") & ")"    '取消失敗
        end if
    else 'if (Ubound(arr_attend_list_data) >= 0) then
        '更新取消課程失敗的訊息
        str_error_code = "-12"
        'str_cancel_fault_msg = getWord("YOU_HAVE_NO_ORDER") & " (" & getWord("CANCEL_FAIL") &")" '此堂課您並沒有預定
    end if
else 'if (isSelectQuerySuccess(arr_attend_list_data)) then
    '更新取消課程失敗的訊息
    str_error_code = "-13"
    'str_cancel_fault_msg = getWord("CANCEL_FAIL")   '"取消失敗"
end if

'Response.Clear

'失敗的訊息
if ( Not isEmptyOrNull(Trim(str_error_code)) ) then
    Response.Clear
    Response.Write Trim(str_error_code)

'若有取消課程成功的話 就發信給客戶
elseif (Not isEmptyOrNull(Trim(str_cancel_success_msg))) then
    dim strAlert : strAlert = ""
    Call sendOrderCancelEmail(str_class_email_url, str_client_login_email_addr, _
                                str_client_full_ename, str_cancel_success_msg_for_email, _
                                str_client_website)

    '發送信件給設定的主管和經理
    str_class_email_subject = "客戶 " & str_client_full_ename & " " & str_class_email_subject
    Call sendOrderCancelEmail(str_class_email_url, str_mail_to_management, _
                                                str_client_full_ename, str_cancel_success_msg_for_email, _
                                                str_client_website)
									  
    '在取一次客戶的總剩餘堂數 目的是為了處理同時有多筆cancel發生
    '客戶資料 COM 物件
    'Set obj_member_opt_exe = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
    'int_member_result_exe = obj_member_opt_exe.prepareData(g_var_client_sn, CONST_VIPABC_RW_CONN)
    'if (int_member_result_exe = CONST_FUNC_EXE_SUCCESS) then
        'flt_client_total_session_last_get = obj_member_opt_exe.getData(CONST_CUSTOMER_REMAIN_CONTRACT_POINT)             '取得客戶總剩餘堂數        
    'end if
    'Set obj_member_opt_exe = Nothing

    '剩餘堂數和取消前不一樣 代表課程已經從table刪除了
    'if (flt_client_total_session_last_get <> flt_client_total_session) then
        'flt_client_total_session = flt_client_total_session_last_get
    '剩餘堂數和取消前一樣 代表課程還未從table刪除(update 可能沒那麼立即)
    'else
        flt_client_total_session = Round(flt_client_total_session + flt_refund_session, 2)
    'end if

    '定時定額產品跑以下規則
    '定時定額產品跑以下規則
    'if ( CONST_PRODUCT_QUOTA = int_client_purchase_product_type ) then
    '    '檢查訂課時間，當週和非當週的總扣堂數，更新到客戶剩餘點數總表
    '        arr_tmp = Array(int_client_sn, str_purchase_account_sn, str_now_use_product_sn, _
    '                                    int_account_cycle_day, int_week_deduct_session, "cancel", _
    '                                    int_week_use_session_count, CONST_VIPABC_WEBSITE, _
    '                                    Request.ServerVariables("PATH_INFO"), 0, str_service_start_date, str_service_end_date)
    '              
    '    '檢查訂課時間，當週和非當週的總扣堂數，更新到客戶剩餘點數總表
    '    Call updatePointAfterReservationOrCancel(arr_success_reservation_class, arr_tmp, CONST_VIPABC_RW_CONN)
    'end if

    Response.Clear
    Response.Write getWord("CANCEL_SUCCESS") & chr(10) '您已成功取消
    Response.Write str_cancel_success_msg & getWord("COURSE") & "1" & getWord("CLASS") & chr(10)   '課程1堂
    
    'strAlert = getWord("CANCEL_SUCCESS") & chr(10) '您已成功取消
    'strAlert = strAlert & str_cancel_success_msg & getWord("COURSE") & "1" & getWord("CLASS") & chr(10)   '課程1堂

    '20120316 阿捨新增 無限卡產品判斷	
    Dim bolUnlimited : bolUnlimited = false
    if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
        bolUnlimited = true 
    end if

    if ( true = bolUnlimited ) then
    else
        Response.Write getWord("TOTAL_REFUND") & flt_refund_session & getWord("SESSION_POINT_COUNT") & chr(10)   '共计归还 xx 课时数
        Response.Write getWord("SESSION_POINT_FOR_USE") & flt_client_total_session  '可用课时数
        'strAlert = strAlert & getWord("TOTAL_REFUND") & flt_refund_session & getWord("SESSION_POINT_COUNT") & chr(10)   '共计归还 xx 课时数
        'strAlert = strAlert & getWord("SESSION_POINT_FOR_USE") & flt_client_total_session  '可用课时数
    end if
    'Call useASPtoJavascript("alert('"&strAlert&"'); ChangePage('"&intPage&"', '"&intPageType&"');")

    '例外的錯誤
else 'if (Not isEmptyOrNull(Trim(str_error_code))) then
    Response.Clear
    Response.Write "-1"
end if
%>