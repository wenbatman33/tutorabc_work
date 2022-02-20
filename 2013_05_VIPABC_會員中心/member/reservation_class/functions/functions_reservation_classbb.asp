<%
function getClassTypeDescription(ByVal p_int_class_type)
'/****************************************************************************************
'描述      ：取得課程類型的中文描述
'傳入參數  ：[p_int_class_type:Integer] 課程類型 一對一 or 一對多 or 大會堂
'回傳值    ：課程類型的中文描述 "小班(1-6人)课程" or "1对1课程" or "大会堂"
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2009/12/23 Created
'*****************************************************************************************/
    if (p_int_class_type = CONST_CLASS_TYPE_ONE_ON_THREE) then
        'getClassTypeDescription = getWord("NORMAL_CLASS_NOTE_2") '"小班(1-3人)课程"
        getClassTypeDescription = "小班制课程"
    elseif (p_int_class_type = CONST_CLASS_TYPE_ONE_ON_FOUR) then
        getClassTypeDescription = "小班制课程" '"1对4课程"
    elseif (p_int_class_type = CONST_CLASS_TYPE_ONE_ON_ONE) then
        getClassTypeDescription = getWord("VIP_CLASS_NOTE_2") '"1对1课程"
    elseif (p_int_class_type = CONST_CLASS_TYPE_LOBBY) then
        getClassTypeDescription = getWord("SPECIAL_SESSION") '"大会堂"
    else
        getClassTypeDescription = ""
    end if
end function


function getClassMsgForEmail(ByVal p_int_class_type, ByVal p_str_class_date, _
                             ByVal p_str_class_time, ByVal p_int_opt)
'/****************************************************************************************
'描述      ：取得新增訂位或取消訂位的課程訊息 (寄信時會用到)
'傳入參數  ：[p_int_class_type:Integer] 課程類型 一對一 or 一對多 or 大會堂
'            [p_str_class_date:String] 課程開始日期
'            [p_str_class_time:String] 課程開始時間
'            [p_int_opt:String] 
'               CONST_RESERVATION_CLASS = 101       '預定課程
'               CONST_RE_RESERVATION_CLASS = 102    '重新預定課程
'               CONST_CANCEL_CLASS = 103            '取消課程
'回傳值    ：新增訂位或取消訂位的課程訊息 (寄信時會用到)
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2009/12/24 Created
'*****************************************************************************************/
    Dim str_class_type, str_class_datetime

    '2009/12/19 21:30 - 22:15 取得這樣格式的字串
    str_class_datetime = p_str_class_date & " " & Right("0" & p_str_class_time, 2) & ":30" & " - " & Right("0" & (Int(p_str_class_time) + 1), 2) & ":15"

    '新增訂位
    if (p_int_opt = CONST_RESERVATION_CLASS) then
        '( 大会堂 -- 新增訂位 ) 取得這樣格式的字串
        str_class_type = "( " & getClassTypeDescription(p_int_class_type) & " -- " & getWord("NEW_ORDER") & " )"    '新增订位
    '重新訂位
    elseif (p_int_opt = CONST_RE_RESERVATION_CLASS) then
        '( 大会堂 -- 重新訂位 ) 取得這樣格式的字串
        str_class_type = "( " & getClassTypeDescription(p_int_class_type) & " -- " & getWord("NEW_REORDER") & " )"    '重新订位
    '取消訂位
    elseif (p_int_opt = CONST_CANCEL_CLASS) then
        '( 取消該堂預約 ) 取得這樣格式的字串
        str_class_type = "( " & getWord("CANCEL_ORDER") & " )" '取消該堂預約
    end if

    getClassMsgForEmail = str_class_datetime & str_class_type
end function


sub sendOrderCancelEmail(ByVal p_str_class_email_url, ByVal p_str_email_to_addr, _
                         ByVal p_str_email_to_name, ByVal p_str_class_msg, ByVal p_str_client_website)
'/****************************************************************************************
'描述      ：發送 新增/取消訂位 通知信給客戶
'傳入參數  ：[p_str_class_email_url:String] 信件內容的url
'            [p_str_email_to_addr:String] 收件者的電子信箱
'            [p_str_email_to_name:String] 收件者的名字 (允許空值)
'            [p_str_class_msg:String] 新增或取消課程訊息
'            [p_str_client_website:String] 'T=TutorABC, C=VIPABC
'回傳值    ：
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2009/12/27 Created
'*****************************************************************************************/
    Dim str_send_email_body_url, int_send_email_result, int_website_code
    
    if ("C" = UCase(p_str_client_website)) then
        int_website_code = CONST_VIPABC_WEBSITE
        str_send_email_body_url = p_str_class_email_url & "?client_full_ename=" & Escape(p_str_email_to_name) & "&add_or_cancel_class_info=" & Escape(p_str_class_msg)    

    elseif ("T" = UCase(p_str_client_website)) then
        int_website_code = CONST_TUTORABC_WEBSITE
        str_send_email_body_url = p_str_class_email_url & "?BookName=" & Escape(p_str_email_to_name) & "&Bookdate=" & Escape(p_str_class_msg)
    end if
    int_send_email_result = sendEmailFromUrl(p_str_email_to_addr, p_str_email_to_name, str_class_email_subject, str_send_email_body_url, int_website_code)
end sub


function isCalenderYearMonthBiggerThanNow(ByVal p_int_cyear, p_int_cmonth)
'/****************************************************************************************
'描述      ：檢查行事曆上要顯示的日期是否大於系統日期
'傳入參數  ：[p_int_cyear:Integer] 行事曆上顯示的年
'            [p_int_cmonth:Integer] 行事曆上顯示的月
'回傳值    ：true: 大於 false:小於
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2009/12/27 Created
'*****************************************************************************************/
    if ((p_int_cyear > Year(now)) or (p_int_cyear = Year(now) and p_int_cmonth >= Month(now))) then
        isCalenderYearMonthBiggerThanNow = true
    else
        isCalenderYearMonthBiggerThanNow = false
    end if
end function


function isCalenderDateBiggerThanNow(ByVal p_int_cyear, ByVal p_int_cmonth, p_int_cday)
'/****************************************************************************************
'描述      ：檢查行事曆上要顯示的日期是否大於系統日期
'傳入參數  ：[p_int_cyear:Integer] 行事曆上顯示的年
'            [p_int_cmonth:Integer] 行事曆上顯示的月
'            [p_int_cday:Integer] 行事曆上顯示的天
'回傳值    ：true: 大於 false:小於
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2009/12/27 Created
'*****************************************************************************************/
    if ((p_int_cyear > Year(now)) or (p_int_cyear = Year(now) and p_int_cmonth > Month(now)) or (p_int_cyear = Year(now) and p_int_cmonth = Month(now) and p_int_cday >= Day(now))) then
        isCalenderDateBiggerThanNow = true
    else
        isCalenderDateBiggerThanNow = false
    end if
end function


function isContractBegin(ByVal p_str_class_date)
'/****************************************************************************************
'描述      ：檢查選到的訂課日期使否在合約的服務開始日之後
'傳入參數  ：[p_str_class_date:String] 訂課日期
'回傳值    ：true: 合約開始日之後 false:之前
'牽涉數據表：
'備註      ：str_service_start_date 定義在 include/include_prepare_and_check_data.inc
'歷程(作者/日期/原因) ：[Ray Chien] 2010/01/11 Created
'*****************************************************************************************/
    if (DateDiff("d", str_service_start_date, p_str_class_date) >= 0) then
        isContractBegin = true
    else
        isContractBegin = false
    end if
end function

function isTemporalContractNoConfirm(ByVal p_str_purchase_contract_sn)
'/****************************************************************************************
'描述      ：檢查客戶的電子合約是否未確認(如果有電子合約存在的話且和client_purahcse.contract_sn一樣)
'傳入參數  ：[p_str_purchase_contract_sn:String] client_purchase.contract_sn
'回傳值    ：true: 電子合約未確認  false: 沒有電子合約或已確認
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/01/15 Created
'*****************************************************************************************/
    Dim bol_result

    bol_result = false

    '判斷合約未確認不能上課 (client_temporal_contract.sn = client_purchase.contract_sn)
    if (Not isEmptyOrNull(Session("contract_sn")) and Not isEmptyOrNull(p_str_purchase_contract_sn)) then
        if (CStr(Session("contract_sn")) = CStr(p_str_purchase_contract_sn)) then
            bol_result = true
        else
            bol_result = false
        end if
    end if
    isTemporalContractNoConfirm = bol_result
end function

function isAlreadyLobbyReservation(ByVal p_str_reservation_date, ByVal p_str_reservation_time, ByVal p_str_lobby_sn)
'/****************************************************************************************
'描述      ：檢查客戶是否已經預定了此時段的課程
'傳入參數  ：[p_str_reservation_date:String] 訂課日期
'            [p_str_reservation_time:String] 訂課時間
'回傳值    ：true: 已經預定  false: 沒有
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/01/15 Created
'*****************************************************************************************/
    Dim str_sql, var_arr_param_value, arr_attend_list_data, int_result

    int_result = false

    '檢查是否有此時段是否已經訂過
	str_sql = "SELECT sn FROM client_attend_list WHERE " &_
              "(client_sn = @client_sn) AND (attend_date = @attend_date) AND (attend_sestime = @attend_sestime) AND (special_sn = @special_sn) AND (valid = 1)"
    var_arr_param_value = Array(g_var_client_sn, p_str_reservation_date, p_str_reservation_time, p_str_lobby_sn)
    'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
    arr_attend_list_data = excuteSqlStatementRead(str_sql, var_arr_param_value, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_attend_list_data)) then
        if (Ubound(arr_attend_list_data) >= 0) then
            int_result = true
        else
            int_result = false
        end if
    end if

    isAlreadyLobbyReservation = int_result
end function


function isAlreadyLobbyReservationJR(ByVal p_str_reservation_date, ByVal p_str_reservation_time, ByVal p_str_lobby_sn)
'/****************************************************************************************
'描述      ：檢查客戶是否已經預定了此時段的課程
'傳入參數  ：[p_str_reservation_date:String] 訂課日期
'            [p_str_reservation_time:String] 訂課時間
'回傳值    ：true: 已經預定  false: 沒有
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/01/15 Created
'*****************************************************************************************/
    Dim str_sql, var_arr_param_value, arr_attend_list_data, int_result

    int_result = false

    '檢查是否有此時段是否已經訂過
	'str_sql = "SELECT sn FROM client_attend_list WHERE " &_
    '          "(client_sn = @client_sn) AND (attend_date = @attend_date) AND (attend_sestime = @attend_sestime) AND (special_sn = @special_sn) AND (valid = 1)"
    str_sql = " SELECT ClassRecordId FROM member.ClassRecord "
    str_sql = str_sql & " WHERE (client_sn = @client_sn) AND (ClassStartDate = @ClassStartDate) AND (ClassStartTime = @ClassStartTime) AND (ClassTypeId = 3) AND (ClassStatusId = 2) "
    var_arr_param_value = Array(g_var_client_sn, p_str_reservation_date, p_str_reservation_time, p_str_lobby_sn)
    'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
    arr_attend_list_data = excuteSqlStatementRead(str_sql, var_arr_param_value, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_attend_list_data)) then
        if (Ubound(arr_attend_list_data) >= 0) then
            int_result = true
        else
            int_result = false
        end if
    end if

    isAlreadyLobbyReservationJR = int_result
end function

function isAlreadyReservation(ByVal p_str_reservation_date, ByVal p_str_reservation_time)
'/****************************************************************************************
'描述      ：檢查客戶是否已經預定了此時段的課程
'傳入參數  ：[p_str_reservation_date:String] 訂課日期
'            [p_str_reservation_time:String] 訂課時間
'回傳值    ：true: 已經預定  false: 沒有
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/01/15 Created
'*****************************************************************************************/
    Dim str_sql, var_arr_param_value, arr_attend_list_data, int_result

    int_result = false

    '檢查是否有此時段是否已經訂過
	str_sql = "SELECT sn FROM client_attend_list WHERE " &_
              "(client_sn = @client_sn) AND (attend_date = @attend_date) AND (attend_sestime = @attend_sestime) AND (valid = 1)"
    var_arr_param_value = Array(g_var_client_sn, p_str_reservation_date, p_str_reservation_time)
    arr_attend_list_data = excuteSqlStatementRead(str_sql, var_arr_param_value, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_attend_list_data)) then
        if (Ubound(arr_attend_list_data) >= 0) then
            int_result = true
        else
            int_result = false
        end if
    end if

    isAlreadyReservation = int_result
end function

function isAlreadyReservationJR(ByVal p_str_reservation_date, ByVal p_str_reservation_time)
'/****************************************************************************************
'描述      ：檢查客戶是否已經預定了此時段的課程
'傳入參數  ：[p_str_reservation_date:String] 訂課日期
'            [p_str_reservation_time:String] 訂課時間
'回傳值    ：true: 已經預定  false: 沒有
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/01/15 Created
'*****************************************************************************************/
    Dim str_sql, var_arr_param_value, arr_attend_list_data, int_result

    int_result = false

    '檢查是否有此時段是否已經訂過
	'str_sql = "SELECT sn FROM client_attend_list WHERE " &_
    '          "(client_sn = @client_sn) AND (attend_date = @attend_date) AND (attend_sestime = @attend_sestime) AND (valid = 1)"
    str_sql = " SELECT ClassRecordId FROM member.ClassRecord "
    str_sql = str_sql & " WHERE (client_sn = @client_sn) AND (ClassStartDate = @ClassStartDate) AND (ClassStartTime = @ClassStartTime) AND (ClassStatusId = 2) "
    var_arr_param_value = Array(g_var_client_sn, p_str_reservation_date, p_str_reservation_time)
    arr_attend_list_data = excuteSqlStatementRead(str_sql, var_arr_param_value, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_attend_list_data)) then
        if (Ubound(arr_attend_list_data) >= 0) then
            int_result = true
        else
            int_result = false
        end if
    end if

    isAlreadyReservationJR = int_result
end function

function isCheckSpecialRule()
'/****************************************************************************************
'描述      ：是否檢查特定規則 觀察課表頁面不檢查
'傳入參數  ：
'回傳值    ：true: 檢查  false: 不檢查
'牽涉數據表：
'備註      ：	guide_EnvirVolumn_ok.asp 為電腦測試通過時不秀alert訊息
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/11 Created
'*****************************************************************************************/
    Dim str_path_info, bol_result

    bol_result = true
    str_path_info = Request.ServerVariables("PATH_INFO")

    if (str_path_info = "/program/member/reservation_class/include/ajax_class_list_calender_head.asp" _
        or str_path_info = "/program/member/reservation_class/include/ajax_class_list_calender.asp" _
        or str_path_info = "/program/member/reservation_class/include/ajax_class_list_week.asp" _
        or str_path_info = "/program/member/reservation_class/include/ajax_class_list_week_head.asp" _
		or str_path_info = "/program/member/environment_test/guide_EnvirVolumn_ok.asp" _
		or str_path_info = "/program/member/learning_dcgs_exe.asp" _
		or str_path_info = "/program/member/first_session_exe.asp" _
        or str_path_info = "/program/member/reservation_class/watch_class.asp") then

        bol_result = false
    end if

    isCheckSpecialRule = bol_result
end function


function getHtmlSelectClassTime(ByVal p_str_empty_msg, ByVal p_str_default_opt_value, ByVal p_str_order_class_date, ByVal p_int_class_type, ByVal p_str_client_sn, ByVal p_last_first_session_datetime, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：回傳HTML控制項<select> 可供客戶訂課的時間
'傳入參數  ：[p_str_empty_msg:String] 第一個<option value="">要顯示的Text
'            [p_str_default_opt_value:String] 預設要 selected 的 <option> 的 value EX:16 (表示預設要顯示 16:30) (允許空值)
'            [p_str_order_class_date:String] 客戶訂課日期
'            [p_int_class_type:Integer] 一對一 OR 一對多 OR 大會堂
'            [p_str_client_sn:Integer] 客戶編號
'            [p_last_first_session_datetime:String] 最近一次預約First Session的日期和時間
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：成功: 回傳HTML控制項<select> 失敗: 空字串
'牽涉數據表： 
'備註      ：預設 1. id = sel_class_time
'                 2. name = sel_class_time
'                 3. class = class_select
'                 4. 若沒有傳入 p_str_order_class_date 則會回傳的<select>會是系統預設的可上課時間
'                 5. 回傳-2 代表 四小時前沒有開課
'                 6. 回傳-3 24小時內臨時訂課僅開放一對多
'                 7. 回傳-99 代表選到了 人事部門規定之不開放上課之日期 或 選到的日期並無可供選擇的時段
'歷程(作者/日期/原因) ：[Ray Chien] 2009/12/1 Created
'*****************************************************************************************/
    Dim str_select_elm, str_now_sys_datetime, int_now_sys_hour, i, int_start_time, int_end_time, str_selected
    Dim str_sql, arr_attend_list_data, arr_param_value, arr_company_calendar, str_select_empty
    Dim str_booking_time : str_booking_time = ""                        '客戶已經訂課的時間 (data format = ",13,23,")
    Dim arr_block_time(24)                                              '人事部門規定之不開放上課之日期的時間 EX: ,1,2,
    Dim str_holiday_description : str_holiday_description = ""          '人事部門規定之不開放上課之日期的放假的描述(原因)
    Dim int_total_open_time_count : int_total_open_time_count = 0       '開放的時段數目
    Dim str_last_first_session_hour : str_last_first_session_hour = 0   '檢查客戶是否是首次上課且有無預約First Session，若有預約的話只能訂一小時後的課
    Dim str_select_width : str_select_width = "style=""width:160px;"""  'Select css style
    Dim int_no_order_time_in_minute : int_no_order_time_in_minute = 0   '若選到當日且系統時間(hr) + 幾分鐘內不能訂課 + 1 超過 23點，則表示今日無法訂課
    Dim arr_product_session_rule, int_now_sys_minute

    'TODO 開課前四小時取消 module
    'TODO 其他 module

    str_now_sys_datetime = now()
    int_now_sys_hour = Hour(str_now_sys_datetime)
    int_now_sys_minute = Minute(str_now_sys_datetime)
    str_select_elm = ""
    str_selected = ""
    int_start_time = -1  '可供客戶訂課的開始時間
    int_end_time = -1    '可供客戶訂課的結束時間

    if (Not isEmptyOrNull(p_str_empty_msg)) then
        if (Not isEmptyOrNull(p_str_order_class_date) and IsDate(p_str_order_class_date) and Not isEmptyOrNull(p_int_class_type)) then

            '檢查訂課日期是否為 人事部門規定之不開放上課之日期 取得時間及描述
            Call getCompanyHolidayInfo(p_str_order_class_date, str_holiday_description, arr_block_time, p_int_connect_type)
          
            '取得 產品對應課程類型的扣堂規則
            arr_product_session_rule = getProductSessionRule(str_now_use_product_sn, CInt(p_int_class_type), p_int_connect_type)
            int_no_order_time_in_minute = ((int_now_sys_hour*60) + int_now_sys_minute + CInt(arr_product_session_rule(1)) + 60)

            '檢查客戶選到的訂課日期是否 等於 目前系統日期
            if (DateDiff("d",DateValue(str_now_sys_datetime), DateValue(p_str_order_class_date)) = 0) then

                '若選到當日且系統時間(hr) + 幾分鐘內不能訂課 + 1 超過 23點，則表示今日無法訂課
                if (int_no_order_time_in_minute >= 1380) then
                    '大於一小時
                    if (CInt(arr_product_session_rule(1)) >= 60) then
                        '小时前没有开课，请选择其他上课日期或咨询时段
                        str_select_empty = Round(arr_product_session_rule(1) / 60, 1) &  getWord("CLASS_NO_OPEN_1") & "，" & getWord("SELECT_OTHER_CLASS_DATE_OR_TIME")
                    '一小時以下
                    else
                        '分钟前没有开课，请选择其他上课日期或咨询时段
                        str_select_empty = arr_product_session_rule(1) & getWord("CLASS_NO_OPEN_IN_MINUTE") & "，" & getWord("SELECT_OTHER_CLASS_DATE_OR_TIME")
                    end if                  

                    str_select_elm = "<select id=""sel_class_time"" name=""sel_class_time"" class=""class_select"" "&str_select_width&" title="""&str_select_empty&""">"        
                    str_select_elm = str_select_elm & "<option value=""-2"">" & str_select_empty & "</option>"    '"N小时前没有开课，请选择其他上课日期或咨询时段"
                    str_select_elm = str_select_elm & "</select>"
                else
                    int_start_time = int_no_order_time_in_minute \ 60
                    int_end_time = 23
                end if
                
            '檢查客戶選到的訂課日期是否 大於 目前系統日期+1
            elseif (DateDiff("d",DateValue(str_now_sys_datetime), DateValue(p_str_order_class_date)) = 1) then   
                int_start_time = (int_no_order_time_in_minute \ 60)
                if (int_start_time >= 24) then
                    int_start_time = int_start_time - 24
                    int_end_time = 23
                else
                    int_start_time = 0
                    int_end_time = 23                
                end if

            '檢查客戶選到的訂課日期是否 大於 目前系統日期+1之後
            elseif (DateDiff("d",DateValue(str_now_sys_datetime), DateValue(p_str_order_class_date)) > 1) then
                int_start_time = 0
                int_end_time = 23

            '例外的錯誤 客戶選到系統目前日期之前的 (前台就有擋)
            'else
                
            end if
        else
            int_start_time = 0
            int_end_time = 23
        end if

        if (int_start_time <> -1) then           
            '若有傳入日期 則檢查該日期客戶是否已有訂課
            str_booking_time = getCustomerBookingTime(p_str_order_class_date, p_str_client_sn, p_int_connect_type)          

            '檢查客戶是否是首次上課且有無預約First Session，若有預約的話只能訂一小時後的課
            if (isDate(p_last_first_session_datetime)) then
                if (DateDiff("d",DateValue(p_last_first_session_datetime), DateValue(p_str_order_class_date)) = 0) then
                    str_last_first_session_hour = Hour(p_last_first_session_datetime) + 1
                end if
            end if

            str_select_elm = "<select id=""sel_class_time"" name=""sel_class_time"" class=""class_select"" "&str_select_width&">"
            str_select_elm = str_select_elm & "<option value="""">"&p_str_empty_msg&"</option>"
            for i = int_start_time to int_end_time
                '若選到的日期客戶有訂課的話 則不顯示
                if (isEmptyOrNull(str_booking_time) or (Not isEmptyOrNull(str_booking_time) and InStr(str_booking_time, ","&i&",") <= 0)) then
                    '檢查是否要預設值
                    if (Not isEmptyOrNull(p_str_default_opt_value)) then
                        str_selected = ""
                        if (CStr(i) = CStr(p_str_default_opt_value)) then
                            str_selected = "selected=""selected"""
                        end if
                    end if

                    '該時段有開課
                    if (isEmptyOrNull(arr_block_time(i))) then
                        '咨訊時段請將 00:30～5:30此時段拿掉
                        if (i > 5) then
                            if (i >= str_last_first_session_hour) then
                                str_select_elm = str_select_elm & " <option value=""" & Right("0"&i, 2) & """ "&str_selected&">" & Right("0"&i, 2) & ":30</option>"                    
                                int_total_open_time_count = int_total_open_time_count + 1
                            end if
                        end if
                    end if
                end if
            next
            str_select_elm = str_select_elm & "</select>"

            '若沒有任何開課時段可供選擇 則顯示 请选择其他上课日期
            if (int_total_open_time_count = 0) then
                if (str_holiday_description <> "") then
                    str_select_empty = str_holiday_description
                else
                    str_select_empty = getWord("SELECT_OTHER_CLASS_DATE")
                end if
                str_select_elm = ""
                str_select_elm = "<select id=""sel_class_time"" name=""sel_class_time"" class=""class_select"" "&str_select_width&" title="""&str_select_empty&""">"        
                str_select_elm = str_select_elm & "<option value=""-99"">" & str_select_empty & "</option>"    '"请选择其他上课日期"
                str_select_elm = str_select_elm & "</select>"
            end if

        end if
    else
        'TODO 錯誤代碼定義
        g_int_run_time_err_code = CONST_ERROR_CODE_PARAM_VALID_FAILED
    end if

    getHtmlSelectClassTime = str_select_elm
end function


function getSelectableClassTime(ByVal p_str_order_class_date, ByVal p_int_class_type, ByVal p_str_client_sn, ByVal p_last_first_session_datetime, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：回傳HTML控制項<select> 可供客戶訂課的時間
'傳入參數  ：[p_str_order_class_date:String] 客戶訂課日期
'            [p_int_class_type:Integer] 一對一 OR 一對多 OR 大會堂
'            [p_str_client_sn:Integer] 客戶編號
'            [p_last_first_session_datetime:String] 最近一次預約First Session的日期和時間
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：成功: 回傳HTML控制項<select> 失敗: 空字串
'牽涉數據表： 
'備註      ：預設 1. id = sel_class_time
'                 2. name = sel_class_time
'                 3. class = class_select
'                 4. 若沒有傳入 p_str_order_class_date 則會回傳是空值
'                 5. 回傳-2 代表 四小時前沒有開課
'                 6. 回傳-3 24小時內臨時訂課僅開放一對多
'                 7. 回傳-99 代表選到了 人事部門規定之不開放上課之日期 或 選到的日期並無可供選擇的時段
'歷程(作者/日期/原因) ：[Lucy Chen] 2010/03/09 從getHtmlSelectClassTime這個function修改過來的
'*****************************************************************************************/
    Dim str_select_elm, str_now_sys_datetime, int_now_sys_hour, i, int_start_time, int_end_time, str_selected
    Dim str_sql, arr_attend_list_data, arr_param_value, arr_company_calendar, str_select_empty
    Dim str_booking_time : str_booking_time = ""                        '客戶已經訂課的時間 (data format = ",13,23,")
    Dim arr_block_time(24)                                              '人事部門規定之不開放上課之日期的時間 EX: ,1,2,
    Dim str_holiday_description : str_holiday_description = ""          '人事部門規定之不開放上課之日期的放假的描述(原因)
    Dim int_total_open_time_count : int_total_open_time_count = 0       '開放的時段數目
    Dim str_last_first_session_hour : str_last_first_session_hour = 0   '檢查客戶是否是首次上課且有無預約First Session，若有預約的話只能訂一小時後的課
    Dim str_select_width : str_select_width = "style=""width:160px;"""  'Select css style
    Dim int_no_order_time_in_minute : int_no_order_time_in_minute = 0   '若選到當日且系統時間(hr) + 幾分鐘內不能訂課 + 1 超過 23點，則表示今日無法訂課
    Dim arr_product_session_rule, int_now_sys_minute

	'客戶資料 COM 物件
	'Dim obj_member_opt
	'Dim int_member_result
	
    'TODO 開課前四小時取消 module
    'TODO 其他 module
	
	getSelectableClassTime = ""
	
    str_now_sys_datetime = now()
    int_now_sys_hour = Hour(str_now_sys_datetime)
    int_now_sys_minute = Minute(str_now_sys_datetime)
    str_select_elm = ","
    str_selected = ""
    int_start_time = -1  '可供客戶訂課的開始時間
    int_end_time = -1    '可供客戶訂課的結束時間
	
	Dim str_now_use_product_sn		'客戶合約的產品序號
	
	if (Not isEmptyOrNull(p_str_order_class_date) and IsDate(p_str_order_class_date) and Not isEmptyOrNull(p_int_class_type)) then
		
		'客戶資料 COM 物件
		Dim obj_member_opt : Set obj_member_opt = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
		Dim int_member_result : int_member_result = obj_member_opt.prepareData(g_var_client_sn, CONST_VIPABC_RW_CONN)
		if (int_member_result = CONST_FUNC_EXE_SUCCESS) then
			str_now_use_product_sn = obj_member_opt.getData(CONST_CONTRACT_PRODUCT_SN) '取得客戶合約的產品序號
		else
			Set obj_member_opt = Nothing
			exit Function
		end if
		Set obj_member_opt = Nothing

		If (IsNull(str_now_use_product_sn) or IsEmpty(str_now_use_product_sn) or str_now_use_product_sn = "") Then
			exit Function
		End If
		'TODO: test
		str_now_use_product_sn = 12
		
		'檢查訂課日期是否為 人事部門規定之不開放上課之日期 取得描述及時間 (str_holiday_description, arr_block_time)
		Call getCompanyHolidayInfo(p_str_order_class_date, str_holiday_description, arr_block_time, p_int_connect_type)
		'取得 產品對應課程類型的扣堂規則
		arr_product_session_rule = getProductSessionRule(str_now_use_product_sn, CInt(p_int_class_type), p_int_connect_type)
		If (IsNull(arr_product_session_rule)) Then
			exit Function
		End If
		
		int_no_order_time_in_minute = ((int_now_sys_hour*60) + int_now_sys_minute + CInt(arr_product_session_rule(1)) + 60)
		
		'檢查客戶選到的訂課日期是否 等於 目前系統日期
		if (DateDiff("d",DateValue(str_now_sys_datetime), DateValue(p_str_order_class_date)) = 0) then

			'若選到當日且系統時間(hr) + 幾分鐘內不能訂課 + 1 超過 23點，則表示今日無法訂課
			if (int_no_order_time_in_minute >= 1380) then
				'do nothing
				'exit Function
				int_start_time = -1
			else
				int_start_time = int_no_order_time_in_minute \ 60
				int_end_time = 23
			end if
			
		'檢查客戶選到的訂課日期是否 大於 目前系統日期+1
		elseif (DateDiff("d",DateValue(str_now_sys_datetime), DateValue(p_str_order_class_date)) = 1) then   
			int_start_time = (int_no_order_time_in_minute \ 60)
			if (int_start_time >= 24) then
				int_start_time = int_start_time - 24
				int_end_time = 23
			else
				int_start_time = 0
				int_end_time = 23                
			end if

		'檢查客戶選到的訂課日期是否 大於 目前系統日期+1之後
		elseif (DateDiff("d",DateValue(str_now_sys_datetime), DateValue(p_str_order_class_date)) > 1) then
			int_start_time = 0
			int_end_time = 23

		'例外的錯誤 客戶選到系統目前日期之前的 (前台就有擋)
		else
			'do nothing
			'exit Function
			int_start_time = -1
		end if
	else
		exit Function
	end if

	
	if (int_start_time <> -1) then           
		'若有傳入日期 則檢查該日期客戶是否已有訂課
		str_booking_time = getCustomerBookingTime(p_str_order_class_date, p_str_client_sn, p_int_connect_type)          

		'檢查客戶是否是首次上課且有無預約First Session，若有預約的話只能訂一小時後的課
		if (isDate(p_last_first_session_datetime)) then
			if (DateDiff("d",DateValue(p_last_first_session_datetime), DateValue(p_str_order_class_date)) = 0) then
				str_last_first_session_hour = Hour(p_last_first_session_datetime) + 1
			end if
		end if

		for i = int_start_time to int_end_time

		'若選到的日期客戶有訂課的話 則不顯示
			if (isEmptyOrNull(str_booking_time) or (Not isEmptyOrNull(str_booking_time) and InStr(str_booking_time, ","&i&",") <= 0)) then

				'該時段有開課
				if (isEmptyOrNull(arr_block_time(i))) then
					
					'咨訊時段請將 00:30～5:30此時段拿掉
					if (i > 5) then
						if (i >= str_last_first_session_hour) then
							str_select_elm = str_select_elm & Right("0"&i, 2) & ","
							int_total_open_time_count = int_total_open_time_count + 1
						end if
					end if
				end if
			end if
		next

	end if
	
    getSelectableClassTime = str_select_elm
end function


function getHtmlSelectLevel(ByVal p_str_empty_msg, ByVal p_str_default_opt_value, ByVal p_str_disable_lev)
'/****************************************************************************************
'描述      ：回傳HTML控制項<select> 可供選擇的等級
'傳入參數  ：[p_str_empty_msg:String] 第一個<option value="">要顯示的Text
'            [p_str_default_opt_value:String] 預設要 selected 的 <option> 的 value EX:"1" (表示預設要顯示 1) (允許空值)
'            [p_str_disable_lev:String] 不要顯示的 <option> 的 value 多個用","隔開 EX:"2,3" (表示不顯示 2 和 3) (允許空值)
'回傳值    ：成功: 回傳HTML控制項<select> 失敗: 空字串
'牽涉數據表： 
'備註      ：預設 1. id = sel_level
'                 2. name = sel_level
'                 3. class = class_select
'                 4. 預設產生 1~12 lev
'歷程(作者/日期/原因) ：[Ray Chien] 2009/12/10 Created
'*****************************************************************************************/
    Dim int_default_lev, str_arr_tmp, str_select_elm, str_selected, i, str_disable_lev, bol_enable_opt

    str_select_elm = ""
    str_selected = ""
    str_disable_lev = ""
    int_default_lev = 0
    bol_enable_opt = true

    if (Not isEmptyOrNull(p_str_empty_msg)) then
        str_select_elm = "<select id=""sel_level"" name=""sel_level"" class=""class_select"" >"
        str_select_elm = str_select_elm & "<option value="""">"&p_str_empty_msg&"</option>"

        '是否有設定預設要顯示的等級
        if (Not isEmptyOrNull(p_str_default_opt_value)) then
            int_default_lev = Int(Trim(p_str_default_opt_value))
        end if

        '是否有設定不要顯示的等級
        if (Not isEmptyOrNull(p_str_disable_lev)) then
            '原先的字串可能是 "1,2" 必須在前後加上 "," 這樣跑迴圈的時候就可以用 ",1," 這樣的字串來search了
            str_disable_lev = "," & Trim(p_str_disable_lev) & ","
        end if

        for i = 1 To 12
            bol_enable_opt = true

            str_selected = ""
            if (int_default_lev <> 0 and int_default_lev = i) then
                str_selected = "selected=""selected"""
            end if

            '檢查是否需要隱藏<option>
            if (Not isEmptyOrNull(str_disable_lev)) then
                if (InStr(str_disable_lev, ","&i&",") > 0) then
                    bol_enable_opt = false
                end if
            end if

            if (bol_enable_opt) then
                str_select_elm = str_select_elm & " <option value=""" & i & """ "&str_selected&">" & i & "</option>"
            end if
        next

        str_select_elm = str_select_elm & "</select>"
    else
        'TODO 錯誤代碼定義
        g_int_run_time_err_code = CONST_ERROR_CODE_PARAM_VALID_FAILED
    end if

    getHtmlSelectLevel = str_select_elm
end function


function getCustomerBookingTime(ByVal p_str_order_class_date, ByVal p_str_client_sn, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：則檢查該日期客戶是否已有訂課 並取得訂課的時間
'傳入參數  ：[p_str_order_class_date:String] 客戶訂課日期
'            [p_str_client_sn:Integer] 客戶編號
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：訂課的時間 (Format: ",6,16,15,")
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/19 Created
'*****************************************************************************************/
    Dim str_sql, arr_param_value, arr_attend_list_data, i, str_booking_time

    str_booking_time = ""

    '若有傳入日期 則檢查該日期客戶是否已有訂課
    if (Not isEmptyOrNull(p_str_order_class_date) and IsDate(p_str_order_class_date) and Not isEmptyOrNull(p_str_client_sn) and Not isEmptyOrNull(p_int_connect_type)) then            
        str_sql = " SELECT attend_sestime FROM client_attend_list " &_
                  " WHERE (valid = 1) AND (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, @sdate, 111)) " &_
                  " AND (client_attend_list.client_sn = @client_sn) "
        arr_param_value = Array(p_str_order_class_date, p_str_client_sn)
        arr_attend_list_data = excuteSqlStatementRead(str_sql, arr_param_value, p_int_connect_type)
        if (isSelectQuerySuccess(arr_attend_list_data)) then
            if (Ubound(arr_attend_list_data) >= 0) then
                for i = 0 to Ubound(arr_attend_list_data, 2)
                    str_booking_time = str_booking_time & arr_attend_list_data(0, i) & ","
                next
                str_booking_time = "," & str_booking_time
            end if
        end if
    end if
    getCustomerBookingTime = str_booking_time
end function

function getCustomerBookingTimeJR(ByVal p_str_order_class_date, ByVal p_str_client_sn, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：則檢查該日期客戶是否已有訂課 並取得訂課的時間
'傳入參數  ：[p_str_order_class_date:String] 客戶訂課日期
'            [p_str_client_sn:Integer] 客戶編號
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：訂課的時間 (Format: ",6,16,15,")
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/19 Created
'*****************************************************************************************/
    Dim str_sql, arr_param_value, arr_attend_list_data, i, str_booking_time

    str_booking_time = ""

    '若有傳入日期 則檢查該日期客戶是否已有訂課
    if (Not isEmptyOrNull(p_str_order_class_date) and IsDate(p_str_order_class_date) and Not isEmptyOrNull(p_str_client_sn) and Not isEmptyOrNull(p_int_connect_type)) then            
        'str_sql = " SELECT attend_sestime FROM client_attend_list " &_
        '          " WHERE (valid = 1) AND (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, @sdate, 111)) " &_
        '          " AND (client_attend_list.client_sn = @client_sn) "
        str_sql = " SELECT ClassStartTime FROM member.ClassRecord "
        str_sql = str_sql & " WHERE (client_sn = @client_sn) AND (ClassStartDate = @ClassStartDate) AND (ClassStatusId = 2) "
        arr_param_value = Array(p_str_order_class_date, p_str_client_sn)
        arr_attend_list_data = excuteSqlStatementRead(str_sql, arr_param_value, p_int_connect_type)
        if (isSelectQuerySuccess(arr_attend_list_data)) then
            if (Ubound(arr_attend_list_data) >= 0) then
                for i = 0 to Ubound(arr_attend_list_data, 2)
                    str_booking_time = str_booking_time & arr_attend_list_data(0, i) & ","
                next
                str_booking_time = "," & str_booking_time
            end if
        end if
    end if
    getCustomerBookingTimeJR = str_booking_time
end function

sub setNoOrderAndCancelMessage(ByVal p_int_product_sn, ByRef p_arr_msg_no_order_in_hour, ByRef p_arr_msg_no_cancel_in_hour, ByRef p_arr_no_cancel_in_minute, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：設定 一對一和一對三和大會堂 幾小時內不能訂課 和 不能取消 要顯示的訊息
'傳入參數  ：[p_int_product_sn:Integer] 產品編號
'            [p_arr_msg_no_order_in_hour:Reference] Array (0)一對一和(1)一對三和(2)大會堂幾小時內不能訂課 訊息
'            [p_arr_msg_no_cancel_in_hour:Reference] Array (0)一對一和(1)一對三和(2)大會堂幾小時內不能取消 訊息
'            [p_arr_no_cancel_in_minute:Reference] Array (0)一對一和(1)一對三和(2)大會堂幾分鐘內不能取消 minute
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/22 Created
'*****************************************************************************************/
    Dim arr_product_session_rule, i, str_no_order_msg, str_no_cancel_msg

    '取得 一對一和一對三和大會堂 幾小時內不能訂課 和 不能取消
    arr_product_session_rule = getProductSessionRule(p_int_product_sn, 0, p_int_connect_type)
    if (IsArray(arr_product_session_rule)) then
        for i = 0 to Ubound(arr_product_session_rule, 2)
            '大於一小時
            if (CInt(arr_product_session_rule(1, i)) >= 60) then
                '小时前没有开课，请选择其他上课日期或咨询时段
                str_no_order_msg = Round(arr_product_session_rule(1, i) / 60, 1) &  getWord("CLASS_NO_OPEN_1") & "，" & getWord("SELECT_OTHER_CLASS_DATE_OR_TIME")
            '一小時以下
            else
                '分钟前没有开课，请选择其他上课日期或咨询时段
                str_no_order_msg = arr_product_session_rule(1, i) & getWord("CLASS_NO_OPEN_IN_MINUTE") & "，" & getWord("SELECT_OTHER_CLASS_DATE_OR_TIME")
            end if

            '大於一小時
            if (CInt(arr_product_session_rule(2, i)) >= 60) then
                '开课前n小时内无法取消
                str_no_cancel_msg = getWord("CLASS_START") & Round(arr_product_session_rule(2, i) / 60, 1) & getWord("NO_CANCEL")
            '一小時以下
            else
                '开课前n分钟内无法取消
                str_no_cancel_msg = getWord("CLASS_START") & arr_product_session_rule(2, i) & getWord("NO_CANCEL_IN_MINUTE")
            end if

            '一對一
            if (CStr(arr_product_session_rule(3, i)) = CStr(CONST_CLASS_TYPE_ONE_ON_ONE)) then
                p_arr_msg_no_order_in_hour(0) = str_no_order_msg
                p_arr_msg_no_cancel_in_hour(0) = str_no_cancel_msg
                p_arr_no_cancel_in_minute(0) = arr_product_session_rule(2, i)
            '一對三
            elseif (CStr(arr_product_session_rule(3, i)) = CStr(CONST_CLASS_TYPE_ONE_ON_THREE)) then
                p_arr_msg_no_order_in_hour(1) = str_no_order_msg
                p_arr_msg_no_cancel_in_hour(1) = str_no_cancel_msg
                p_arr_no_cancel_in_minute(1) = arr_product_session_rule(2, i)
            '一對4
            elseif (CStr(arr_product_session_rule(3, i)) = CStr(CONST_CLASS_TYPE_ONE_ON_FOUR)) then
                p_arr_msg_no_order_in_hour(1) = str_no_order_msg
                p_arr_msg_no_cancel_in_hour(1) = str_no_cancel_msg
                p_arr_no_cancel_in_minute(1) = arr_product_session_rule(2, i)
            '大會堂
            elseif (CStr(arr_product_session_rule(3, i)) = CStr(CONST_CLASS_TYPE_LOBBY)) then
                p_arr_msg_no_order_in_hour(2) = str_no_order_msg
                p_arr_msg_no_cancel_in_hour(2) = str_no_cancel_msg
                p_arr_no_cancel_in_minute(2) = arr_product_session_rule(2, i)
            end if
        next
    end if
end sub

function getCustomerBookingTimeAndType(ByVal p_str_order_class_date, ByVal p_str_client_sn, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：則檢查該日期客戶是否已有訂課 並取得訂課的時間與訂哪一種課程
'傳入參數  ：[p_str_order_class_date:String] 客戶訂課日期
'            [p_str_client_sn:Integer] 客戶編號
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：訂課的時間與type (Format: {{"2010-03-10 00:00:00", 6}, {2010-03-10 00:00:00", 7}, {2010-03-10 00:00:00", 99}, ...etc})
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Lucy Chen] 2010/03/09 Created
'*****************************************************************************************/
    Dim str_sql, arr_param_value, arr_attend_list_data, i
	Dim arr_booking_data
	Dim arr_temp

    '若有傳入日期 則檢查該日期客戶是否已有訂課
    if (Not isEmptyOrNull(p_str_order_class_date) and IsDate(p_str_order_class_date) and Not isEmptyOrNull(p_str_client_sn) and Not isEmptyOrNull(p_int_connect_type)) then            
        str_sql = " SELECT attend_sestime, isNull(special_sn, 0) as special_sn, attend_livesession_types, sn " &_
                  " FROM client_attend_list " &_
                  " WHERE (valid = 1) AND (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, @sdate, 111)) " &_
                  " AND (client_attend_list.client_sn = @client_sn) " &_
				  " ORDER BY attend_sestime "
        arr_param_value = Array(p_str_order_class_date, p_str_client_sn)
		
		arr_attend_list_data = excuteSqlStatementRead(str_sql, arr_param_value, p_int_connect_type)
        if (isSelectQuerySuccess(arr_attend_list_data)) then
            if (Ubound(arr_attend_list_data) >= 0) then
			
				ReDim arr_booking_data(Ubound(arr_attend_list_data, 2))
				
                for i = 0 to Ubound(arr_attend_list_data, 2)
					ReDim arr_temp(2)
					arr_temp(0) = arr_attend_list_data(0, i)	'attend_sestime
					
					'special_sn有值: 此訂位資料為大會堂訂位
					If (arr_attend_list_data(1, i)<>0) Then
						arr_temp(1) = CONST_CLASS_TYPE_LOBBY
					Else
						'attend_livesession_types: 7 一對一
						If (arr_attend_list_data(2, i)=7) Then
							arr_temp(1) = CONST_CLASS_TYPE_ONE_ON_ONE
						Else
						'attend_livesession_types: 4 一對多
							arr_temp(1) = CONST_CLASS_TYPE_ONE_ON_THREE
						End If
					End If
					arr_temp(2) = arr_attend_list_data(3, i)
					arr_booking_data(i) = arr_temp
                next
            end if
        end if
    end if
	
    getCustomerBookingTimeAndType = arr_booking_data
end function


function calculateAvailablePoint(ByVal p_flt_available_points, ByVal p_flt_week_points, ByVal p_flt_deduct_session_points, ByVal p_flt_bonus_points, ByVal p_str_bonus_begin_datetime)
'/****************************************************************************************
'描述      ：大小錢包和要扣除的點數做運算的結果
'傳入參數  ：[p_flt_available_points:Float] 大錢包 = 已使用堂數(client_attend_list.valid=1) + return - 錄影檔使用
'            [p_flt_week_points:Float] 小錢包 = 本週剩餘可用的點數
'            [p_flt_deduct_session_points:Float] 要扣除的點數
'            [p_flt_bonus_points:Float] 紅利包
'            [p_str_bonus_begin_datetime:str] 紅利包開通時間
'回傳值    ：一維陣列 (0) = 大錢包，(1) = 小錢包, (2) = 紅利包, (3)=紅利包開通時間
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/04/20 Created
'*****************************************************************************************/
    Dim flt_session_deduct_point, flt_available_points, flt_week_points , flt_bouns_points
    Dim arr_point(3)
	Dim bolBuyPointsOver : bolBuyPointsOver = false

    flt_available_points = CDbl(p_flt_available_points)
    flt_week_points = CDbl(p_flt_week_points)
	flt_bouns_points = CDbl(p_flt_bonus_points)
	str_bonus_begin_datetime = p_str_bonus_begin_datetime
    arr_point(0) = 0
    arr_point(1) = 0
	arr_point(2) = 0
	arr_point(3) = ""
	
	
	'如果大錢包用完，則將紅利包暫時充當大錢包計算
	if (flt_available_points <= 0) then
		bolBuyPointsOver =true
		flt_available_points = flt_bouns_points
	else
		bolBuyPointsOver = false
	end if
	
	

    '預訂課程的扣點數
    flt_session_deduct_point = p_flt_deduct_session_points

    '小錢包沒有點數可以扣除，從大錢包扣除
    if (flt_week_points <= 0) then
        flt_available_points = flt_available_points - flt_session_deduct_point

    '小錢包有足夠的點數可以扣除
    elseif (flt_week_points > 0 and flt_week_points >= flt_session_deduct_point) then
        flt_week_points = flt_week_points - flt_session_deduct_point

    '小錢包沒有足夠的點數可以扣除，要從大錢包扣除一部分
    elseif (flt_week_points > 0 and flt_week_points < flt_session_deduct_point) then
        flt_session_deduct_point = flt_session_deduct_point - flt_week_points
        flt_week_points = 0
        flt_available_points = flt_available_points - flt_session_deduct_point
    end if
	
	'如果大錢包用完，再將暫時的大錢包從之前算出來的值塞回紅利包，大錢包改為0
	if (bolBuyPointsOver) then
		flt_bouns_points = flt_available_points
		flt_available_points = 0
	end if

    arr_point(0) = flt_available_points
    arr_point(1) = flt_week_points
	arr_point(2) = flt_bouns_points
	arr_point(3) = str_bonus_begin_datetime
    calculateAvailablePoint = arr_point
end function

function updatePointAfterReservationOrCancel(ByVal p_arr_reservation_session, ByVal p_arr_data, _                                         
                                             ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：1. 檢查訂課時間，當週和非當週的總扣堂數，更新到客戶剩餘點數總表
'            2. 檢查取消課程的時間，當週和非當週的，更新到客戶剩餘點數總表
'傳入參數  ：[p_arr_reservation_session:Integer] 成功預訂課程的日期和扣堂 (0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn
'            [p_arr_data:Array]
'               [0] = 客戶編號 client_basic.sn
'               [1] = client_purchase.account_sn
'               [2] = 產品編號
'               [3] = 結算週期
'               [4] = 每週期(期末)應扣除堂數
'               [5] = 預訂課程"reservation" or 取消課程"cancel"
'               [6] = 本週已使用堂數(排除return) (允許空值)
'               [7] = 來源網站:
'                       1 = TutorABC, 2 = TutorABCJr,  3 = TutorMing 
'                       4 = NTUtorMing,  5 = Columbia,  6 = VIPABC, 7=TutorABCIMS, 
'                       8 = TUTORABCJR_IMS, 9=TUTORMING_IMS, 10 = NTUTORMING_IMS,  11 = CLUE_IMS, 12 = VIPABC_IMS
'               [8] = 來源程式位置:
'               [9] = 員工編號
'               [10] = 服務開始日
'               [11] = 服務結束日
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：true or false
'牽涉數據表：customer_session_points
'備註      ：錯誤回報的等級 WARN
'            定時定額產品才使用此函式
'歷程(作者/日期/原因) ：[Ray Chien] 2010/04/19 Created
'*****************************************************************************************/
    Dim int_client_purchase_sn, int_client_sn, i, arr_session_cycle_info
    Dim arr_reservation_session, arr_customer_session_points, arr_client_purchase, arr_cfg_product
    Dim arr_point, arr_update_data, bol_result, str_note, arr_client_attend_list
    Dim objConn, str_sql, objRs, arr_log_data, arr_log_total_data

    str_note = ""
    arr_customer_session_points = null

    '是否需要回報錯誤
    Dim is_report_error : is_report_error = false

    '小錢包 : 本週剩餘可用的點數 
    Dim flt_week_points : flt_week_points = 0

    '大錢包 : 可用點數  = 已使用堂數(client_attend_list.valid=1) + return - 錄影檔使用
    Dim flt_available_points : flt_available_points = 0
	
	'紅利包
	Dim flt_bonus_points : flt_bonus_points = 0
	
	'紅利包開通時間
	dim str_bonus_begin_datetime : str_bonus_begin_datetime = ""

    '預訂或取消課程預扣除的扣點數
    Dim int_session_deduct_point : int_session_deduct_point = 0

    '來源網站
    Dim int_source_website : int_source_website = 0

    '來源程式位置
    Dim str_source_code : str_source_code = ""

    '要回報的錯誤訊息
    Dim str_report_error_msg : str_report_error_msg = ""

    '詳細的錯誤訊息描述
    Dim str_error_msg_detail_desc : str_error_msg_detail_desc = ""

    '每週期(期末)應扣除堂數
    Dim int_week_deduct_session : int_week_deduct_session = 0
    
    '本週已使用堂數(排除return)
    Dim int_week_use_session_count : int_week_use_session_count = 0

    '員工編號
    Dim int_staff_sn : int_staff_sn = 0

    '預訂課程"reservation" or 取消課程"cancel"
    Dim str_opt : str_opt = ""

    '產品編號
    Dim int_product_sn : int_product_sn = 0

    '結算週期
    Dim int_account_cycle_day : int_account_cycle_day = 0

    '服務開始日
    Dim dat_contract_service_sdate : dat_contract_service_sdate = ""

    '服務結束日
    Dim dat_contract_service_edate : dat_contract_service_edate = ""

    '何種情況下新增的資料
    '1.帳號開通 
    '2.系統排程(週期初、週期末檢查) 
    '3.首次訂課 
    '4.訂當週課程，扣除小錢包csh_week_points
    '5.預訂下週或未來課程，扣除大錢包csh_total_points
    '6.取消本週課程，歸還到大錢包
    '7.取消下週或未來課程，歸還到大錢包或小錢包
    '8.return，歸還到大錢包
    '9.退費時
    '10. 觀看錄影檔時，從大錢包扣除點數
    Dim int_csh_type : int_csh_type = 0


    if (IsArray(p_arr_reservation_session) and _
        IsArray(p_arr_data) and _
        Not isEmptyOrNull(p_int_connect_type)) then

        int_client_sn = p_arr_data(0)
        int_client_purchase_sn = p_arr_data(1)
        int_product_sn = p_arr_data(2)
        int_account_cycle_day = p_arr_data(3)
        int_week_deduct_session = p_arr_data(4)
        str_opt = p_arr_data(5)
        int_week_use_session_count = p_arr_data(6)
        int_source_website = p_arr_data(7)
        str_source_code = p_arr_data(8)
        int_staff_sn = p_arr_data(9)
        dat_contract_service_sdate = p_arr_data(10)
        dat_contract_service_edate = p_arr_data(11)
        

        '取得成功預訂課程的日期
        arr_reservation_session = p_arr_reservation_session
        if (Ubound(arr_reservation_session) >= 0) then

                'if (str_opt = "reservation") then
                    ''檢查是否需要 初始化 客戶剩餘點數總表 和 大小錢包異動紀錄資料表
                    ''並取得 客戶剩餘點數總表 和 大小錢包異動紀錄資料表 是否有資料
                    'arr_customer_session_points = checkCustomerSessionPointsExist(int_client_purchase_sn, int_client_sn, _
                                                                                  'int_source_website, str_source_code, _
                                                                                  'p_int_connect_type)
                'else
                    '取得 客戶剩餘點數總表 和 大小錢包異動紀錄資料表 是否有資料
                    'arr_customer_session_points = getCustomerSessionPoints(int_client_purchase_sn, int_client_sn, p_int_connect_type)                    
                'end if

                'if (Not IsArray(arr_customer_session_points)) then
                    'is_report_error = true
                    'if (str_opt = "reservation") then
                        'str_error_msg_detail_desc = arr_customer_session_points
                    'else
                        'str_error_msg_detail_desc = "<br/>取得 客戶剩餘點數總表 和 大小錢包異動紀錄資料表時，失敗!!<br/>"
                    'end if
                'end if
            
                Set objConn = Server.CreateObject("ADODB.Connection")
	            objConn.connectionString = g_strVipWriteConnString
	            objConn.open
	    	    
                ' Begin Transtraction
                objConn.BeginTrans

                str_sql = " SELECT "
                '0
                str_sql = str_sql & " csp_sn, "
                '1 合約總點數  client_purchase.add_points(產品總點數) + client_purchase.access_points(贈送點數)
                str_sql = str_sql & " csp_total_points, "
                '2 大錢包 : 可用點數  = 已使用堂數(client_attend_list.valid=1) + return - 錄影檔使用 + 小錢包        
                str_sql = str_sql & " ISNULL(csp_available_points, 0) AS csp_available_points, "
                '3 小錢包 : 本週剩餘可用的點數
                str_sql = str_sql & " ISNULL(csp_week_points, 0) AS csp_week_points, "
                '4 週期初小錢包的值，每週週期初會更新
               	str_sql = str_sql & " ISNULL(csp_week_start_points, 0) AS csp_week_start_points, "
				'5 紅利包 : 贈送的堂數，當大錢包用完後才扣
				str_sql = str_sql & " ISNULL(csp_bonus_points, 0) AS csp_bonus_points, "
				'6 紅利包開通時間 : 大錢包用完的當天
				str_sql = str_sql & " csp_bonus_begin_datetime "
                str_sql = str_sql & " FROM customer_session_points "
                str_sql = str_sql & " WHERE (csp_client_sn = '"&int_client_sn&"') AND (csp_purchase_sn = '"&int_client_purchase_sn&"') "

                '取得舊的剩餘點數
                Set objRs = Server.CreateObject("Adodb.RecordSet")
                objRs.CursorLocation = CONST_AD_USE_CLIENT
                objRs.CursorType = CONST_AD_OPEN_STATIC
                objRs.LockType = CONST_AD_LOCK_OPTIMISTIC
                objRs.open str_sql,objConn
                if ( not objRs.eof ) then
                    arr_customer_session_points  = objRs.GetRows()    
                end if
                objRs.close
                Set objRs = Nothing

                if (objConn.Errors.Count <> 0 or err.Number <> 0 or isEmptyOrNull(arr_customer_session_points)) then
                    str_error_msg_detail_desc = str_error_msg_detail_desc & "<br/>取得 客戶剩餘點數總表資訊時，失敗!!<br/>"
                    is_report_error = true
                    objConn.RollBackTrans
                end if
                
                if (false = is_report_error) then
                    '**** 將每筆訂課或取消課程資訊分當週和非當週，各別塞入 customer_session_points_log START ****
                    flt_available_points = CDbl(arr_customer_session_points(2, 0))
                    flt_week_points = CDbl(arr_customer_session_points(3, 0))
					flt_bonus_points = CDbl(arr_customer_session_points(5, 0))
					str_bonus_begin_datetime = arr_customer_session_points(6, 0)
                    Redim arr_log_total_data(Ubound(arr_reservation_session, 2))

                    for i = 0 to Ubound(arr_reservation_session, 2)
                        str_note = ""
                        int_session_deduct_point = 0
                        if (Not isEmptyOrNull(Trim(arr_reservation_session(0, i)))) then
                            '取得課程在結算週期內的資訊，只有大陸定時定額產品才需要判斷
                            arr_session_cycle_info = getSessionCycleDateInfo(Array(int_account_cycle_day, dat_contract_service_sdate), arr_reservation_session(0, i))
                            if (Not IsArray(arr_session_cycle_info)) then                            
                                is_report_error = true
                                str_error_msg_detail_desc = "<br/>判斷課程是否在結算週期內的資訊，只有大陸定時定額產品才需要判斷時，失敗!!<br/>"
                            else
                                '將課程的扣堂數換算成點數
                                int_session_deduct_point = CDbl(arr_reservation_session(1, i)) * CONST_SESSION_TRANSFORM_POINT

                                '預訂或取消當週課程
                                if (arr_session_cycle_info(0)) then
                                    '預訂課程
                                    if (str_opt = "reservation") then
                                        int_csh_type = 4
                                        str_note = str_note & "預訂當週課程, "
                                        '取得大小錢包運算後的值
                                        arr_point = calculateAvailablePoint(flt_available_points, flt_week_points, int_session_deduct_point, flt_bonus_points, str_bonus_begin_datetime)

                                        str_note = str_note & "扣除大錢包" & Round((flt_available_points - arr_point(0)) / CONST_SESSION_TRANSFORM_POINT, 2) & "堂, "
                                        str_note = str_note & "扣除小錢包" & Round((flt_week_points - arr_point(1)) / CONST_SESSION_TRANSFORM_POINT, 2) & "堂, "
										str_note = str_note & "扣除紅利包" & Round((flt_bonus_points - arr_point(2)) / CONST_SESSION_TRANSFORM_POINT, 2) & "堂, "

                                        flt_available_points = arr_point(0)
                                        flt_week_points = arr_point(1)
										flt_bonus_points = arr_point(2)
										str_bonus_begin_datetime = arr_point(3)
                                    '取消課程
                                    else
                                        int_csh_type = 6
                                        str_note = str_note & "取消當週課程, "
                                        
                                        '**** 取得產品的資訊 的 每週期(期末)應扣除堂數 START ****
                                        if (int_week_deduct_session = 0) then
                                            is_report_error = true
                                            str_error_msg_detail_desc = str_error_msg_detail_desc & "<br/>每週期(期末)應扣除堂數不該是空值!!<br/>"
                                        end if
                                        '**** 取得產品的資訊 的 每週期(期末)應扣除堂數 END ****

                                        '發生錯誤必須停止迴圈
                                        if (true = is_report_error) then
                                            exit for
                                        end if

                                        '取消課程規則:
                                        '1	取消本周課程取消確認時，先將”本週已使用堂數”(假設數值為X)，比對每週期(期末)應扣除堂數(假設其值為Y)：
                                        '1.1  若X <= Y，代表未有超額訂課或預訂課程的情況，此時將”所取消堂數” 立刻歸還到小錢包。
                                        '1.2  若X > Y，代表有超額訂課或預訂課程的情況，此時將X – Y(假設數值為W)，”所取消堂數”(假設數值為C)，有以下兩種情況:
                                        '1.21 若C <= W，則將”所取消堂數”的部份立刻歸還到大錢包。
                                        '1.22 若C > W，則將W的部份立刻歸還到大錢包，且將C – W的值立刻歸還到小錢包。

                                        if (int_week_use_session_count <= int_week_deduct_session) then
                                            '將堂數歸還到小錢包
                                            flt_week_points = flt_week_points + int_session_deduct_point
                                            str_note = str_note & "歸還" & arr_reservation_session(1, i) & "堂到小錢包, "
                                        else
                                            if (CDbl(arr_reservation_session(1, i)) <= (int_week_use_session_count - int_week_deduct_session)) then
                                                '將堂數歸還到大錢包或者已經開始使用紅利包則回到紅利包
												if ( flt_available_points <= 0 or not isEmptyOrNull(str_bonus_begin_datetime) ) then 
	                                                flt_bonus_points = flt_bonus_points + int_session_deduct_point
													str_note = str_note & "歸還" & arr_reservation_session(1, i) & "堂到紅利包, "
												else
													flt_available_points = flt_available_points + int_session_deduct_point
    	                                            str_note = str_note & "歸還" & arr_reservation_session(1, i) & "堂到大錢包, "
												end if
                                            else  
												'將堂數歸還到大錢包或者已經開始使用紅利包則回到紅利包
												if ( flt_available_points <= 0 or not isEmptyOrNull(str_bonus_begin_datetime) ) then
													flt_bonus_points = flt_bonus_points + (CDbl(int_week_use_session_count - int_week_deduct_session) * CONST_SESSION_TRANSFORM_POINT)
													str_note = str_note & "歸還" & CDbl(int_week_use_session_count - int_week_deduct_session) & "堂到紅利包, "
												else                                        
                                                	flt_available_points = flt_available_points + (CDbl(int_week_use_session_count - int_week_deduct_session) * CONST_SESSION_TRANSFORM_POINT)
													str_note = str_note & "歸還" & CDbl(int_week_use_session_count - int_week_deduct_session) & "堂到大錢包, "
												end if
												flt_week_points = flt_week_points + ((CDbl(arr_reservation_session(1, i)) - CDbl(int_week_use_session_count - int_week_deduct_session)) * CONST_SESSION_TRANSFORM_POINT)                                             
                                                str_note = str_note & "歸還" & (CDbl(arr_reservation_session(1, i)) - CDbl(int_week_use_session_count - int_week_deduct_session)) & "堂到小錢包, "
                                            end if
                                        end if
                                        int_week_use_session_count = int_week_use_session_count - CDbl(arr_reservation_session(1, i))
                                    end if

                                '預訂或取消下週或未來課程
                                else
                                    '預訂課程
                                    if (str_opt = "reservation") then
                                        int_csh_type = 5
                                        str_note = str_note & "預訂下週或未來的課程, "

                                        '預訂非本周課時，從大錢包扣除點數，紅利包開通後只能從紅利包扣除
										if (flt_available_points <= 0 or not isEmptyOrNull(str_bonus_begin_datetime) ) then
	                                        flt_bonus_points = flt_bonus_points - int_session_deduct_point
											str_note = str_note & "扣除紅利包" & arr_reservation_session(1, i) & "堂, "
										else
											flt_available_points = flt_available_points - int_session_deduct_point
											str_note = str_note & "扣除大錢包" & arr_reservation_session(1, i) & "堂, "
										end if
                                        
                                    '取消課程
                                    else
                                        int_csh_type = 7
                                        str_note = str_note & "取消下週或未來的課程, "
                                        '取消非本周課時，將堂數歸還到大錢包
										if (flt_available_points <= 0 or not isEmptyOrNull(str_bonus_begin_datetime) ) then
											flt_bonus_points = flt_bonus_points + int_session_deduct_point
	                                        str_note = str_note & "歸還" & arr_reservation_session(1, i) & "堂到紅利包, "
										else
											flt_available_points = flt_available_points + int_session_deduct_point
	                                        str_note = str_note & "歸還" & arr_reservation_session(1, i) & "堂到大錢包, "
										end if
                                        
                                    end if
                                end if

                                if (str_opt = "reservation") then                                   
                                    str_note = str_note & "總扣除的堂數: " & arr_reservation_session(1, i)
                                else
                                    str_note = str_note & "總歸還的堂數: " & arr_reservation_session(1, i)
                                end if

                                arr_log_data = Array(arr_customer_session_points(0, 0), int_client_sn, int_client_purchase_sn, _
                                                     flt_available_points, flt_week_points, int_csh_type, _
                                                     arr_reservation_session(2, i), "", str_note,  _
                                                     arr_session_cycle_info(1), arr_session_cycle_info(2), arr_session_cycle_info(3), _
                                                     int_source_website, str_source_code, flt_bonus_points, str_bonus_begin_datetime)
                                arr_log_total_data(i) = arr_log_data

                                '新增一筆資料到 大小錢包異動紀錄資料表 總剩餘堂數和每週使用堂數異動紀錄
                                'if (Not insertCustomerSessionPointsLog(arr_log_data, p_int_connect_type)) then
                                    'is_report_error = true
                                    'str_error_msg_detail_desc = str_error_msg_detail_desc & "<br/>新增一筆資料到 大小錢包異動紀錄資料表 總剩餘堂數和每週使用堂數異動紀錄時，失敗!!<br/>"
                                    'str_error_msg_detail_desc = str_error_msg_detail_desc & "失敗的client_attend_list.sn: " & arr_reservation_session(2, i)
                                'end if
                            end if 'end of 取得課程在結算週期內的資訊
                        end if
                    next
                    '更新一筆到 客戶剩餘點數總表
                    str_sql = " UPDATE customer_session_points SET csp_available_points = '"&flt_available_points&"', "
                    str_sql = str_sql & " csp_week_points = '"&flt_week_points&"', "
					str_sql = str_sql & " csp_bonus_points = '"&flt_bonus_points&"', "
                    str_sql = str_sql & " csp_last_modify_datetime = getdate() "
                    str_sql = str_sql & " WHERE (csp_sn = '"&arr_customer_session_points(0, 0)&"') AND "
                    str_sql = str_sql & " (csp_client_sn = '"&int_client_sn&"') AND (csp_purchase_sn = '"&int_client_purchase_sn&"')"
					objConn.Execute(str_sql)

                    if (objConn.Errors.Count = 0 And err.Number = 0) then
                        objConn.CommitTrans
                    else
                        str_error_msg_detail_desc = str_error_msg_detail_desc & "<br/>更新一筆到 客戶剩餘點數總表時，失敗!!<br/>"
                        is_report_error = true
                        objConn.RollBackTrans
                    end if

                    'arr_update_data = Array(arr_customer_session_points(0, 0), int_client_sn, int_client_purchase_sn, _
                                            'flt_available_points, flt_week_points)
                    'if (Not updateCustomerSessionPoints(arr_update_data, p_int_connect_type)) then
                        'is_report_error = true
                        'str_error_msg_detail_desc = str_error_msg_detail_desc & "<br/>更新一筆到 客戶剩餘點數總表時，失敗!!<br/>"
                    'end if
                    '**** 將每筆訂課或取消課程資訊分當週和非當週，各別塞入 customer_session_points_log END ****
                end if 'end of if (false = is_report_error) then

                objConn.close
                set objConn = Nothing
                '新增一筆資料到 大小錢包異動紀錄資料表 總剩餘堂數和每週使用堂數異動紀錄
                if (IsArray(arr_log_total_data)) then
                    for i = 0 to Ubound(arr_log_total_data)
                        if (IsArray(arr_log_total_data(i))) then
                            if (Not insertCustomerSessionPointsLog(arr_log_total_data(i), p_int_connect_type)) then
                                is_report_error = true
                                str_error_msg_detail_desc = str_error_msg_detail_desc & "<br/>新增一筆資料到 大小錢包異動紀錄資料表 總剩餘堂數和每週使用堂數異動紀錄時，失敗!!<br/>"
                                str_error_msg_detail_desc = str_error_msg_detail_desc & "失敗的client_attend_list.sn: " & arr_reservation_session(2, i)
                            end if
                        end if
                    next
                end If

        end if  'end of 取得成功預訂課程的日期

    end if

    '回報錯誤
    if (true = is_report_error) then
        Dim objCollectionMyData, bol_err_report, arr_web_info
		Set objCollectionMyData = CreateObject("Scripting.Dictionary")

        objCollectionMyData("Error_Description") = "客戶編號: " & int_client_sn & "<br/>" &_
                            "<font color='red'>錯誤描述: 預訂/取消課程時，" &_
                            "<br/>呼叫/program/member/reservation_class/functions/functions_reservation_class.asp 函式" &_
                            "updatePointAfterReservationOrCancel 時發生錯誤。<br>" &_
                            str_error_msg_detail_desc & "</font>"
        arr_web_info = getSystemEmailAddressInfo(int_source_website)
        if (isarray(arr_web_info)) then
            if (ubound(arr_web_info) >= 2) then
                objCollectionMyData("Website") = arr_web_info(2)
            end if
        end if
        objCollectionMyData("Exec_Source_Code_Info") = str_source_code

        bol_err_report = reportSysErrorWithAction(CONST_ERROR_LEVEL_WARN, str_source_code, _
								  ObjCollectionMyData, CONST_ERROR_REPORT_EMAIL)
        Set ObjCollectionMyData= Nothing
        updatePointAfterReservationOrCancel = false
    else
        updatePointAfterReservationOrCancel = true
    end if   
end function

function getErrorMessageForUpdatePointAfterBooking(ByVal p_arr_reservation_session, ByVal p_int_client_purchase_sn, _
                                 ByVal p_int_client_sn, ByVal p_int_staff_sn, _
                                 ByVal p_str_error_msg)
'/****************************************************************************************
'描述      ：取得 更新到客戶剩餘點數總表 失敗時 要回報的錯誤訊息
'傳入參數  ：[p_arr_reservation_session:Integer] 成功預訂課程的日期和扣堂 (0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn, 
'            [p_int_client_purchase_sn:Integer] client_purchase.account_sn
'            [p_int_client_sn:Integer] 客戶編號 client_basic.sn
'            [p_int_staff_sn:Integer] 員工編號
'            [p_str_error_msg:String] 額外的錯誤訊息
'回傳值    ：成功: 錯誤訊息, 失敗:空字串
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/04/20 Created
'*****************************************************************************************/
    Dim str_report_reservation_msg, str_report_error_msg, i

    str_report_reservation_msg = ""
    str_report_error_msg = ""

    '擷取成功預訂課程的日期和client_attend_list.sn
    if (IsArray(p_arr_reservation_session)) then
        for i = 0 to ubound(p_arr_reservation_session, 2)
            str_report_reservation_msg = str_report_reservation_msg & (i+1) & ". 訂課日期: " & p_arr_reservation_session(0, i)
            str_report_reservation_msg = str_report_reservation_msg & " , 扣堂數: " & p_arr_reservation_session(1, i)
            str_report_reservation_msg = str_report_reservation_msg & " , client_attend_list.sn: " & p_arr_reservation_session(2, i)
            str_report_reservation_msg = str_report_reservation_msg & "<br/>"
        next
    end if

    str_report_error_msg = "客戶編號: " & p_int_client_sn & "<br/>" &_
                           "client_purchase.account: " & p_int_client_purchase_sn & "<br/>" &_
                           "員工編號(若是從IMS訂課): " & p_int_staff_sn & "<br/>" &_
                           "此次預訂課程的資訊如下: <br/>" & str_report_reservation_msg &_
                           "<font color='red'>錯誤描述: " & p_str_error_msg & "</font>"

    getErrorMessageForUpdatePointAfterBooking = str_report_error_msg
end function

function getStaffHandleData(ByVal p_client_sn )
'/****************************************************************************************
'描述      ：查詢負責業務資料
'傳入參數   ： [p_client_sn:int] client_basic.sn
'回傳值    ：回傳業務資料
'牽涉數據表： 
'備註      ：
'歷程(作者/日期/原因) ：[JoyceWang] 2010/07/29 Created
'*****************************************************************************************/
	Dim str_sql : str_sql = ""
	Dim var_arr : var_arr = null
	Dim arr_result : arr_result = null
	
	str_sql = " SELECT fname + ' ' + lname as ename, userid  AS mail "
	str_sql = str_sql & " FROM staff_basic "
	str_sql = str_sql & " WHERE ( "
	str_sql = str_sql & "	sn in ( "
	str_sql = str_sql & "		SELECT in_charge FROM client_record(nolock) "
	str_sql = str_sql & "		WHERE (valid = @valid) AND (client_sn IN ( "
	str_sql = str_sql & "			SELECT lead_sn FROM client_temporal_contract  "
	str_sql = str_sql & "			WHERE (client_sn =@client_sn)"
	str_sql = str_sql & "			))"
	str_sql = str_sql & "		)"
	str_sql = str_sql & "	)"
	
	var_arr = Array(1,p_client_sn)
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)	
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) > 0) then  	'有資料		
			getStaffHandleData = arr_result
		else
			'無資料
			getStaffHandleData = -1
		end if 
	else
		getStaffHandleData = -1
	end if 

end function


function isBonusBegin( Byval p_flt_available_points , ByVal p_str_bonus_begin_datetime)
'/****************************************************************************************
'描述      ：定時定額產品是否需要開通紅利包
'傳入參數  ：[p_flt_available_points:Float] 大錢包 = 已使用堂數(client_attend_list.valid=1) + return - 錄影檔使用
'            [p_str_bonus_begin_datetime:str] 紅利包開通時間
'回傳值    ：布林值: true = 現在開通紅利包 , false = 未開通或者已被開通過
'牽涉數據表：
'備註      ：每次訂課的時候檢查是否要立即開通紅利包，以更新紅利包開通時間
'歷程(作者/日期/原因) ：[Ray Chien] 2010/04/20 Created
'*****************************************************************************************/
	dim flt_available_points : flt_available_points = p_flt_available_points
	dim str_bonus_begin_datetime : str_bonus_begin_datetime = p_str_bonus_begin_datetime
	dim bolBonusBegin : bolBonusBegin = false
	if (flt_available_points <= 0 and isEmptyOrNull(str_bonus_begin_datetime) ) then
		bolBonusBegin = true
	end if
	isBonusBegin = bolBonusBegin
end function





sub checkDrFriendThenSendMail(ByVal p_intClientSn, ByVal p_strEmailUrl, _
                              ByVal p_strCustomerName, ByVal p_strClassMsg, _
                              ByVal p_strWebsite, ByVal p_intConnectType)
'/****************************************************************************************
'描述      ：回傳HTML控制項<select> 可供客戶訂課的時間
'傳入參數  ：[p_intClientSn:Integer] 客戶編號
'            [p_strEmailUrl:String] Email url
'            [p_strCustomerName:String] 客戶名字
'            [p_strClassMsg:String] 新增或取消課程訊息
'            [p_strWebsite:String] 'T=TutorABC, C=VIPABC
'            [p_intConnectType:Integer] 連線的主機型態
'回傳值    ：成功: 回傳HTML控制項<select> 失敗: 空字串
'牽涉數據表： 
'備註      ：博士友人訂課時要發信通知，舊的訂課規則
'歷程(作者/日期/原因) ：[Ray Chien] 2010/08/15 Created
'*****************************************************************************************/
    Dim strSql, arrQueryResult, bolDrVipFriend, intClientSn, strClientSn, strReceiverStaffSn, i, cname, ename
    Dim strEmailUrl : strEmailUrl = p_strEmailUrl
    Dim strCustomerName : strCustomerName = p_strCustomerName
    Dim strClassMsg : strClassMsg = p_strClassMsg
    Dim strWebsite : strWebsite = p_strWebsite
    Dim intConnectType : intConnectType = p_intConnectType


    bolDrVipFriend = false
    intClientSn = p_intClientSn
    strClientSn = cstr(intClientSn)


	strSql = "SELECT category, cname, fname+' '+lname as ename FROM client_basic WHERE (sn = @sn) AND (ISNULL(category, 0) = 1)"
    arrQueryResult = excuteSqlStatementRead(strSql, Array(intClientSn), intConnectType)
    if ( isSelectQuerySuccess(arrQueryResult) ) then
	    if ( Ubound(arrQueryResult) >= 0 ) then
		    bolDrVipFriend = true
            cname = arrQueryResult(1, 0)
            ename = arrQueryResult(2, 0)
	    End If
    End If
		 
    strReceiverStaffSn = ""
    if (true = bolDrVipFriend or _
        strClientSn = "184645" Or strClientSn = "184527" Or strClientSn = "214178" Or strClientSn = "214180") then
        if (strClientSn = "181302" or strClientSn = "181305" or strClientSn = "181306" ) then
            strReceiverStaffSn = "1,2,572,50,465,11,463,537,302,615,289,65,54,387,167,344,265"
        elseif (strClientSn = "184645" or strClientSn = "184527") then
            strReceiverStaffSn = "1,2,725,968,54,228,167,11,572,795,481,265"
        elseif (strClientSn = "193925" or  strClientSn = "193926" or strClientSn = "193927" or  strClientSn = "193928") then
            strReceiverStaffSn = "1,2,54,167,11,572,265"
        elseif (strClientSn = "194880") then
            strReceiverStaffSn = "1,2,54,167,11,265"
        elseif (strClientSn = "214178" or strClientSn = "214180") then
        'CS-John/Joan/Sabrina/Kent
        'PM-Sharon W/Kathy/Tina/Serena/Jack H/Jin Chan 
        '如果是博士友人訂課會寄信給CS 20091223增加VIP兩位訂課要寄MAIL通知特定主管 工單:CS09122106
            strReceiverStaffSn = "11,167,265,463,615,486,302,289,65,462,481"
        else
            strReceiverStaffSn = "167,11,265"
        end if
		'發送信件給設定的主管和經理 
		str_class_email_subject = "VIP博士友人" & cname & " " & ename & "_使用課程預約新增/取消系統"
   



        response.Write "CONST_EMAIL_DR_FRIENDS_RESERVATION_OR_CANCEL_SESSION =  " & CONST_EMAIL_DR_FRIENDS_RESERVATION_OR_CANCEL_SESSION & "<br>"
        response.Write "strClassMsg =  " & strClassMsg & "<br>"
        response.Write "strWebsite =  " & strWebsite & "<br>"
    


		Call sendOrderCancelEmail(strEmailUrl, CONST_EMAIL_DR_FRIENDS_RESERVATION_OR_CANCEL_SESSION, _
                                              strCustomerName, strClassMsg, _
                                              strWebsite)
		
		
        'strSql = "select fname+' '+lname as ename,userid+'@tutorabc.com' as email from staff_basic where sn in (" & strReceiverStaffSn &") and status=1"
'        arrQueryResult = excuteSqlStatementReadQuick(strSql, intConnectType)
'        if ( isSelectQuerySuccess(arrQueryResult) ) then
'	        if ( Ubound(arrQueryResult) >= 0 ) then                
'                for i = 0 to Ubound(arrQueryResult, 2)
'                    '發送信件給設定的主管和經理
'                    str_class_email_subject = "VIP博士友人_使用課程預約新增/取消系統"
'                    Call sendOrderCancelEmail(strEmailUrl, arrQueryResult(1,i), _
'                                              strCustomerName, strClassMsg, _
'                                              strWebsite)
'                next
'	        End If
'        End If
    end if
end sub

function getComboProductSession(ByVal p_intClientSn, ByVal p_intContractSn, ByVal p_intClassType,  ByVal p_intShowMode, ByVal p_intDeBugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述 ： 取得套餐課程類型的剩餘堂數
'傳入參數 ：
'[p_intClientSn:Integer] 客戶編號
'[p_intProductSn:Integer] 產品編號
'[p_intClassType:Integer] 課程類型 7 = 一對一, 3 = 一對多, 99 = 大會堂
'[p_intShowMode:Integer] 顯示模式 1= 次數 2 =實際堂數
'[p_intDeBugMode:Integer] 除錯模式 1= 開啟(會秀出相關Sql) 0 =關閉
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： 套餐課程類型的剩餘堂數
'牽涉數據表： cfgComboProductSession, v_contract_session_points_new
'備註 ：
'歷程(作者/日期/原因) ： [ArTHur Wu] 2012/07/23 Created
'*****************************************************************************************/
    dim strConditionColumn : strConditionColumn = ""
    dim arrParam : arrParam = null
    dim strSql : strSql = ""
    dim objClientProductData : objClientProductData = null
    dim objProductData : objProductData = null
    dim arrSqlResult : arrSqlResult = null
    dim bolDeBugMode : bolDeBugMode = false
    dim intHavingSession : intHavingSession = 0
    dim objProductTypeData : objProductTypeData = null

    '只有1會開啟, 其他預設均不開啟
    if ( Not isEmptyOrNull(p_intDeBugMode) ) then
        if ( 1 = sCInt(p_intDeBugMode) ) then
            bolDeBugMode = true
        end if
    end if
    
    '20121108 阿捨新增 改版優化 改為讀取資料庫
    '先讀取產品服務別
    arrParam = Array(p_intContractSn)
    strSql =  " SELECT cfg_product.web_site, cfg_product.sn FROM cfg_product INNER JOIN client_temporal_contract ON cfg_product.sn = client_temporal_contract.prodbuy "
    strSql =  strSql & " WHERE client_temporal_contract.sn =@sn "
    Set objProductData = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for讀取產品服務別:" & g_str_sql_statement_for_debug & "<br>")
    end if
    if ( not objProductData.eof ) then
        intProductSn = objProductData("sn")
        strWebsite = objProductData("web_site")
    end if
    '再讀取課程類型
    Set objProductTypeData = getProductTypeData(p_intClassType, strWebsite, p_intDeBugMode, p_intConnectType)
    if ( not objProductTypeData.eof ) then
        strSessionName = objProductTypeData("sst_description")
        intSessionTypeSn = objProductTypeData("sn")
    end if
    '帶入已設定好之課程堂數view
    strConditionColumn = " CONVERT(NUMERIC(10, 2), ISNULL(v_contract_session_points_new." & strSessionName & "預約使用堂數, 0.0)) "
    strConditionColumn1 = " CONVERT(INT, ISNULL(v_contract_session_points_new." & strSessionName & "預約使用堂數, 0)) "

    '讀取客戶套餐課程類型的剩餘堂數
    '20121108 阿捨新增 改版優化 改為多筆讀取
    if ( 1 = sCInt(p_intShowMode) ) then
        arrParam = Array(intSessionTypeSn, p_intClientSn, intProductSn, p_intContractSn)
        strSql =  " SELECT  ( CONVERT(NUMERIC(10, 2), ISNULL(cfgComboProductSession_1.MaxSession, 0.0) ) - " & strConditionColumn & " ) AS HavingSession FROM v_contract_session_points_new WITH (NOLOCK) "
        strSql =  strSql & " LEFT JOIN ( SELECT MaxSession, product_sn FROM cfgComboProductSession WITH (NOLOCK) WHERE SessionType=@SessionType AND valid = 1 ) AS cfgComboProductSession_1 "
        strSql =  strSql & " ON v_contract_session_points_new.product_sn = cfgComboProductSession_1.product_sn "
        strSql =  strSql & " WHERE v_contract_session_points_new.client_sn = @client_sn AND v_contract_session_points_new.product_sn = @product_sn AND v_contract_session_points_new.contract_sn = @contract_sn "
    elseif ( 2 = sCInt(p_intShowMode) ) then 'for VIPABC官網訂課頁面
        arrParam = Array(intSessionTypeSn, p_intClientSn, intProductSn, p_intContractSn, p_intClassType, strWebsite)
        strSql =  " SELECT  ( CONVERT(INT, product_session_rule.psr_deduct_session) * ( CONVERT(INT, ISNULL(cfgComboProductSession_1.MaxSession, 0)) - " & strConditionColumn1 & ") ) AS HavingSession FROM v_contract_session_points_new WITH (NOLOCK) "
        strSql =  strSql & " LEFT JOIN ( SELECT MaxSession, product_sn FROM cfgComboProductSession WITH (NOLOCK) WHERE SessionType=@SessionType AND valid = 1 ) AS cfgComboProductSession_1 "
        strSql =  strSql & " ON v_contract_session_points_new.product_sn = cfgComboProductSession_1.product_sn "
        strSql =  strSql & " INNER JOIN product_session_rule ON cfgComboProductSession_1.product_sn = product_session_rule.product_sn "
        strSql =  strSql & " INNER JOIN cfgSessionType ON cfgSessionType.sst_number = product_session_rule.sst_number "
        strSql =  strSql & " WHERE v_contract_session_points_new.client_sn = @client_sn AND v_contract_session_points_new.product_sn = @product_sn AND v_contract_session_points_new.contract_sn = @contract_sn "
        strSql =  strSql & " AND cfgSessionType.sst_number = @sst_number AND cfgSessionType.web_site = @web_site "
    end if
     
    Set objClientProductData = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for 讀取客戶套餐課程類型的剩餘堂數:" & g_str_sql_statement_for_debug & "<br>")
    end if
    if ( not objClientProductData.eof ) then
        intHavingSession = objClientProductData("HavingSession")
    end if

    getComboProductSession = intHavingSession
end function

function getComboProductMaxSession(ByVal p_intProductSn, ByVal p_intClassType, ByVal p_intDeBugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述 ： 取得套餐課程類型的最大堂數
'傳入參數 ：
'[p_intProductSn:Integer] 產品編號
'[p_intClassType:Integer] 課程類型 7 = 一對一, 3 = 一對多, 99 = 大會堂
'[p_intDeBugMode:Integer] 除錯模式 1= 開啟(會秀出相關Sql) 0 =關閉
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： 套餐課程類型的最大堂數
'牽涉數據表： cfgComboProductSession, v_contract_session_points_new_new
'備註 ：
'歷程(作者/日期/原因) ： [ArTHur Wu] 2012/07/23 Created
'*****************************************************************************************/
    dim strColumnName : strColumnName = ""
    dim arrParam : arrParam = null
    dim strSql : strSql = ""
    dim objClientProductData : objClientProductData = null
    dim bolDeBugMode : bolDeBugMode = false
    dim intMaxSession : intMaxSession = 0
    dim objProductData : objProductData = null
    dim objProductTypeData : objProductTypeData = null

    '只有1會開啟, 其他預設均不開啟
    if ( Not isEmptyOrNull(p_intDeBugMode) ) then
        if ( 1 = sCInt(p_intDeBugMode) ) then
            bolDeBugMode = true
        end if
    end if

    '20121108 阿捨新增 改版優化 改為讀取資料庫
    '先讀取產品服務別
    arrParam = Array(p_intProductSn)
    strSql =  " SELECT web_site FROM cfg_product WHERE sn =@sn "
    Set objProductData = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for讀取產品服務別:" & g_str_sql_statement_for_debug & "<br>")
    end if
    if ( not objProductData.eof ) then
        strWebsite = objProductData("web_site")
    end if
    '再讀取課程類型
    Set objProductTypeData = getProductTypeData(p_intClassType, strWebsite, p_intDeBugMode, p_intConnectType)
    if ( not objProductTypeData.eof ) then
        intSessionTypeSn = objProductTypeData("sn")
    end if

    '讀取套餐課程類型的最大堂數
    '20121108 阿捨新增 改版優化 改為多筆讀取
    arrParam = Array(p_intProductSn, intSessionTypeSn)
    strSql =  " SELECT CONVERT(NUMERIC(10, 2), ISNULL(MaxSession, 0.0)) AS MaxSession FROM cfgComboProductSession WITH (NOLOCK) "
    strSql = strSql & " WHERE product_sn = @product_sn AND SessionType=@SessionType AND valid = 1 "

    Set objClientProductData = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for 讀取套餐課程類型的最大堂數:" & g_str_sql_statement_for_debug & "<br>")
    end if
    if ( not objClientProductData.eof ) then
        intMaxSession = objClientProductData("MaxSession")
    end if

    getComboProductMaxSession = intMaxSession
end function

function getProductTypeData(ByVal p_intClassType, ByVal p_strWebSite, ByVal p_intDeBugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述 ： 取得課程類型的相關資料
'傳入參數 ：
'[p_intClassType:Integer] 課程類型 7 = 一對一, 3 = 一對多, 99 = 大會堂
'[p_strWebSite:String] 服務別區分 T: TutorABC C:VIPABC
'[p_intDeBugMode:Integer] 除錯模式 1= 開啟(會秀出相關Sql) 0 =關閉
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： 課程類型的相關資訊
'牽涉數據表： cfgSessionType
'備註 ：
'歷程(作者/日期/原因) ： [ArTHur Wu] 2012/11/08 Created
'*****************************************************************************************/
    dim arrParam : arrParam = null
    dim strSql : strSql = ""
    dim objProductTypeData : objProductTypeData = null
    dim bolDeBugMode : bolDeBugMode = false

    if ( Not isEmptyOrNull(p_intDeBugMode) ) then
        if ( 1 = sCInt(p_intDeBugMode) ) then
            bolDeBugMode = true
        end if
    end if

    '讀取課程類型的相關資訊
    
    strSql =  " SELECT sn, sst_number, sst_description, sst_description_cn, sst_description_en FROM dbo.cfgSessionType WHERE valid = 1 AND web_site = @web_site "
    '給空值可以查詢全部服務別的課程類型
    if ( Not isEmptyOrNull(p_intClassType) ) then
        arrParam = Array(p_strWebSite, p_intClassType)
        strSql = strSql & " AND sst_number = @sst_number "
    else
        arrParam = Array(p_strWebSite)
    end if
    Set objProductTypeData = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for 讀取課程類型的相關資訊:" & g_str_sql_statement_for_debug & "<br>")
    end if
    'if ( not objProductTypeData.eof ) then
        Set getProductTypeData = objProductTypeData
    'end if
end function

function isComboProduct(ByVal p_intProductSn, ByVal p_intDeBugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述 ： 判斷產品是否為套餐產品
'傳入參數 ：
'[p_intProductSn:Integer] 產品編號
'[p_intDeBugMode:Integer] 除錯模式 1= 開啟(會秀出相關Sql) 0 =關閉
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： true or false
'牽涉數據表： cfgComboProductSession,
'備註 ：
'歷程(作者/日期/原因) ： [ArTHur Wu] 2012/11/12 Created
'*****************************************************************************************/
    dim arrParam : arrParam = null
    dim strSql : strSql = ""
    dim objClientProductData : objClientProductData = null
    dim bolDeBugMode : bolDeBugMode = false

    '只有1會開啟, 其他預設均不開啟
    if ( Not isEmptyOrNull(p_intDeBugMode) ) then
        if ( 1 = sCInt(p_intDeBugMode) ) then
            bolDeBugMode = true
        end if
    end if

    arrParam = Array(p_intProductSn)
    strSql =  " SELECT sn FROM cfgComboProductSession WITH (NOLOCK) "
    strSql = strSql & " WHERE product_sn = @product_sn AND CONVERT(int, ISNULL(MaxSession, 0)) > 0 AND valid = 1 "
    
    Set objClientProductData = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for 讀取是否符合套餐產品:" & g_str_sql_statement_for_debug & "<br>")
    end if
    if ( not objClientProductData.eof ) then
        isComboProduct = true
    else
        isComboProduct = false
    end if

end function

function getProductSessionType(ByVal p_intProductSn, ByVal p_intDeBugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述 ： 取得產品可上課類型 
'傳入參數 ：
'[p_intProductSn:Integer] 產品編號
'[p_intDeBugMode:Integer] 除錯模式 1= 開啟(會秀出相關Sql) 0 =關閉
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： data role
'牽涉數據表： cfgSessionType, product_session_rule
'備註 ：
'歷程(作者/日期/原因) ： [ArTHur Wu] 2012/12/05 Created
'*****************************************************************************************/
    dim arrParam : arrParam = null
    dim strSql : strSql = ""
    dim objProductSessionData : objProductSessionData = null
    dim bolDeBugMode : bolDeBugMode = false

    '只有1會開啟, 其他預設均不開啟
    if ( Not isEmptyOrNull(p_intDeBugMode) ) then
        if ( 1 = sCInt(p_intDeBugMode) ) then
            bolDeBugMode = true
        end if
    end if

    arrParam = Array(p_intProductSn)
    strSql =  " SELECT product_session_rule.sst_number, cfgSessionType.sst_description_en "
    strSql = strSql & " FROM product_session_rule WITH (NOLOCK) INNER JOIN cfgSessionType "
    strSql = strSql & " ON product_session_rule.sst_number = cfgSessionType.sst_number "
    strSql = strSql & " WHERE product_session_rule.product_sn = @product_sn AND cfgSessionType.valid = 1 AND web_site = 'C' "
    
    Set objProductSessionData = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDeBugMode ) then
        Response.Write("錯誤訊息Sql for 取得產品可上課類型:" & g_str_sql_statement_for_debug & "<br>")
    end if
    
    Set getProductSessionType = objProductSessionData
end function

function setComboProductGift(ByVal p_intContractSn, ByVal p_objSessionTypeGift, ByVal p_strNote, ByVal p_intStaffSn, ByVal p_intUpdateMode, ByVal p_intDebugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述：新增套餐產品每種上課類型之贈送堂數
'傳入參數：
'[p_intContractSn:Integer] 合約編號
'[p_objSessionTypeGift]  : 課程類型編號(cfgSessionType.sn)與贈送堂數
'[p_strNote] 贈送原因
'[p_intStaffSn] 設定者(員工編號)
'[p_intUpdateMode] 更新模式 0: 關閉(一定要新增, 不用判斷資料存在) 1:開啟(若有資料就更新, 若無資料就新增)
'[p_intDebugMode] 除錯模式 0: 關閉 1:開啟
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： ture or false
'牽涉數據表：
'備註 ：
'歷程(作者/日期/原因) ：[ArTHurWu] 2013/01/03 Created
'*****************************************************************************************/
    dim arrParam : arrParam = null
    dim bolSucc : bolSucc = true
    dim strSql : strSql = ""
    dim arrSqlResult : arrSqlResult = null
    dim bolDebugMode : bolDebugMode = false
    dim bolUpdateMode : bolUpdateMode = false
    dim intTotalNum : intTotalNum = 0
    dim intDataGiftSession : intDataGiftSession = 0
    dim bolNewData : bolNewData = false
    dim strConditionSql : strConditionSql = ""

    if ( 1 = sCInt(p_intDebugMode) ) then
        bolDebugMode = true
    end if

    if ( 1 = sCInt(p_intUpdateMode) ) then
        bolUpdateMode = true
    end if

    intTotalNum = p_objSessionTypeGift.Count
    if ( true = bolDebugMode ) then
        response.Write "intTotalNum: " & intTotalNum & "<br />"
    end if   

    '字典檔存在時
    if ( 0 < intTotalNum ) then
        dim intSessionTypeSn : intSessionTypeSn = "" '字典檔的key 也是課程類型編號
        dim intGiftSession : intGiftSession = 0
        for each intSessionTypeSn in p_objSessionTypeGift.Keys
            intGiftSession = p_objSessionTypeGift.Item(intSessionTypeSn) '字典檔的value 也是課程類型對應到的贈送點數
            intGiftSession = sCInt(intGiftSession) * CONST_SESSION_TRANSFORM_POINT
            if ( true = bolDebugMode ) then
                response.Write "intSessionTypeSn: " & intSessionTypeSn & "<br />"
                response.Write "intGiftSession: " & intGiftSession & "<br />"
            end if   
             '若開啟更新模式 要判斷資料是否存在
             if ( true = bolUpdateMode ) then
                intDataGiftSession = getComboProductGift(p_intContractSn, "", "", intSessionTypeSn, p_intDebugMode, p_intConnectType)
                 if ( true = bolDebugMode ) then
                    response.Write "intDataGiftSession: " & intDataGiftSession & "<br />"
                end if 
                if ( 0 = sCInt(intDataGiftSession) ) then
                    bolNewData = true  '沒有資料
                else
                    bolNewData = false
                end if
            end if

            if ( true = bolDebugMode ) then
                response.Write "bolNewData: " & bolNewData & "<br />"
            end if 

            arrParam = Array(intGiftSession, p_strNote, p_intStaffSn, intSessionTypeSn, p_intContractSn)
            if ( (true = bolUpdateMode AND true = bolNewData) OR false = bolUpdateMode ) then '開啟更新模式且沒有資料就新增 or 關閉更新模式(一定要新增, 不用判斷資料存在)
                if ( 0 < sCInt(intGiftSession) ) then '若0不新增
                    strSql = " INSERT INTO ContractChangeRecord (ChangeTypeSn, GiftPoint, Note, StaffSn, InputeDate, Valid, SessionTypeSn, ContractSn) VALUES "
                    strSql = strSql & " ( '3', @GiftPoint, @Note, @StaffSn, GETDATE(), 1, @SessionTypeSn, @ContractSn ) "
                    'strSql = strSql & " SELECT SCOPE_IDENTITY() "

                    arrSqlResult = excuteSqlStatementScalar(strSql, arrParam, p_intConnectType)
                    if ( true = bolDebugMode ) then
                        Response.Write("錯誤訊息sql for 新增套餐產品每種上課類型之贈送堂數:" & g_str_sql_statement_for_debug & "<br>")
                    end if
                    if ( isSelectQuerySuccess(arrSqlResult) ) then
                        if ( Ubound(arrSqlResult) >= 0 ) then
                            'intNewRecordSn = arrSqlResult(0)
                        end if
                    else
                        bolSucc = false
                    end if
                end if
            else
                '若贈送為0 直接將log更新為失效
                if ( 0 = sCInt(intGiftSession) ) then
                    strConditionSql = " ,ContractChangeRecord.Valid = '0' "
                else
                    strConditionSql = ""
                end if
                '有資料就update
                strSql = " UPDATE ContractChangeRecord SET "
                strSql = strSql & " ContractChangeRecord.ChangeTypeSn = '3', ContractChangeRecord.GiftPoint = @GiftPoint, "
                strSql = strSql & " ContractChangeRecord.Note = @Note, ContractChangeRecord.StaffSn = @StaffSn, "
                strSql = strSql & " ContractChangeRecord.InputeDate = GETDATE() "
                strSql = strSql & strConditionSql
                strSql = strSql & " WHERE ContractChangeRecord.SessionTypeSn = @SessionTypeSn AND ContractChangeRecord.ContractSn = @ContractSn AND ContractChangeRecord.valid = 1 "

                arrSqlResult = excuteSqlStatementWrite(strSql, arrParam, p_intConnectType)
                if ( true = bolDebugMode ) then
                    Response.Write("錯誤訊息Sql for 更新套餐產品每種上課類型之贈送堂數:" & g_str_sql_statement_for_debug & "<br>")
                end if
                if ( arrSqlResult >= 0 ) then
                else
                    bolSucc = false
                end if
            end if
        next
    end if

    setComboProductGift = bolSucc
end Function

function getComboProductGift(ByVal p_intContractSn, ByVal p_intAccountSn, ByVal p_intSessionType, ByVal p_intSessionTypeSn, ByVal p_intDebugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述：讀出課程類型之贈送堂數
'傳入參數：
'[p_intContractSn:Integer] 開通前合約編號
'[p_intAccountSn:Integer] 開通後合約流水號
'[p_intSessionType] 課程類型 可為空, 
'[p_intSessionTypeSn] 課程類型編號 可為空, 
'注意 兩者為空 秀出全部 但keys 為 p_intSessionType (cfgSessionType.sst_number)
'[p_intDebugMode] 除錯模式 0: 關閉 1:開啟
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： 單一課程類型之贈送堂數 or 全部類型之贈送堂數(Dictionary)
'牽涉數據表：
'備註 ：
'歷程(作者/日期/原因) ：[ArTHurWu] 2013/01/03 Created
'*****************************************************************************************/
    dim arrParam : arrParam = null
    dim bolSucc : bolSucc = true
    dim strSql : strSql = ""
    dim objContractRecord : objContractRecord = null
    dim bolDebugMode : bolDebugMode = false
    dim intTotalNum : intTotalNum = 0
    dim objSessionTypeGift : objSessionTypeGift = null

    if ( 1 = sCInt(p_intDebugMode) ) then
        bolDebugMode = true
    end if

    if ( Not isEmptyOrNull(p_intSessionType) ) then
        arrParam = Array(p_intContractSn, p_intSessionType)
        strSql = " SELECT SUM(CONVERT(NUMERIC(8, 0), CEILING(ROUND((ISNULL(ContractChangeRecord.GiftPoint, 0.00) + 0.00) / 65, 4) * 100) / 100)) AS GiftPoint FROM ContractChangeRecord WITH (NOLOCK) "
        strSql = strSql & " INNER JOIN cfgSessionType WITH (NOLOCK) ON ContractChangeRecord.SessionTypeSn = cfgSessionType.sn "
        strSql = strSql & " WHERE ContractChangeRecord.ContractSn = @ContractSn AND cfgSessionType.sst_number = @sst_number AND ContractChangeRecord.Valid = '1' "
        strSql = strSql & " AND cfgSessionType.web_site = 'C' "
        Set objContractRecord = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
        if ( true = bolDebugMode ) then
            Response.Write("錯誤訊息sql for 讀出課程類型之贈送堂數:" & g_str_sql_statement_for_debug & "<br>")
        end if
        if ( not objContractRecord.eof ) then
            getComboProductGift = objContractRecord("GiftPoint")
        else
            getComboProductGift = 0
        end if
    elseif ( Not isEmptyOrNull(p_intSessionTypeSn) ) then
        arrParam = Array(p_intContractSn, p_intSessionTypeSn)
        strSql = " SELECT SUM(CONVERT(NUMERIC(8, 0), CEILING(ROUND((ISNULL(ContractChangeRecord.GiftPoint, 0.00) + 0.00) / 65, 4) * 100) / 100)) AS GiftPoint FROM ContractChangeRecord WITH (NOLOCK) "
        strSql = strSql & " WHERE ContractChangeRecord.ContractSn = @ContractSn AND ContractChangeRecord.SessionTypeSn = @SessionTypeSn AND ContractChangeRecord.Valid = '1' "
        Set objContractRecord = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
        if ( true = bolDebugMode ) then
            Response.Write("錯誤訊息sql for 讀出課程類型之贈送堂數:" & g_str_sql_statement_for_debug & "<br>")
        end if
        if ( not objContractRecord.eof ) then
            getComboProductGift = objContractRecord("GiftPoint")
        else
            getComboProductGift = 0
        end if
    else
        arrParam = Array(p_intContractSn, p_intAccountSn)
        strSql = " SELECT SUM(CONVERT(NUMERIC(8, 0), CEILING(ROUND((ISNULL(ContractChangeRecord.GiftPoint, 0.00) + 0.00) / 65, 4) * 100) / 100)) AS GiftPoint, cfgSessionType.sst_number FROM ContractChangeRecord WITH (NOLOCK) "
        strSql = strSql & " INNER JOIN cfgSessionType WITH (NOLOCK) ON ContractChangeRecord.SessionTypeSn = cfgSessionType.sn "
        strSql = strSql & " WHERE ( ContractChangeRecord.ContractSn = @ContractSn OR ContractChangeRecord.AccountSn = @AccountSn ) AND ContractChangeRecord.SessionTypeSn >'' AND ContractChangeRecord.Valid = '1' "
        strSql = strSql & " GROUP BY GiftPoint, sst_number "
        Set objContractRecord = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
        if ( true = bolDebugMode ) then
            Response.Write("錯誤訊息sql for 讀出所有類型贈送堂數:" & g_str_sql_statement_for_debug & "<br>")
        end if
        Set objSessionTypeGift = CreateObject("Scripting.Dictionary")
        if ( not objContractRecord.eof ) then
            while not objContractRecord.eof
                intSessionType = objContractRecord("sst_number")
                intGiftSession = objContractRecord("GiftPoint")
                if ( true = bolDebugMode ) then
                    response.Write "intSessionType: " & intSessionType & "<br />"
                    response.Write "intGiftSession: " & intGiftSession & "<br />"
                end if  
                objSessionTypeGift(intSessionType) = intGiftSession
            objContractRecord.movenext
            wend
            set getComboProductGift = objSessionTypeGift
        else
            set getComboProductGift =objSessionTypeGift
        end if
            
    end if
end Function

function getComboAllGiftPoints(ByVal p_intContractSn, ByVal p_intAccountSn, ByVal p_intDebugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述：取出所有套餐贈送總點數 for 開通時寫入合約client_purchase.access_points
'傳入參數：
'[p_intContractSn:Integer] 開通前合約編號
'[p_intAccountSn:Integer] 開通後合約流水號
'[p_intDebugMode] 除錯模式 0: 關閉 1:開啟
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ： ture or false
'牽涉數據表：
'備註 ：
'歷程(作者/日期/原因) ：[ArTHurWu] 2013/01/09 Created
'*****************************************************************************************/
    dim objProductTypeData : objProductTypeData = null
    dim objSessionTypeGift : objSessionTypeGift = null
    dim objProductSessionType : objProductSessionType = null
    dim intSessionType : intSessionType = 0
    dim intGiftSession : intGiftSession = 0
    dim intProductSession : intProductSession = 0
    dim intRealSession : intRealSession = 0

    if ( 1 = sCInt(p_intDebugMode) ) then
        bolDebugMode = true
    end if
    
    '讀取贈送課程類型點數 空值表讀出全部
    Set objSessionTypeGift = getComboProductGift(p_intContractSn, p_intAccountSn, "", "", p_intDebugMode, p_intConnectType)

    '產品課程扣堂堂數 空值表讀出全部
    Set objProductSessionType = getSessionTypeRealPoints(p_intContractSn, "", p_intDebugMode, p_intConnectType)

    '課程類型 空值表讀出全部
    Set objProductTypeData = getProductTypeData("", "C", p_intDebugMode, p_intConnectType)
    while not objProductTypeData.eof
        intSessionType = objProductTypeData("sst_number")
        intGiftSession = objSessionTypeGift(intSessionType)
        intProductSession = objProductSessionType(intSessionType)
        if ( true = bolDebugMode ) then
            response.Write "intSessionType: " & intSessionType & "<br />"
            response.Write "intGiftSession: " & intGiftSession & "<br />"
            response.Write "intProductSession: " & intProductSession & "<br />"
        end if  
        intRealSession = sCDbl(intGiftSession) * sCDbl(intProductSession)
        intTotalSession = intTotalSession + intRealSession
        if ( true = bolDebugMode ) then
            response.Write "intRealSession: " & intRealSession & "<br />"
            response.Write "intTotalSession: " & intTotalSession & "<br />"
        end if  
    objProductTypeData.movenext
    wend

    intTotalSession = intTotalSession * CONST_SESSION_TRANSFORM_POINT
    getComboAllGiftPoints = intTotalSession
end Function

function getSessionTypeRealPoints(ByVal p_intContractSn, ByVal p_intSessionType, ByVal p_intDebugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述：取出課程類型實際扣除點數
'傳入參數：
'[p_intContractSn:Integer] 開通前合約編號
'[p_intSessionType:Integer] 課程類型編號 若給空值, 秀出全部
'[p_intDebugMode] 除錯模式 0: 關閉 1:開啟
'[p_intConnectType:Integer] 連線的主機型態
'回傳值 ：堂數 或 整個物件
'牽涉數據表：
'備註 ：
'歷程(作者/日期/原因) ：[ArTHurWu] 2013/01/09 Created
'*****************************************************************************************/
    dim arrParam : arrParam = null
    dim strSql : strSql = ""
    dim strConditionSql : strConditionSql = ""
    dim objProductSessionType : objProductSessionType = null
    dim objProductSession : objProductSession = null
    dim bolDebugMode : bolDebugMode = false
    dim intSessionType : intSessionType = 0
    dim intProductSession : intProductSession = 0

    if ( 1 = sCInt(p_intDebugMode) ) then
        bolDebugMode = true
    end if

    if ( Not isEmptyOrNull(p_intSessionType) ) then
        arrParam = Array(p_intContractSn, p_intSessionType)
        strConditionSql = " AND product_session_rule.sst_number = @sst_number "
    else
        arrParam = Array(p_intContractSn)
    end if

    strSql = " SELECT product_session_rule.psr_deduct_session,  product_session_rule.sst_number FROM product_session_rule WITH (NOLOCK) "
    strSql = strSql & " INNER JOIN client_temporal_contract WITH (NOLOCK) ON product_session_rule.product_sn = client_temporal_contract.prodbuy "
    strSql = strSql & " WHERE client_temporal_contract.sn = @sn "
    strSql = strSql & strConditionSql & " ORDER BY product_session_rule.sst_number "
    Set objProductSessionType = excuteSqlStatementReadEach(strSql, arrParam, p_intConnectType)
    if ( true = bolDebugMode ) then
        Response.Write("錯誤訊息sql for 取出課程類型實際扣除點數:" & g_str_sql_statement_for_debug & "<br>")
    end if

     if ( Not isEmptyOrNull(p_intSessionType) ) then
         if ( Not objProductSessionType.eof ) then
            getSessionTypeRealPoints = objProductSessionType("psr_deduct_session")
        else
            getSessionTypeRealPoints = 0
        end if
    else
        Set objProductSession = CreateObject("Scripting.Dictionary")
        if ( not objProductSessionType.eof ) then
            while not objProductSessionType.eof
                intSessionType = objProductSessionType("sst_number")
                intProductSession = objProductSessionType("psr_deduct_session")
                if ( true = bolDebugMode ) then
                    response.Write "intSessionType: " & intSessionType & "<br />"
                    response.Write "intProductSession: " & intProductSession & "<br />"
                end if  
                objProductSession(intSessionType) = intProductSession
            objProductSessionType.movenext
            wend
            set getSessionTypeRealPoints = objProductSession
        else
            set getSessionTypeRealPoints = objProductSession
        end if
    end if

end Function
%>