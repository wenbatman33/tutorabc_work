<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"--><!--TREE VIP系統信-->
<!--#include virtual="/lib/functions/functions_ABCSystemSendMail/functions_ABCSystemSendMail.asp"--><%'ABC系統信%>
<% dim bolDebugMode : bolDebugMode = false %>
<!--#include virtual="/program/member/reservation_class/include/getComboProductData.asp"-->
<%
'會有六個session 五個變數 for 套餐判斷
'bolComboProduct, bolComboHaving1On1, bolComboHaving1On3, bolComboHavingLobby, bolComboHavingRecord
'20121206 阿捨新增 一對四判斷 bolComboHaving1On4
if ( true = bolDebugMode ) then
    response.Write "bolComboProduct: " & bolComboProduct & "<br />"
    response.Write "bolComboHaving1On1: " & bolComboHaving1On1 & "<br />"
    response.Write "bolComboHaving1On3: " & bolComboHaving1On3 & "<br />"
    response.Write "bolComboHaving1On4: " & bolComboHaving1On4 & "<br />"
    response.Write "bolComboHavingLobby: " & bolComboHavingLobby & "<br />"
    response.Write "bolComboHavingRecord: " & bolComboHavingRecord & "<br />"
end if

dim bolTestCase : bolTestCase = false '測試模式

'20120926 阿捨新增 特殊關懷訊息
dim strOrderClassInfo : strOrderClassInfo = ""
dim strCancelClassInfo : strCancelClassInfo = ""
dim strClassType : strClassType = ""

    'TODO: 修改取得剩餘堂數

    '預訂&重訂課程成功頁面 20091208 raychien

    '停用頁面的快取暫存
    Call setPageNoCache()
    
    '**** 參數設定 START ****
    Dim int_db_connect : int_db_connect = CONST_VIPABC_RW_CONN                          'DB Connection Code
    Dim int_website_code : int_website_code = CONST_VIPABC_WEBSITE                      'Website Code
    Dim str_class_minute : str_class_minute = "30"
    '**** 參數設定 END   ****

    '20120316 阿捨新增 無限卡產品判斷	
    Dim bolUnlimited : bolUnlimited = false
    if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
        bolUnlimited = true 
    end if


    '**** 成功和錯誤訊息設定 START ****
    Dim str_arr_success_msg, str_arr_fail_msg
    '[0] = "訂位失敗，請與客服人員連絡：4006-30-30-30"
    '[1] = '"合约剩余课时数不足，请至「学习记录」查询详细资讯。"
    '[2] = "取消課程失敗，請與客服人員連絡：4006-30-30-30"
    '[3] = "取消失敗(上課前$hour$小時內無法取消課程)"
    '[4] = "取消失敗(已經超過上課時間的課程無法取消)"
    '[5] = "取消失敗(此堂課您並沒有預定)"
    '[6] = "預定失敗"
    '[7] = "預定失敗(已經超過上課時間的課程無法預訂)"
    '[8] = 預定失敗(請預訂$contract_service_start$之後的課程)
    '[9] = 預定失敗(請預訂$contract_service_end$之前的課程)
    '[10] = 四小时前没有开课，请选择其他上课日期或咨询时段
    '[11] = 24小时内临时订课仅开放一对多，请选择其他上课日期或咨询时段
    '[12]' = 大会堂只允许开课前五分钟可以订课，请选择其他上课日期或咨询时段
    '[13] = 預定失敗(合约剩余课时数不足)
    '[14] = 預定失敗(此堂课您已经预定过了)
    '[15] = 重新预定失败
    '[16] = 请选择预约或取消的课程时段
    str_arr_fail_msg = Array("订位失败，请与客服人员连络：4006-30-30-30\n", _
                            "合约剩余课时数不足，请至「学习记录」查询详细资讯。\n", _
                            "取消课程失败，请与客服人员连络：4006-30-30-30\n", _
                            "取消失败(上课前$hour$小时内无法取消课程)\n", _
                            "取消失败(已经超过上课时间的课程无法取消)\n", _
                            "取消失败(此堂课您并没有预定)\n", _
                            "预定失败\n", _
                            "预定失败(已经超过上课时间的课程无法预订)\n", _
                            "预定失败(请预订$contract_service_start$之後的课程)\n", _
                            "预定失败(请预订$contract_service_end$之前的课程)\n", _
                            "四小时前没有开课，请选择其他上课日期或咨询时段\n", _
                            "24小时内临时订课仅开放一对多，请选择其他上课日期或咨询时段\n", _
                            "大会堂只允许开课前五分钟可以订课，请选择其他上课日期或咨询时段\n", _
                            "预定失败(合约剩余课时数不足)\n", _
                            "预定失败(此堂课您已经预定过了)\n", _
                            "重新预定失败\n", _
                            "请选择预约或取消的课程时段")

    '[0] = 大會堂
    '[1] = 小班(1-3)人
    '[2] = 1對1
    '[3] = 1對6
    str_arr_success_msg = Array("大会堂课程\n", _
                                "小班制课程\n", _
                                "1对1课程\n", _
                                "1对6\n")

    '**** 成功和錯誤訊息設定 END   ****
    


    Dim str_total_reservation_class_dt
    str_total_reservation_class_dt = Trim(Request("hdn_total_reservation_class_dt"))    '客戶訂課資訊
    Dim str_total_cancel_class_sn
    str_total_cancel_class_sn = getRequest("total_cancel_class_sn", CONST_DECODE_NO)    '客戶欲取消的課程資訊 client_attend_list.sn
    Dim str_total_cancel_class_date
    str_total_cancel_class_date = Trim(Request("total_cancel_class_date"))              '客戶欲取消的課程資訊 client_attend_list.attend_date
    Dim str_total_cancel_class_time
    str_total_cancel_class_time = Trim(Request("total_cancel_class_time"))              '客戶欲取消的課程資訊 client_attend_list.attend_sestime
    Dim str_total_cancel_class_type
    str_total_cancel_class_type = Trim(Request("total_cancel_class_type"))              '客戶欲取消的課程資訊 課程類型
    Dim str_order_type : str_order_type = Trim(Request("hdn_order_or_reorder_type"))    '1:新預定的  2:重訂課程  3:取消課程
    Dim flt_subtract_session                                                            '要扣的堂數
    Dim bol_enabled_set_class                                                           'true: 通過檢查可以排課  false: 不能被排課
    Dim str_order_success_msg : str_order_success_msg = ""                              '訂課成功訊息
    Dim str_cancel_success_msg : str_cancel_success_msg = ""                            '取消課程成功訊息
    Dim str_order_fail_msg : str_order_fail_msg = ""                                    '訂課失敗訊息
    Dim str_cancel_fail_msg : str_cancel_fail_msg = ""                                  '取消課程失敗訊息
    Dim flt_total_subtract_session                                                      '總共要扣的堂數
    Dim flt_total_cancel_return_session                                                 '總共要取消後歸還的堂數
    Dim int_total_success_order_count                                                   '總共成功預定的課程筆數
    Dim int_total_fail_order_count                                                      '總共失敗預定的課程筆數
    Dim int_total_success_cancel_count                                                  '總共成功取消的課程筆數
    Dim int_total_fail_cancel_count                                                     '總共失敗取消的課程筆數
    Dim str_arr_total_cancel_class_sn                                                   '欲取消的課程的client_attend_list.sn，以逗點隔開多筆 
    Dim str_arr_total_cancel_class_date                                                 '欲取消的課程的client_attend_list.attend_date，以逗點隔開多筆 
    Dim str_arr_total_cancel_class_time                                                 '欲取消的課程的client_attend_list.attend_sestime，以逗點隔開多筆 
    Dim str_arr_total_cancel_class_type                                                 '欲取消的課程的類型，以逗點隔開多筆 
    Dim str_arr_total_reservation_class                                                 '已預訂的課程，欲取消的課程，以逗點隔開多筆 
                                                                                        '一對一和一對多格式: class_type + date + time 例:072009120820 代表一對一12/8號晚上八點的課
                                                                                        '大會堂格式: class_type + date + time + lobby_session.sn
    
    Dim str_redirect_url : str_redirect_url = ""                                        '執行完後要轉址的頁面
    Dim str_order_success_msg_for_email : str_order_success_msg_for_email = ""          '訂課成功後 寄送給客戶的新增課程訊息
    Dim str_cancel_success_msg_for_email : str_cancel_success_msg_for_email = ""        '取消課程成功後 寄送給客戶的新增課程訊息
    Dim str_class_type_description : str_class_type_description = ""                    '課程類型描述 (使用在 getClassSubtractSession 函式)
    Dim str_subtract_session_error_msg : str_subtract_session_error_msg = ""            '扣堂數錯誤的訊息
    Dim arr_success_reservation_class : arr_success_reservation_class = null            '成功預訂課程的日期和扣堂 (0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn
    Dim arr_success_cancel_class : arr_success_cancel_class = null                      '成功取消課程的日期和扣堂 (0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn    
    Dim int_week_use_session_count : int_week_use_session_count = 0                     '本週已使用堂數(排除return)
    Dim dat_attend_date, int_attend_sestime, str_reservation_class_datetime, arr_client_attend_list, arr_session_cycle_info
    Dim arr_reservation_result, str_class_datetime, i, int_client_sn, int_attend_type_tmp, arr_tmp, intCancelClassType, int_special_sn

    int_total_success_order_count = 0
    int_total_fail_order_count = 0
    int_total_success_cancel_count = 0
    int_total_fail_cancel_count = 0
    flt_total_subtract_session = 0.0
    flt_total_cancel_return_session = 0.0
	
	Dim special1on1value : special1on1value = request("special1on1value")
	
    '******************* 前置作業檢查 (檢查是否符合資格可以訂課) START *******************
    '取得客戶編號
    int_client_sn = getSession("client_sn", CONST_DECODE_NO)

    '設定要轉址到訂課結果頁面
    '設定轉址的頁面
    '課程類型
    Dim strChooseClassType : strChooseClassType = getRequest("clst", CONST_DECODE_NO)
    str_redirect_url = "reservation_class.asp?clst=" & strChooseClassType
    'if (str_order_type = "1") then
        ''設定要轉址到訂課結果頁面
        'str_redirect_url = "/program/member/reservation_class/order_success.asp"    
    'elseif (str_order_type = "2") then
        ''設定要轉址到重新訂課結果頁面 
        'str_redirect_url = "/program/member/reservation_class/reorder_success.asp"
    'end if

    '客戶沒有 有效合約
    if (Not bol_client_have_valid_contract) then
        if (isEmptyOrNull(str_debug_mode)) then
            '"訂位失敗，請與客服人員連絡：02-3365-9999"
            Call alertGo(str_arr_fail_msg(0), "/index.asp", CONST_ALERT_THEN_REDIRECT)
        else
            Call alertGo(str_arr_fail_msg(0) & "(客戶沒有 有效合約)", "/index.asp", CONST_ALERT_THEN_REDIRECT)
        end if
        Response.End
    end if

    '沒有客戶預訂課程的資訊?
    if (isEmptyOrNull(str_total_reservation_class_dt) And isEmptyOrNull(str_total_cancel_class_sn)) then
        if (isEmptyOrNull(str_debug_mode)) then
            '"訂位失敗，請與客服人員連絡：02-3365-9999"
            Call alertGo(str_arr_fail_msg(16), str_redirect_url, CONST_ALERT_THEN_REDIRECT)
        else
            Call alertGo(str_arr_fail_msg(16) & "(沒有客戶預訂(重訂)和取消課程的資訊)", str_redirect_url, CONST_ALERT_THEN_REDIRECT)
        end if
        Response.End
    end if
    
    '******************* 前置作業檢查 (檢查是否符合資格可以訂課) END   *******************


    '******************* 取消課程 START *******************
    if (Not isEmptyOrNull(str_total_cancel_class_sn)) then

        Dim arr_allow_cancel, arr_cancel, arr_chk_cancel_param
        '取消的課程 client_attend_list.sn
        str_arr_total_cancel_class_sn = Split(str_total_cancel_class_sn, ",")
        '取消的課程 client_attend_list.attend_date
        str_arr_total_cancel_class_date = Split(str_total_cancel_class_date, ",")
        '取消的課程 client_attend_list.attend_ststime
        str_arr_total_cancel_class_time = Split(str_total_cancel_class_time, ",")
        '取消的課程 課程類型
        str_arr_total_cancel_class_type = Split(str_total_cancel_class_type, ",")


        '**** 取得課程在結算週期內的資訊，只有大陸定時定額產品才需要判斷 START ****
        if (CONST_PRODUCT_QUOTA = int_client_purchase_product_type) then
            arr_session_cycle_info = getSessionCycleDateInfo(Array(int_account_cycle_day, str_service_start_date), now)
            if (IsArray(arr_session_cycle_info)) then
                '取出這個週期內的上課筆數(排除return)
                arr_client_attend_list = getCustomerBookingCount(int_client_sn, arr_session_cycle_info(2), arr_session_cycle_info(3), int_db_connect)
                int_week_use_session_count = CInt(arr_client_attend_list(4))
            end if
        end if
        '**** 取得課程在結算週期內的資訊，只有大陸定時定額產品才需要判斷 END   ****


        '**** 判斷是否有需要取消的課程 START ****
        '重新定義 成功預訂課程的日期和扣堂 陣列
        Redim arr_success_cancel_class(2, Ubound(str_arr_total_cancel_class_sn))
        for i = 0 to Ubound(str_arr_total_cancel_class_sn)
            if (Not isEmptyOrNull(str_arr_total_cancel_class_sn(i)) and _
                Not isEmptyOrNull(str_arr_total_cancel_class_date(i)) and _
                Not isEmptyOrNull(str_arr_total_cancel_class_time(i)) and _
                Not isEmptyOrNull(str_arr_total_cancel_class_type(i))) then

                '課程日期/時間
                str_class_datetime = str_arr_total_cancel_class_date(i) & " " &_
                                     Right("0" & str_arr_total_cancel_class_time(i), 2) & ":" &_
                                     str_class_minute
            
                arr_chk_cancel_param = Array(str_now_use_product_sn, _
                                             Trim(str_arr_total_cancel_class_type(i)), _
                                             Trim(str_arr_total_cancel_class_date(i)), _
                                             Trim(str_arr_total_cancel_class_time(i)))
                arr_allow_cancel = isAllowCancelClass(int_client_sn, _
                                               arr_chk_cancel_param, _
                                               int_db_connect)    
                if (Not IsArray(arr_allow_cancel)) then
                    int_total_fail_cancel_count = int_total_fail_cancel_count + 1
                    if (isEmptyOrNull(str_debug_mode)) then
                        str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2)
                    else
                        str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2) &_
                                              "(不允許取消課程，isAllowCancelClass，函式傳入值錯誤)"
                    end if
                else
                    '不允許取消
                    if (arr_allow_cancel(0) <> CONST_FUNC_EXE_SUCCESS) then
                        int_total_fail_cancel_count = int_total_fail_cancel_count + 1
                    
                        '不能取消課程的原因 - 因為在上課前幾小時內不能取消課程
                         if (CONST_ERROR_NO_CANCEL_IN_TIME = arr_allow_cancel(0)) then
                            str_cancel_fail_msg = str_cancel_fail_msg & Replace(str_arr_fail_msg(3), "$hour$", (arr_allow_cancel(1)\60))
                        '不能取消課程的原因 - 已超過上課時間
                         elseif (CONST_ERROR_NO_CANCEL_OVER_TIME = arr_allow_cancel(0)) then
                            str_cancel_fail_msg = str_cancel_fail_msg & str_arr_fail_msg(4)
                        '不能取消課程 - 產品對應課程類型的扣點規則未設定
                         elseif (CONST_ERROR_NO_CANCEL_NO_RULE = arr_allow_cancel(0)) then
                            if (isEmptyOrNull(str_debug_mode)) then
                                str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2)
                            else
                                '不能取消課程 - 產品對應課程類型的扣點規則未設定'
                                str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2) &_
                                                      "(產品對應課程類型的扣點規則未設定)"
                            end if
                        '其他'
                         else
                            if (isEmptyOrNull(str_debug_mode)) then
                                str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2)
                            else
                                'isAllowCancelClass 函式回傳值錯誤
                                str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2) &_
                                                      "(isAllowCancelClass 函式回傳值錯誤)"
                            end if
                         end if
                    '允許取消
                    else
                        arr_cancel = cancelClass(Trim(str_arr_total_cancel_class_sn(i)), int_client_sn, _
                                                 int_db_connect)
    
                        if (IsArray(arr_cancel)) then

                            '取消失敗
                            if (CONST_FUNC_EXE_SUCCESS <> arr_cancel(1)) then
                                int_total_fail_cancel_count = int_total_fail_cancel_count + 1
                                '並沒有預訂該堂課程
                                if (CONST_ERROR_NO_ORDER_CLASS = arr_cancel(2)) then
                                    '取消失敗(此堂課您並沒有預定)
                                    str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(5)
                                '取消課程失敗
                                elseif (CONST_CANCEL_FAIL = arr_cancel(2)) then
                                    if (isEmptyOrNull(str_debug_mode)) then
                                        str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2)
                                    else
                                        str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2) &_
                                                              "(cancelClass函式，update client_attend_list 失敗)"
                                    end if
                                else
                                    if (isEmptyOrNull(str_debug_mode)) then
                                        str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2)
                                    else
                                        str_cancel_fail_msg = str_cancel_fail_msg & str_class_datetime & " " & str_arr_fail_msg(2) &_
                                                              "(cancelClass函式，異常的回傳值)"
                                    end if
                                end if
                        
                            '取消成功
                            else
                                int_total_success_cancel_count = int_total_success_cancel_count + 1
                                '紀錄已成功訂課的資訊
                                '(0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn
                                arr_success_cancel_class(0, i) = arr_cancel(5)
                                arr_success_cancel_class(1, i) = arr_cancel(4)
                                arr_success_cancel_class(2, i) = arr_cancel(2)
                                intCancelClassType = sCInt(arr_cancel(3))

                                '更新要顯示的取消成功訊息
                                if (CONST_CLASS_TYPE_LOBBY = intCancelClassType) then
                                    str_class_type_description = str_arr_success_msg(0)     'getWord("SPECIAL_SESSION") '"大會堂"
                                    strClassType = "大會堂"
                                elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_THREE) = intCancelClassType OR sCInt(CONST_CLASS_TYPE_ONE_ON_FOUR) = intCancelClassType) then
                                    str_class_type_description = str_arr_success_msg(1)     'getWord("NORMAL_CLASS_NOTE_1") '"小班(1-3人)"
                                    'strClassType = "一對三"
                                    strClassType = "小班制"
                                elseif (CONST_CLASS_TYPE_ONE_ON_ONE = intCancelClassType) then
                                    str_class_type_description = str_arr_success_msg(2)     'getWord("VIP_CLASS_NOTE_1") '"1对1"
                                    strClassType = "一對一"
                                '一對六
                                elseif (CONST_CLASS_TYPE_NORMAL = intCancelClassType) then
                                    'TODO: 目前大陸客戶一對多課程 當成是一對三
                                    if ("C" = UCase(str_client_website)) then
                                        str_class_type_description = str_arr_success_msg(1)     'getWord("NORMAL_CLASS_NOTE_1") '"小班(1-3人)"
                                        intCancelClassType = CONST_CLASS_TYPE_ONE_ON_THREE
                                        strClassType = "一對三"
                                    else
                                        str_class_type_description = str_arr_success_msg(3)     '"1对6"
                                    end if
                                end if

                                '更新取消課程成功的訊息
                                str_cancel_success_msg = str_cancel_success_msg &_
                                                         arr_cancel(5) & " " & arr_cancel(6) & ":" & str_class_minute & " " &_
                                                         str_class_type_description
                                
                                '20120926 阿捨新增 特殊關懷訊息
                                if ( int_total_success_cancel_count > 1 ) then
                                    strComma = ","
                                else
                                    strComma = ""
                                end if
                                strClassDatetime = arr_cancel(5) & " " & arr_cancel(6) & ":" & str_class_minute
                                strCancelClassInfo = strCancelClassInfo & strComma & right(getFormatDateTime(strClassDatetime, 1), 14)

                                '更新要寄送email給客戶的取消課程成功 課程訊息
                                str_cancel_success_msg_for_email = str_cancel_success_msg_for_email & getClassMsgForEmail(intCancelClassType, arr_cancel(5), arr_cancel(6), arr_cancel(0)) & "$br$"

                                flt_total_cancel_return_session = flt_total_cancel_return_session + CDbl(arr_cancel(4))
                                flt_client_total_session = flt_client_total_session + CDbl(arr_cancel(4))
								
								'如果是Special 1 on 1 就寄信
								if (instr(special1on1value,"#"&replace(arr_cancel(5),"/","")&right("0"&arr_cancel(6),2)&"#")>0) then
									Set objInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
									sql = "SELECT  a.sn , "
									sql = sql & "        b.userid+'@tutorabc.com' submitterEmail, "
									sql = sql & "        CONVERT(VARCHAR(10), SessionDate, 111) + ' ' + RIGHT('0'+ CONVERT(VARCHAR(2), SessionTime),2)+':30' SessionDateTime , "
									sql = sql & "        c.basic_fname+' '+c.basic_lname consultantName, "
									sql = sql & "        a.ClientSn, "
									sql = sql & "        d.lead_sn, "
									sql = sql & "        d.cname "
									sql = sql & "FROM    Special1on1Session a WITH ( NOLOCK ) "
									sql = sql & "INNER JOIN staff_basic b WITH (NOLOCK) ON a.Creator = b.sn "
									sql = sql & "LEFT JOIN con_basic c WITH (NOLOCK) ON a.Consultant = c.con_sn "
									sql = sql & "INNER JOIN client_basic d WITH (NOLOCK) ON a.ClientSn = d.sn "
									sql = sql & "WHERE   ClientBookingSn = "&arr_cancel(2)&" "
									sql = sql & "        AND a.Valid = 1 " '有效, 未被顧問取消或客戶取消
									sql = sql & "        AND ClientPlatform = 2 " 'VIPABC
									sql = sql & "        AND a.Status NOT IN ( 7, 8 ) "
									intQueryResult = objInfo.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_RW_CONN)
									if not objInfo.eof then
										'更新狀態
										Set rs1 = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
										sql = "UPDATE Special1on1Session SET Status=7,isProcess=0,Valid=0 WHERE sn=@sn"
										arrParam = Array(objInfo("sn"))
										intQueryResult = rs1.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_RW_CONN)
										rs1.close
										Set rs1=nothing
										
										Dim SendMailabc :set SendMailabc = new ABCSystemSendMail
										strEmailTo = objInfo("submitterEmail")&",CSOP@tutorabc.com"
										strSessionDateTime = objInfo("SessionDateTime")
										strConsultantName = objInfo("consultantName")
										SendMailabc.str_sendlettermail_sn = "special1on1_0001"
										SendMailabc.strEmailTitle = "Special 1 on 1 客戶 "&objInfo("cname")&" 取消訂位"
										SendMailabc.client_cname = objInfo("cname") '客戶姓名
										SendMailabc.str_para2 = "http://"&CONST_VIPABC_IMSHOST_NAME&"/asp/cs/client_3.asp?lead="&objInfo("lead_sn")&"&client="&objInfo("ClientSn")&"&search=yes"
										SendMailabc.str_para4 = strSessionDateTime
										SendMailabc.str_para5 = strConsultantName
										if CFG_SPECIAL1ON1_DEBUG_MODEL = 1 then			
											SendMailabc.strEmailToAddr = CFG_SPECIAL1ON1_MAIL_TO '客戶的電子郵件
											testMailTo = strEmailTo
											SendMailabc.str_para1 = "<div style='background:#99FF88'>{內為正式站時會發的人mail 正式機時不會看到此訊息: To: "&testMailTo &" / CC:}</div>"
											result = SendMailabc.MainSendMail()
										else
											SendMailabc.strEmailToAddr = strEmailTo
											result = SendMailabc.MainSendMail()
										end if
										Set SendMailabc = nothing
									end if
									objInfo.close
									Set objInfo=nothing
								end if
								
                            end if
                        end if 'end of if (IsArray(arr_cancel)) then
                    end if  'end of 允許或不允許取消課程
                end if 'end of if (Not IsArray(arr_allow_cancel)) then
            end if 'end of 取消課程的資料驗證'
        next 'for str_arr_total_cancel_class_sn'

        '定時定額產品跑以下規則
        if (int_total_success_cancel_count > 0 and CONST_PRODUCT_QUOTA = int_client_purchase_product_type) then
            arr_tmp = Array(int_client_sn, str_purchase_account_sn, str_now_use_product_sn, _
                            int_account_cycle_day, int_week_deduct_session, "cancel", _
                            int_week_use_session_count, int_website_code, _
                            Request.ServerVariables("PATH_INFO"), 0, str_service_start_date, str_service_end_date)
            
            '檢查訂課時間，當週和非當週的總扣堂數，更新到客戶剩餘點數總表
            Call updatePointAfterReservationOrCancel(arr_success_cancel_class, arr_tmp, int_db_connect)
        end if
    end if 'end of if (Not isEmptyOrNull(str_total_cancel_class_sn)) then
    '******************* 取消課程 END *******************
    

    '******************* 預定/重訂課程 START *******************  
    if (Not isEmptyOrNull(str_total_reservation_class_dt)) then
		
		 '若取不到客戶的剩餘堂數
		if (isEmptyOrNull(flt_client_total_session) and isEmptyOrNull(flt_client_bonus_session) and isEmptyOrNull(flt_client_week_session)) then
			if (isEmptyOrNull(str_debug_mode)) then
				'"合约剩余课时数不足，请至「学习记录」查询详细资讯。"
				Call alertGo(str_arr_fail_msg(1), "/index.asp", CONST_ALERT_THEN_REDIRECT)
			else
				Call alertGo(str_arr_fail_msg(1) & "(客戶的剩餘堂數是空的)", "/index.asp", CONST_ALERT_THEN_REDIRECT)
			end if
			Response.End
		elseif (flt_client_total_session <= 0 and flt_client_bonus_session <= 0 and flt_client_week_session <= 0) then
			if (isEmptyOrNull(str_debug_mode)) then
				'"合约剩余课时数不足，请至「学习记录」查询详细资讯。"
				Call alertGo(str_arr_fail_msg(1), "/index.asp", CONST_ALERT_THEN_REDIRECT)
			else
				Call alertGo(str_arr_fail_msg(1) & ")", "/index.asp", CONST_ALERT_THEN_REDIRECT)
			end if
			Response.End
		end if
		
        '預訂的課程 class_type + date + time
        str_arr_total_reservation_class = Split(str_total_reservation_class_dt, ",")

        '重新定義 成功預訂課程的日期和扣堂 陣列
        Redim arr_success_reservation_class(2, Ubound(str_arr_total_reservation_class))

        for i = 0 to Ubound(str_arr_total_reservation_class)
            bol_enabled_set_class = true
        
            str_reservation_class_datetime = Trim(str_arr_total_reservation_class(i))
            if (Not isEmptyOrNull(str_reservation_class_datetime)) then

                '課程資訊的字串格式不正確
                if (Len(Trim(str_reservation_class_datetime)) < 12) then
                    int_total_fail_order_count = int_total_fail_order_count + 1
                    '更新預定課程失敗的訊息
                    if (isEmptyOrNull(str_debug_mode)) then
                        '預定失敗
                        str_order_fail_msg = str_order_fail_msg & str_arr_fail_msg(6)   'getWord("ORDER_FAIL") & "<br/>"
                    'DEBUG MODE 參數 若有設定的話 會輸出比較多的錯誤訊息給工程師看
                    else
                        str_order_fail_msg = str_order_fail_msg & str_arr_fail_msg(6) & "(課程資訊的字串格式不正確)"
                    end if
                else
                    '取得課程日期
                    dat_attend_date =  Mid(Trim(str_reservation_class_datetime), 3, 4) &_
                                       "/" & Mid(Trim(str_reservation_class_datetime), 7, 2) &_
                                       "/" & Mid(Trim(str_reservation_class_datetime), 9 , 2)

                    '取得課程時間(24H)
                    int_attend_sestime = Mid(str_reservation_class_datetime, 11, 2)

                    '取得課程類型
                    int_attend_type_tmp = Int(Left(str_reservation_class_datetime, 2))
                    
                    if ( true = bolDebugMode ) then
                        response.Write "int_attend_type_tmp: " & int_attend_type_tmp & "<br />"
                    end if
                    '若是大會堂課程 則 attend_livesession_types 還是塞4 所以將原始的先存到 int_attend_type_tmp 供其他函式和判斷式使用
                    
                    '一對一和一對三和一對六 教室的等級就是客戶的等級(+-1)
                    if (int_attend_type_tmp <> CONST_CLASS_TYPE_LOBBY) then
                        int_special_sn = 0 '大會堂才有值
                    '大會堂 教室的等級是 lobby_session.sn
                    else
                        int_special_sn = Mid(str_reservation_class_datetime, 13, Len(str_reservation_class_datetime)-12)
                    end if

                    '課程日期/時間
                    str_class_datetime = dat_attend_date & " " &_
                                         Right("0" & int_attend_sestime, 2) & ":" &_
                                         str_class_minute

                    '**** 排除掉已超過上課時間的 START ****
                    if (compareCurrentTime(str_class_datetime, "n") <> CONST_TIME_AFTER) then
                        bol_enabled_set_class = false
                        int_total_fail_order_count = int_total_fail_order_count + 1
                        '預定失敗(已經超過上課時間的課程無法預訂)
                        str_order_fail_msg = str_order_fail_msg & str_class_datetime & str_arr_fail_msg(7)
                    end if
                    '**** 排除掉已超過上課時間的 END ****

                    '******************* 訂課日期必須>=合約開始日且<=合約結束日 (檢查是否符合資格可以訂課) START *******************
                    '合約開始日之前
                    if (compareTime(dat_attend_date, str_service_start_date, "d") = CONST_TIME_BEFORE) then
                        bol_enabled_set_class = false
                        int_total_fail_order_count = int_total_fail_order_count + 1
                        '預定失敗(請預訂YYYY/MM/DD(合約開始)之後的課程。)
                        str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                             getClassTypeDescription(int_attend_type_tmp) &_
                                             Replace(str_arr_fail_msg(8), "$contract_service_start$", str_service_start_date)
                    '合約結束日之後
                    elseif (compareTime(dat_attend_date, str_service_end_date, "d") = CONST_TIME_AFTER) then
                        bol_enabled_set_class = false
                        int_total_fail_order_count = int_total_fail_order_count + 1
                        '預定失敗(請預訂YYYY/MM/DD(合約結束)之前的課程)'
                        str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                             getClassTypeDescription(int_attend_type_tmp) &_
                                             Replace(str_arr_fail_msg(9), "$contract_service_end$", str_service_end_date)
                    end if
                    '******************* 訂課日期必須>=合約開始日且<=合約結束日 (檢查是否符合資格可以訂課) END   *******************
           
                    if (true = bol_enabled_set_class) then

                        '取得要扣的堂數 
                        flt_subtract_session = getSubtractSession(Array(str_class_datetime, int_attend_type_tmp, str_now_use_product_sn, int_special_sn), _
                                                                  int_db_connect)

                        flt_subtract_session = CDbl(flt_subtract_session)
                        '******************* 該時段並無開課或不允許訂課 (檢查是否符合資格可以訂課) START *******************
                        if (flt_subtract_session < 0) then
                
                            '若是在四小時內的話 有可能是因為訂課頁面放太久才送出 導致四小時前的課程變成四小時內 所以還是讓客戶訂
                            if (flt_subtract_session = -2) then
                                flt_subtract_session = getClassSubtractSession(int_attend_type_tmp, Array(str_now_use_product_sn, "under4"), int_db_connect)
                            end if

                            if (flt_subtract_session < 0) then
                                bol_enabled_set_class = false
                                int_total_fail_order_count = int_total_fail_order_count + 1

                                '訂課失敗的錯誤訊息
                                if (flt_subtract_session = -2) then
                                    '"四小时前没有开课，请选择其他上课日期或咨询时段"'
                                    str_subtract_session_error_msg = str_arr_fail_msg(10)   'getWord("CLASS_NO_OPEN_IN_4_HOUR")
                                elseif (flt_subtract_session = -3) then
                                    '"24小时内临时订课仅开放一对多，请选择其他上课日期或咨询时段"
                                    str_subtract_session_error_msg = str_arr_fail_msg(11)   'getWord("CLASS_NO_OPEN_IN_24_HOUR")
                                elseif (flt_subtract_session = -4) then
                                    if (isEmptyOrNull(str_debug_mode)) then
                                        '"訂位失敗，請與客服人員連絡：02-3365-9999"
                                        str_subtract_session_error_msg = str_arr_fail_msg(0) 'getWord("RESERVATION_CLASS_FAIL_CONTACT_CS") & CONST_SERVICE_PHONE
                                    'DEBUG MODE 參數 若有設定的話 會輸出比較多的錯誤訊息給工程師看
                                    else
                                        str_subtract_session_error_msg = str_arr_fail_msg(0) & "(錯誤的訂課日期和時間)"
                                    end if
                                elseif (flt_subtract_session = -5) then
                                    '"大会堂只允许开课前五分钟可以订课，请选择其他上课日期或咨询时段"
                                    str_subtract_session_error_msg = str_arr_fail_msg(12)   'getWord("CLASS_NO_OPEN_IN_5_MINUTE")
                                end if
                        
                                if (Not isEmptyOrNull(str_subtract_session_error_msg)) then
                                    str_subtract_session_error_msg = "(" & str_subtract_session_error_msg & ")"
                                end if

                                '更新預定課程失敗的訊息
                                '預定失敗
                                str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                                     getClassTypeDescription(int_attend_type_tmp) &_
                                                     str_arr_fail_msg(6) & str_subtract_session_error_msg
                            end if 'end of if (flt_subtract_session < 0) then
                        end if
                        '******************* 該時段並無開課或不允許訂課 (檢查是否符合資格可以訂課) END    *******************
                    end if 'end of if (bol_enabled_set_class) then

                    if (true = bol_enabled_set_class) then
                        '******************* 確認目前剩餘堂數是否足夠 (檢查是否符合資格可以訂課) START *******************
                        if (((flt_client_total_session + flt_client_week_session - flt_total_subtract_session) < flt_subtract_session  ) And _
							((flt_client_bonus_session + flt_client_week_session - flt_total_subtract_session) < flt_subtract_session  )) then
                            bol_enabled_set_class = false
                            int_total_fail_order_count = int_total_fail_order_count + 1
                
                            '更新預定課程失敗的訊息
                            '預定失敗(合约剩余课时数不足)
                            str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                                 getClassTypeDescription(int_attend_type_tmp) &_
                                                 str_arr_fail_msg(13)                    
                        end if
                        '******************* 確認目前剩餘堂數是否足夠 (檢查是否符合資格可以訂課) END   *******************
                        '20120723 阿捨新增 套餐產品堂數						 
						'免费大会堂2堂，一对三真人授课3堂
                        '20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493                    
                        '20121112 阿捨新增 套餐產品改版
                        '改讀/program/member/reservation_class/getComboProductData.asp
                        '當主架構判斷無誤時, 再判斷套餐產品的上限數
                        if ( true = bolComboProduct  AND true = bol_enabled_set_class ) then
                            dim intHavingSession : intHavingSession = 0
                            '20130102 阿捨新增 套餐產品判斷 for 贈送堂數
                            dim intComboGiftSession : intComboGiftSession = 0
                            dim intRealSessionPoint : intComboGiftSession = 1
						    '可上堂數
                            intHavingSession = getComboProductSession(int_client_sn, str_now_use_contract_sn, int_attend_type_tmp, 2, 0, CONST_VIPABC_R_CONN)
                            intComboGiftSession = getComboProductGift(str_now_use_contract_sn, "", int_attend_type_tmp, "", 0, CONST_VIPABC_R_CONN)
                            intRealSessionPoint = getSessionTypeRealPoints(str_now_use_contract_sn, int_attend_type_tmp, 0, CONST_VIPABC_R_CONN)
                            intComboGiftSession = sCDbl(intComboGiftSession) * sCDbl(intRealSessionPoint)
                            if ( true = bolDebugMode ) then
                                response.Write "intHavingSession:" & intHavingSession & "<br/>"
                                response.Write "intComboGiftSession:" & intComboGiftSession & "<br/>"
                                response.Write "flt_subtract_session:" & flt_subtract_session & "<br/>"
                            end if
                            if ( (sCInt(intHavingSession) + sCInt(intComboGiftSession)) < sCInt(flt_subtract_session) ) then
                            bol_enabled_set_class = false
                            int_total_fail_order_count = int_total_fail_order_count + 1
                                
                               
                            '更新預定課程失敗的訊息
                            '預定失敗(合约剩余课时数不足)
                            str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                                    getClassTypeDescription(int_attend_type_tmp) &_
                                                    str_arr_fail_msg(13)                    
                            end if
                        end if
                    end if
                    
                    strClassType = "大會堂"
                    '通過檢查 開始檢查要新增或修改或刪除 客戶預定的課程 START
                    if (true = bol_enabled_set_class) then                
                        '若是大陸人的一般產品 則一對一都扣兩堂，一對三都扣一堂
                        'VIPABC
                        if ( UCase(sCStr(str_client_website)) = "C" ) then

                            '修改訂課堂術規則 tree--開始--
                            '取得 一般產品 或 定期定額 預計扣除堂數( 產品編號, 課程類型編號1v1 1v3, 訂課時間, 連線站台 )
                            flt_subtract_session = getFixedAmountTimeSubtractSession(str_now_use_product_sn, int_attend_type_tmp, str_class_datetime, CONST_VIPABC_R_CONN)  '取得要扣的堂數 
                            if ( flt_subtract_session = "Error" ) then
                                Call alertGo("(订课资料异常请重新执行)", str_redirect_url, CONST_ALERT_THEN_REDIRECT)'可能是產品編號不存在 或 課程類型編號錯誤 或 資料庫異常
                                response.end'資料異常
                            end if
                            '修改訂課堂術規則 tree--結束--
                            
                            '一對三和六
                            if (CONST_CLASS_TYPE_NORMAL = int_attend_type_tmp or _
                                CONST_CLASS_TYPE_ONE_ON_THREE = int_attend_type_tmp or _
                                CONST_CLASS_TYPE_ONE_ON_FOUR = int_attend_type_tmp) then
                               'strClassType = "一對三"
                               strClassType = "小班制"
                            '一對一
                            elseif (CONST_CLASS_TYPE_ONE_ON_ONE = int_attend_type_tmp or _
                                    1 = int_attend_type_tmp) then
                                strClassType = "一對一"
                            else
                                strClassType = "大會堂"
                            end if
                        end if

                        arr_reservation_result = reservationClass(str_reservation_class_datetime, int_client_sn, _
                                                                  Array(str_now_use_product_sn, flt_subtract_session, int_client_attend_level), _
                                                                  int_db_connect)
                        
                        '失敗 (傳入參數有問題)
                        if (Not IsArray(arr_reservation_result)) then
                            int_total_fail_order_count = int_total_fail_order_count + 1
                            if (isEmptyOrNull(str_debug_mode)) then
                                str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " & str_arr_fail_msg(0)
                            else
                                str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " & str_arr_fail_msg(0) &_
                                                     "(reservationClass函式，傳入參數錯誤)"
                            end if
                        else
                            '成功
                            if (arr_reservation_result(1) = CONST_FUNC_EXE_SUCCESS) then

                                '紀錄已成功訂課的資訊
                                '(0,0) = 日期, (1,0)=扣堂, (2,0)=client_attend_list.sn
                                arr_success_reservation_class(0, i) = dat_attend_date
                                arr_success_reservation_class(1, i) = flt_subtract_session
                                arr_success_reservation_class(2, i) = arr_reservation_result(2)

                                '更新總共要扣的堂數
                                flt_total_subtract_session = flt_total_subtract_session + flt_subtract_session
                                
                                '更新總共成功預定的課程筆數
                                int_total_success_order_count = int_total_success_order_count + 1
                                
                                '更新要顯示在頁面上的訂課成功 課程訊息
                                str_order_success_msg = str_order_success_msg & str_class_datetime & " " &_
                                                        getClassTypeDescription(int_attend_type_tmp) & "，$br$"
                                
                                '20120926 阿捨新增 特殊關懷訊息
                                if ( int_total_success_order_count > 1 ) then
                                    strComma = ","
                                else
                                    strComma = ""
                                end if
                                 strOrderClassInfo = strOrderClassInfo & strComma & right(str_class_datetime, 14)


                                '更新要寄送email給客戶的訂課成功 課程訊息
                                str_order_success_msg_for_email = str_order_success_msg_for_email &_
                                                                  getClassMsgForEmail(int_attend_type_tmp, dat_attend_date, int_attend_sestime, arr_reservation_result(0)) & "$br$"

                            '失敗
                            else
                                int_total_fail_order_count = int_total_fail_order_count + 1
                                '該時段的課程已經預訂了
                                if (arr_reservation_result(2) = CONST_ERROR_ALREADY_ORDER) then
                                    ''此堂课您已经预定过了'
                                    str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                                         getClassTypeDescription(int_attend_type_tmp) &_
                                                         str_arr_fail_msg(14)

                                '重訂課程失敗
                                elseif (arr_reservation_result(2) = CONST_RE_RESERVATION_FAIL) then
                                    '更新預定課程失敗的訊息
                                    '重新预定失败
                                    str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                                         getClassTypeDescription(int_attend_type_tmp) &_
                                                         str_arr_fail_msg(15)
                            
                                '預訂課程失敗
                                elseif (arr_reservation_result(2) = CONST_RESERVATION_FAIL) then
                                    '更新預定課程失敗的訊息
                                    if (isEmptyOrNull(str_debug_mode)) then
                                        '預定失敗
                                        str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                                             getClassTypeDescription(int_attend_type_tmp) &_
                                                             str_arr_fail_msg(6)
                                    'DEBUG MODE 參數 若有設定的話 會輸出比較多的錯誤訊息給工程師看
                                    else
                                        str_order_fail_msg = str_order_fail_msg & str_class_datetime & " " &_
                                                             getClassTypeDescription(int_attend_type_tmp) &_
                                                             str_arr_fail_msg(6) & "(reservationClass insert client_attend_list 失敗)"
                                    end if
                                end if
                            end if  'end of 新增成功或失敗

                        end if  'end of if (IsArray(arr_reservation_result)) then

                    end if
                    '通過檢查 開始檢查要新增或修改或刪除 客戶預定的課程 END
                end if 'end of 課程資訊的字串格式不正確
            end if 'end of if (Not isEmptyOrNull(str_reservation_class_datetime)) then'
        next ' for 訂課的時間'

        '若有訂課成功的話 就發信給客戶
        if (int_total_success_order_count > 0) then
            '定時定額產品跑以下規則
            if (CONST_PRODUCT_QUOTA = int_client_purchase_product_type) then
                arr_tmp = Array(int_client_sn, str_purchase_account_sn, str_now_use_product_sn, _
                                int_account_cycle_day, int_week_deduct_session, "reservation", _
                                int_week_use_session_count, int_website_code, _
                                Request.ServerVariables("PATH_INFO"), 0, str_service_start_date, str_service_end_date)
                   
                '檢查訂課時間，當週和非當週的總扣堂數，更新到客戶剩餘點數總表
                Call updatePointAfterReservationOrCancel(arr_success_reservation_class, arr_tmp, int_db_connect)
            end if
        end if
    end if 'end of if (Not isEmptyOrNull(str_total_reservation_class_dt)) then
    '******************* 預定/重訂課程 END    *******************  
    '發送信件給客戶
    if (int_total_success_order_count > 0 or int_total_success_cancel_count > 0) then
'        Call sendOrderCancelEmail(str_class_email_url, str_client_login_email_addr, _
'                                  str_client_full_ename, str_order_success_msg_for_email & str_cancel_success_msg_for_email, _
'								  str_client_website)

		'20110214 tree VB10110205 VIPABC系統信件 --開始--
		set SendMailvip = new VIPSystemSendMail '建立class物件
		SendMailvip.str_sendlettermail_sn = "0008" '第0008封信 主旨：使用课程预约新增						
		SendMailvip.strEmailToAddr = str_client_login_email_addr '客戶的電子郵件
		SendMailvip.client_cname = str_client_full_ename '客戶姓名
		SendMailvip.client_appointment_date = str_order_success_msg_for_email&str_cancel_success_msg_for_email '定退課資訊
		Call SendMailvip.MainSendMail() '呼叫class物件的函式來發信
		set SendMailvip = nothing '釋放CLASS物件
		'20110214 tree VB10110205 VIPABC系統信件 --結束--  

		'20100729 joyce 查詢客戶專屬客服
		dim arr_result : arr_result = ""
		dim str_staff_email : str_staff_email = ""
		arr_result = getStaffHandleData(int_client_sn)
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) > 0) then  	'有資料
				str_staff_email = arr_result(1,0)&"@tutorabc.com"
				str_mail_to_management = str_staff_email &","& str_mail_to_management
			else
				'無資料
			end if 
		end if

            '20110816 tree
            '**** 特殊規則 檢查VIP身分是不是博士友人 是的話 發信 START ****
            Call checkDrFriendThenSendMail(int_client_sn, str_class_email_url, _
                                          str_client_full_ename, str_order_success_msg_for_email & str_cancel_success_msg_for_email, _
                                          str_client_website, int_db_connect)
            '**** 特殊規則 檢查VIP身分是不是博士友人 是的話 發信 END   ****
            '20110816 tree

			'發送信件給設定的主管和經理
			str_class_email_subject = "客戶 " & str_client_full_ename & " " & str_class_email_subject
			'Call sendOrderCancelEmail(str_class_email_url, str_mail_to_management, _
										'str_client_full_ename, str_order_success_msg_for_email & str_cancel_success_msg_for_email, _
										'str_client_website)
			'20101123 阿捨新增 訂課RD相關關懷
			if (isHardCodeComplied("customer_rd_care_client_sn", int_client_sn)) then 
				Call sendOrderCancelEmail(str_class_email_url, CONST_MAIL_TO_RD_CARE_CUSTOMER_ACCOUNT, _
										str_client_full_ename, str_order_success_msg_for_email & str_cancel_success_msg_for_email, _
										str_client_website)
			end if
            
            '20101123 阿捨新增 訂課CS相關關懷
			if (isHardCodeComplied("customer_cs_care_client_sn", int_client_sn)) then 
				Call sendOrderCancelEmail(str_class_email_url, CONST_MAIL_TO_CS_SPECIAL_CARE_CUSTOMER_ACCOUNT, _
										str_client_full_ename, str_order_success_msg_for_email & str_cancel_success_msg_for_email, _
										str_client_website)
			end if

        '20110614 阿捨新增 特殊客戶關懷
        '20120926 阿捨新增 特殊客戶關懷include檔
        dim intNoifiyType : intNoifiyType = ""
        dim bolOrder : bolOrder = false
        dim bolCancel : bolCancel = false
        dim intOrderClassNum : intOrderClassNum = 0
        dim intCancelClassNum : intCancelClassNum = 0
        intNoifiyType = 2
        if ( int_total_success_order_count > 0 ) then
            intOrderClassNum = int_total_success_order_count
            bolOrder = true
        end if
        if ( int_total_success_cancel_count > 0 ) then
            intCancelClassNum = int_total_success_cancel_count
            bolCancel= true
        end if
%>
<!--#include virtual="/program/member/CareClientNotification.asp"-->
<%
end if

Dim strShowResult : strShowResult = ""
'預約成功訊息
if (int_total_success_order_count > 0) then
    strShowResult = strShowResult & int_total_success_order_count & " 堂课程预约成功\n"
    strShowResult = strShowResult & Replace(str_order_success_msg, "$br$", "\n")
    if ( true = bolUnlimited ) then
    else
        '预计扣除课时数
        strShowResult = strShowResult & getWord("SUBTRACT_SESSION_POINT") & " " & flt_total_subtract_session & "\n\n"
    end if
end if

'預約失敗訊息
if (int_total_fail_order_count > 0) then
    strShowResult = strShowResult & int_total_fail_order_count & " 堂课程预约失败\n"
    strShowResult = strShowResult & Replace(str_order_fail_msg, "$br$", "\n") & "\n\n"
end if

'取消成功訊息
if (int_total_success_cancel_count > 0) then
    strShowResult = strShowResult & int_total_success_cancel_count & " 堂课程取消成功\n"
    strShowResult = strShowResult & Replace(str_cancel_success_msg, "$br$", "\n")
    if ( true = bolUnlimited ) then
    else
        '预计归还课时数'
        strShowResult = strShowResult & "预计归还课时数 " & flt_total_cancel_return_session & "\n\n"
    end if  
end if

'取消失敗訊息
if (int_total_fail_cancel_count > 0) then
    strShowResult = strShowResult & int_total_fail_cancel_count & " 堂课程取消失败\n"
    strShowResult = strShowResult & Replace(str_cancel_fail_msg, "$br$", "\n") & "\n\n"
end if

if (isEmptyOrNull(strShowResult)) then
    strShowResult = str_arr_fail_msg(16)
end if                    

'若大會堂訂課完畢 導回大會堂訂課頁面
dim strSessionWord : strSessionWord = getRequest("session", CONST_DECODE_NO)    
dim strTypeWord : strTypeWord = getRequest("type", CONST_DECODE_NO)   '確認訂課後 判斷載入頁面
if ( "lobby" = strSessionWord ) then
    str_redirect_url = "/program/member/reservation_class/lobby_near_order.asp?type=" & strTypeWord
end if

if ( true = bolDeBugMode ) then
    response.Write str_order_fail_msg
else
    Call alertGo(strShowResult, str_redirect_url, CONST_ALERT_THEN_REDIRECT)
    Response.End
end if
%>