<%
'/****************************************************************************************
'程式/元件簡述(目的) ： 所有訂課頁面 要用到的訊息 包括 javascript alert or 控制項的文字
'備註：'被 /program/member/reservation_clas/normal_and_vip.asp include
       '被 /program/member/reservation_class/normal_and_vip_near_order.asp include
'歷程(作者/日期)：[Ray Chien] 2009/12/22 Created
'*****************************************************************************************/


'TODO: 判斷語系讀文字檔
'/****************************************************************************************
'描述： 控制項的文字
'參數描述 / 建立人員 / 建立日期：
'*****************************************************************************************/
'<a>此次预订課程</a>
'include/ajax_reservation_normal_and_vip_class.asp
'normal_and_vip_near_cancel.asp
Dim str_msg_a_have_order_class : str_msg_a_have_order_class = getWord("HAVE_ORDER_CLASS")   '"此次预订課程"

'<a>近期取消課程</a>
'normal_and_vip_near_cancel.asp
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_a_have_cancel_class : str_msg_a_have_cancel_class = getWord("HAVE_CANCEL_CLASS")  '"近期取消課程"

'<a>近期预订大会堂课程</a>
'include/ajax_lobby_class_search.asp
'lobby_near_cancel.asp
Dim str_msg_a_have_order_lobby_class : str_msg_a_have_order_lobby_class = getWord("HAVE_ORDER_LOBBY_CLASS")  '"近期预订大会堂课程"

'<a>近期取消大会堂课程</a>
'include/ajax_lobby_class_search.asp
'lobby_near_cancel.asp
Dim str_msg_a_have_cancel_lobby_class : str_msg_a_have_cancel_lobby_class = getWord("HAVE_CANCEL_LOBBY_CLASS")  '"近期取消大会堂课程"

'<a>您此次尚未预订谘询课程</a>
'include/ajax_lobby_class_search.asp
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_a_no_have_order_class : str_msg_a_no_have_order_class = getWord("NO_HAVE_ORDER_CLASS")  '"尚未预订谘询课程"

'<a>您近期尚未取消課程</a>
'include/ajax_lobby_class_search.asp
'lobby_near_cancel.asp
'normal_and_vip_near_cancel.asp
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_a_no_have_cancel_class : str_msg_a_no_have_cancel_class = getWord("NO_HAVE_CANCEL_CLASS") '"近期尚未取消課程"

'请选择课程类型 / 时间 (一對一和一對多新增)
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_font_normal_class_head : str_msg_font_normal_class_head = getWord("SELECT_CLASS_TYPE_TIME")   '"请选择课程类型 / 时间"

'快速选择大会堂 (大會堂新增)
'include/ajax_lobby_class_search.asp
Dim str_msg_font_lobby_class_head : str_msg_font_lobby_class_head = getWord("FAST_SELECT_LOBBY")   '"快速选择大会堂"

'課程類型下拉選單
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_sel_normal_option : str_msg_sel_normal_option = "一般课程/小班制课程" 'getWord("NORMAL_CLASS")     '"一般课程/小班(1-6人)课程"
Dim str_msg_sel_vip_option : str_msg_sel_vip_option = getWord("VIP_CLASS")              '"一般课程/1对1课程"
Dim str_msg_sel_lobby_option : str_msg_sel_lobby_option = getWord("LOBBY_CLASS")        '"大会堂课程"

'一對一和一對多訂課 新增按鈕
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_btn_add_normal_class : str_msg_btn_add_normal_class = getWord("ADD")            '"新增"

'大會堂訂課 搜尋顧問按鈕
'include/ajax_lobby_class_search.asp
Dim str_msg_btn_search_lobby_class : str_msg_btn_search_lobby_class = getWord("CONFIRM")    '"确认"

'大會堂開課資訊搜尋 選擇主題的下拉選單
'include/ajax_lobby_class_search.asp
Dim str_msg_sel_lobby_topic : str_msg_sel_lobby_topic = getWord("SELECT_LOBBY_TOPIC")       '"请选择主题系列"

'大會堂開課資訊搜尋 選擇顧問的下拉選單
'include/ajax_lobby_class_search.asp
Dim str_msg_sel_lobby_consultant : str_msg_sel_lobby_consultant = getWord("SELECT_LOBBY_CONSULTANT") '"请选择顾问"

'大會堂開課資訊搜尋 選擇適用程度的下拉選單
'include/ajax_lobby_class_search.asp
Dim str_msg_sel_suit_level : str_msg_sel_suit_level = getWord("SUIT_LEVEL")     '"适用程度"

'大會堂開課資訊搜尋 選擇課程時段的下拉選單
'include/ajax_lobby_class_search.asp
Dim str_msg_sel_lobby_time : str_msg_sel_lobby_time = getWord("SELECT_LOBBY_TIME")  '"请选择课程時段"

'include/ajax_lobby_class_info.asp
Dim str_msg_font_lobby_class_info_head : str_msg_font_lobby_class_info_head = getWord("ALL_LOBBY")  '"所有大会堂"

'include/ajax_lobby_class_info.asp
Dim str_msg_font_lobby_class_list_head_1 : str_msg_font_lobby_class_list_head_1 = getWord("DATE_TYPE")  '"日期 / 类型"

'include/ajax_lobby_class_info.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_lobby_class_list_head_2 : str_msg_font_lobby_class_list_head_2 = getWord("CLASS_19")  '"主题"

'include/ajax_lobby_class_info.asp
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_lobby_class_list_head_3 : str_msg_font_lobby_class_list_head_3 = getWord("SUBTRACT_SESSION_POINT") '"预计扣除课时数"

'include/ajax_lobby_class_info.asp
'include/ajax_commend_lobby_class.asp
Dim str_msg_font_suit_level_eng : str_msg_font_suit_level_eng = getWord("SUIT_LV")   '"适用LV"

'取消課程會顯示的 "取消" 文字
'include/ajax_lobby_class_info.asp
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
Dim str_msg_a_cancel_class : str_msg_a_cancel_class = getWord("CANCEL") '"取消"

'include/ajax_lobby_class_info.asp
'include/ajax_commend_lobby_class.asp
Dim str_msg_font_cut_word : str_msg_font_cut_word =  getWord("SUBTRACT")    '"扣"

'include/ajax_lobby_class_info.asp
Dim str_msg_font_limit_full : str_msg_font_limit_full = getWord("FULL") '"已额满"

'include/ajax_lobby_class_info.asp
Dim str_msg_font_already_order_class : str_msg_font_already_order_class = getWord("YOU_ALREADY_ORDER")  '"您已预订"

'include/ajax_lobby_class_info.asp
Dim str_msg_font_search_no_data : str_msg_font_search_no_data = getWord("SEARCH_NO_DATA")   '"无相关大会堂课程，请重新查询。"


'日期
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_date : str_msg_font_date = getWord("CLASS_DATE")   '"日期"

'時間
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_time : str_msg_font_time = getWord("CLASS_TIME")   '"时间"


'顾问
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_consultant : str_msg_font_consultant = getWord("CLASS_22") '"顾问"

'一般课( 1-6人 )
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'watch_class.asp
Dim str_msg_font_normal_class : str_msg_font_normal_class = getWord("NORMAL_CLASS_NOTE")    '"一般课( 1-6人 )"

'一般课( 1對1 )
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'watch_class.asp
Dim str_msg_font_vip_class : str_msg_font_vip_class = getWord("VIP_CLASS_NOTE") '"一般课( 1對1 )"

'大会堂
'watch_class.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_lobby_class : str_msg_font_lobby_class = getWord("SPECIAL_SESSION")    '"大会堂"

'您此次预订课时共
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_order_result_1 : str_msg_font_order_result_1 = getWord("ORDER_RESULT_1")   '"您此次预订课时共"

'次
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_order_result_2 : str_msg_font_order_result_2 =  getWord("CONTRACT_8")  '"次"

'可用课时数
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'order_success.asp
'reorder_success.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_order_result_3 : str_msg_font_order_result_3 =  getWord("SESSION_POINT_FOR_USE")   '"可用课时数"

'共计扣除课时数
'order_success.asp
'reorder_success.asp
Dim str_msg_font_order_result_4 : str_msg_font_order_result_4 = getWord("TOTAL_SUBTRACT_SESSION")   '"共计扣除课时数"

'确认订課
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_btn_confirm_order : str_msg_btn_confirm_order = getWord("CONFIRM_ORDER_CLASS")  '"确认订課"

'您已成功预订
'order_success.asp
Dim str_msg_font_order_success_1 : str_msg_font_order_success_1 = getWord("YOU_ORDER_SUCCESS")  '"您已成功预订"

'課程共
'order_success.asp
'reorder_success.asp
Dim str_msg_font_order_success_2 : str_msg_font_order_success_2 = getWord("TOTAL_CLASS")    '"课程共"

'堂
'order_success.asp
'reorder_success.asp
Dim str_msg_font_order_success_3 : str_msg_font_order_success_3 = getWord("CLASS")  '"堂"

'您已成功重新预订
'reorder_success.asp
Dim str_msg_font_order_success_4 : str_msg_font_order_success_4 = getWord("YOU_REORDER_SUCCESS")    '"您已成功重新预订"

'返回预订课程
'order_success.asp
'reorder_success.asp
Dim str_msg_btn_return_to_order_class : str_msg_btn_return_to_order_class = getWord("RETURN_ORDER_CLASS")   '"返回预订课程"

'查看课表
'order_success.asp
'reorder_success.asp
Dim str_msg_btn_watch_class : str_msg_btn_watch_class = getWord("WATCH_CLASS")  '"查看课表"

'您也可以前往
'watch_class.asp
Dim str_msg_font_you_can_go : str_msg_font_you_can_go = getWord("YOU_CAN_GO_TO")    '"您也可以前往"

'预定课程
'watch_class.asp
Dim str_msg_font_order_class : str_msg_font_order_class = getWord("SHCEDULE_CLASS") '"预定课程"

'月历模式
'watch_class.asp
Dim str_msg_font_calender_mode : str_msg_font_calender_mode = getWord("MONTHLY_MODE")    '"月历模式"

'列表模式
'watch_class.asp
Dim str_msg_font_week_mode : str_msg_font_week_mode =  getWord("WEEKLY_MODE")   '"列表模式"

'备注：请将鼠标移至日期上即可观看详细内容
'watch_class.asp
Dim str_msg_font_calender_note : str_msg_font_calender_note = getWord("WATCH_CLASS_NOTE")    '"备注：请将鼠标移至日期上即可观看详细内容"

'年
'include/ajax_class_list_calender_head.asp
'include/ajax_class_list_week_head.asp
Dim str_msg_font_year : str_msg_font_year = getWord("YEAR") '"年"

'月
'include/ajax_class_list_calender_head.asp
'include/ajax_class_list_week_head.asp
Dim str_msg_font_month : str_msg_font_month = getWord("MONTH")   '"月"


'星期日
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
Dim str_msg_font_sunday : str_msg_font_sunday =  getWord("SUNDAY")  '"星期日"

'星期一
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
Dim str_msg_font_monday : str_msg_font_monday = getWord("MONDAY")   '"星期一"

'星期二
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
Dim str_msg_font_tuesday : str_msg_font_tuesday = getWord("TUESDAY")    '"星期二"

'星期三
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
    Dim str_msg_font_wednesday : str_msg_font_wednesday = getWord("WEDNESDAY")  '"星期三"

'星期四
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
Dim str_msg_font_thursday : str_msg_font_thursday = getWord("THURSDAY") '"星期四"

'星期五
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
Dim str_msg_font_friday : str_msg_font_friday = getWord("FRIDAY")   '"星期五"

'星期六
'include/ajax_class_list_calender.asp
'include/ajax_class_list_week.asp
Dim str_msg_font_saturday : str_msg_font_saturday = getWord("SATURDAY") '"星期六"

'推薦大會堂
'include/ajax_commend_lobby_class.asp
Dim str_msg_font_commend_lobby : str_msg_font_commend_lobby = getWord("COMMEND_LOBBY")   '"推荐大会堂"

'您目前尚未预订咨询课程
'normal_and_vip.asp
Dim str_msg_font_no_order_class : str_msg_font_no_order_class = getWord("LAST_NO_ORDER_CLASS")  '"您目前尚未预订咨询课程"

'该如何预订？
'normal_and_vip.asp
Dim str_msg_font_how_to_order : str_msg_font_how_to_order = getWord("HOW_TO_ORDER") '"该如何预订？"


'開課前n小時內無法取消
'normal_and_vip.asp
'normal_and_vip_near_order.asp
'lobby_near_order.asp
Dim str_msg_font_no_cancel_1 : str_msg_font_no_cancel_1 = getWord("CLASS_START")  '"开课前"
Dim str_msg_font_no_cancel_2 : str_msg_font_no_cancel_2 = getWord("NO_CANCEL")  '"小时内无法取消"
Dim str_msg_font_no_cancel_in_minute : str_msg_font_no_cancel_in_minute = getWord("NO_CANCEL_IN_MINUTE")  '"分钟内无法取消"

'/****************************************************************************************
'描述： Javascript alert 的訊息 javascript/reservationi_class.js 和 javascript/lobby_class.js
'參數描述 / 建立人員 / 建立日期：
'*****************************************************************************************/
'取消此課程 按鈕
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_btn_cancel_class : str_msg_btn_cancel_class = getWord("CANCEL_THIS_CLASS")  '"取消此课程"

'重订此课程 按鈕
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_btn_reorder_class : str_msg_btn_reorder_class = getWord("REORDER_THIS_CLASS")   '"重订此课程"

'按下 取消此課程按鈕後要顯示的 已取消 字樣
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_already_cancel_class : str_msg_font_already_cancel_class = getWord("ALREADY_CANCEL")   '"已取消"

'顯示在訂課列表上的文字 表示客戶訂的課程距離開時時間多久
'include/ajax_commend_lobby_class.asp
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_list_1 : str_msg_font_list_1 = getWord("SESSION_POINT_COUNT")   '"课时数"
Dim str_msg_font_list_before_24_hour : str_msg_font_list_before_24_hour = getWord("CLASS_TIME_BEFORE_24")   '"（开课前24小时）"
Dim str_msg_font_list_between_12_hour : str_msg_font_list_between_12_hour = getWord("CLASS_TIME_IN_12") '"（开课前24~12小时）"
Dim str_msg_font_list_between_4_hour : str_msg_font_list_between_4_hour = getWord("CLASS_TIME_IN_4")    '"（开课前12~4小时）"

'開課前四小時內無法取消
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_font_no_cancel_in_4_hour : str_msg_font_no_cancel_in_4_hour = getWord("NO_CANCEL_CLASS_IN_4_HOUR")  '"开课前四小时内无法取消"

'合約堂數不足或<=0
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_not_enough_session : str_msg_not_enough_session = getWord("CONTRACT_POINT_NOT_ENOUGH")  '"合约剩余课时数不足，请至「学习记录」查询详细资讯。"

'沒有選擇日期就按確認訂課 (一對一和一對多訂課頁面)
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_no_sel_class_date : str_msg_no_sel_class_date = getWord("SELECT_CLASS_DATE")    '"请选择上课日期"

'沒有選擇課程類型就按確認訂課 (一對一和一對多訂課頁面)
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_no_sel_class_type : str_msg_no_sel_class_type = getWord("SELECT_CLASS_TYPE")    '"请选择课程类型"

'沒有選擇諮詢時段就按確認訂課 (一對一和一對多訂課頁面)
'include/ajax_class_info.asp
'include/ajax_reservation_normal_and_vip_class.asp
Dim str_msg_no_sel_class_time : str_msg_no_sel_class_time = getWord("SELECT_CLASS_TIME")    '"请选择咨询时段"

'開課四小時內不允許訂課
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_class_no_open_in_4_hour : str_msg_class_no_open_in_4_hour =  getWord("CLASS_NO_OPEN_IN_4_HOUR") '"四小时前没有开课，请选择其他上课日期或咨询时段"

'24小時內臨時訂課僅開放一對多
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_class_no_open_in_24_hour : str_msg_class_no_open_in_24_hour = getWord("CLASS_NO_OPEN_IN_24_HOUR")   '"24小时内临时订课仅开放一对多，请选择其他上课日期或咨询时段"

'大會堂只允許開課前五分鐘可以訂課
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_class_no_open_in_5_min : str_msg_class_no_open_in_5_min = getWord("CLASS_NO_OPEN_IN_5_MINUTE")  '"大会堂只允许开课前五分钟可以订课，请选择其他上课日期或咨询时段"

'選到同一堂課 (已在頁面上選過了)
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_sel_same_class : str_msg_sel_same_class = getWord("SELECT_REPEAT_CLASS")    '"此堂课已在订课列表中"

'按下 重订此课程 按鈕後 執行失敗
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_reorder_class_fault : str_msg_reorder_class_fault = getWord("REORDER_CLASS_FAIL")  '"重订此课程失败"

'首次按下訂課時會顯示的提醒訊息
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
Dim str_msg_first_remind : str_msg_first_remind =  getWord("ORDER_CLASS_REMIND")    '"欢迎您预订VIPABC课程!\n提醒您，实际扣除堂数将以 ｢确认订课｣ 按钮送出的时间为准。"

'按下 近期預定課程頁面 or 近期预订大会堂课程 "確認訂課" 按鈕 且並無新增任何課程到訂課列表
Dim str_msg_no_add_class : str_msg_no_add_class = getWord("NO_ORDER_CLASS_NOW") '"您目前尚未预订课程"

'取消課程失敗
'javascript/watch_class.js
'javascript/lobby_class.js
Dim str_msg_cancel_class_fault : str_msg_cancel_class_fault = getWord("CANCEL_CLASS_FAIL")  '"取消课程失败"

'关闭
'include/ajax_commend_lobby_class.asp
Dim str_msg_close_commend_lobby : str_msg_close_commend_lobby = getWord("CLOSE")    '"关闭"

'打开
'include/ajax_commend_lobby_class.asp
Dim str_msg_open_commend_lobby : str_msg_open_commend_lobby = getWord("OPEN")       '"打开"

'合約已
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_contract_edate_invalid_0 : str_msg_contract_edate_invalid_0 = getWord("CONTRACT_EDATE_INVALID_0") '"合約已"

'合約于
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_contract_edate_invalid_1 : str_msg_contract_edate_invalid_1 = getWord("CONTRACT_EDATE_INVALID_1") '"合约于"

'到期
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_contract_edate_invalid_2 : str_msg_contract_edate_invalid_2 = getWord("CONTRACT_EDATE_INVALID_2")   '"到期"

'請至 ｢學習紀錄｣ 查詢詳細資訊
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_contract_edate_invalid_3 : str_msg_contract_edate_invalid_3 = getWord("CONTRACT_EDATE_INVALID_3")   '"请至 ｢学习纪录｣ 查询详细资讯。"

'合约剩余课时数不足
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_contract_edate_invalid_4 : str_msg_contract_edate_invalid_4 = getWord("CONTRACT_POINT_NOT_ENOUGH_1")   '"合约剩余课时数不足"


'请选择其他上课日期
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
Dim str_msg_select_other_class_date : str_msg_select_other_class_date = getWord("SELECT_OTHER_CLASS_DATE")   '"请选择其他上课日期"

'訂位失敗，請與客服人員連絡：02-3365-9999
'order_success.asp
'reorder_success.asp
'normal_and_vip_near_order.asp
'lobby_near_order.asp
Dim str_msg_reservation_class_fail_contact_cs : str_msg_reservation_class_fail_contact_cs = getWord("RESERVATION_CLASS_FAIL_CONTACT_CS") & CONST_SERVICE_PHONE

'此堂课您已经预定过了
'order_success.asp
'reorder_success.asp
'lobby_near_order.asp
Dim str_msg_this_class_already_order : str_msg_this_class_already_order = getWord("THIS_CLASS_YOU_ALREADY_ORDER")


'N小时前没有开课，请选择其他上课日期或咨询时段
'normal_and_vip_near_order.asp
'normal_and_vip_near_cancel.asp
'lobby_near_order.asp
'lobby_near_cancel.asp
Dim str_msg_class_no_open_in_hour : str_msg_class_no_open_in_hour =  getWord("CLASS_NO_OPEN_1") '"小时前没有开课"
Dim str_msg_class_no_open_2 : str_msg_class_no_open_2 =  getWord("SELECT_OTHER_CLASS_DATE_OR_TIME") '"请选择其他上课日期或咨询时段"
Dim str_msg_class_no_open_in_minute : str_msg_class_no_open_in_minute =  getWord("CLASS_NO_OPEN_IN_MINUTE") '"分钟前没有开课"

'您新增/取消 预约上课时间如下:
'resevationClass_New.js
'resevationClass_New.asp
Dim str_msg_order_class_as : str_msg_order_class_as =  "您新增/取消 预约上课时间如下"

'/****************************************************************************************
'描述： Javascript confirm 的訊息
'參數描述 / 建立人員 / 建立日期：
'CONST_CLASS_TYPE_NORMAL / RayChien / 20091129
'*****************************************************************************************/
'按下 取消課程的按鈕
'是否要取消此課程 ?
'javascript/lobby_class.js
'lobby_near_order.asp
'watch_class.asp
Dim str_msg_confirm_cancel_class : str_msg_confirm_cancel_class = getWord("CONFIRM_CANCEL_THIS_CLASS")  '"是否要取消此课程 ?"


%>
<script type="text/javascript">
<!--
    // reservation_class.js 要用到的data
    // lobby_class.js 要用到的data
    // watch_class.js 要用到的data
    var g_obj_class_data = {
        
        // 系統時間 YYYYMMDDHHMMSS
        str_now_sys_datetime : "<%=getFormatDateTime(now(), 8)%>",
        
        // 服務開始日期 YYYY/MM/DD
        str_service_start_date : "<%=getFormatDateTime(str_service_start_date, 5)%>",

        // 服務結束日期 YYYY/MM/DD
        str_service_end_date : "<%=getFormatDateTime(str_service_end_date, 5)%>",

        // 合約是否未過期 true:未到期  false:已過期
        bol_client_contract_edate_valid : "<%=bol_client_contract_edate_valid%>",

        // 取得客戶總剩餘堂數 = 大錢包+紅利包
        str_client_total_session : "<%=flt_client_total_session + flt_client_bonus_session%>",
        
        // 取得客戶當週剩餘的堂數
        str_client_week_session : "<%=flt_client_week_session%>",
        
        // 取得客戶本週的週期初開始日
        dat_now_week_account_start_date : "<%=dat_now_week_account_start_date%>",
        
        // 取得客戶本週的週期末結束日 (結算日)
        dat_now_week_account_end_date : "<%=dat_now_week_account_end_date%>",
        
        // 取得客戶購買合約的產品類型 1:一般產品 2:定時定額
        str_client_purchase_product_type : "<%=int_client_purchase_product_type%>",

        // 回傳最近一次預約First Session的日期和時間 YYYYMMDDHHMMSS
        str_last_first_session_datetime : "<%=getFormatDateTime(str_last_first_session_datetime, 8)%>",

        // 課程類型 一對一 OR 一對多 OR 大會堂 OR 一對三
        arr_class_type : {
            "normal" : "<%=Right("0"&CONST_CLASS_TYPE_NORMAL, 2)%>",
            "vip" : "<%=Right("0"&CONST_CLASS_TYPE_ONE_ON_ONE, 2)%>",
            "lobby" : "<%=Right("0"&CONST_CLASS_TYPE_LOBBY, 2)%>",
            "one_on_three" : "<%=Right("0"&CONST_CLASS_TYPE_ONE_ON_THREE, 2)%>"
        },

        // 課程類型 一對一 OR 大會堂 OR 一對三 幾小時內不能訂課 要顯示的訊息
        arr_msg_class_no_order_in_hour : {
            "vip" : "<%=arr_msg_no_order_in_hour(0)%>",
            "one_on_three" : "<%=arr_msg_no_order_in_hour(1)%>",
            "lobby" : "<%=arr_msg_no_order_in_hour(2)%>"
        },

        // 課程類型 一對一 OR 大會堂 OR 一對三 幾小時內不能取消 要顯示的訊息
        arr_msg_class_no_cancel_in_hour : {
            "vip" : "<%=arr_msg_no_cancel_in_hour(0)%>",
            "one_on_three" : "<%=arr_msg_no_cancel_in_hour(1)%>",
            "lobby" : "<%=arr_msg_no_cancel_in_hour(2)%>"
        },

        // 課程類型 一對一 OR 大會堂 OR 一對三 幾分鐘內不能取消 minute
        arr_class_no_cancel_in_minute : {
            "vip" : "<%=arr_no_cancel_in_minute(0)%>",
            "one_on_three" : "<%=arr_no_cancel_in_minute(1)%>",
            "lobby" : "<%=arr_no_cancel_in_minute(2)%>"
        },

        // alert 訊息
        arr_msg : {
            "normal_remind" : "<%=str_msg_first_remind%>",                          // 若是首次按下則顯示提醒訊息
            "no_sel_class_date" : "<%=str_msg_no_sel_class_date%>",                 // 沒有選擇日期就按確認訂課 (一對一和一對多訂課頁面)
            "no_sel_class_type" : "<%=str_msg_no_sel_class_type%>",                 // 沒有選擇課程類型就按確認訂課 (一對一和一對多訂課頁面)
            "no_sel_class_time" : "<%=str_msg_no_sel_class_time%>",                 // 沒有選擇諮詢時段就按確認訂課 (一對一和一對多訂課頁面)
            "class_no_open_in_4_hour" : "<%=str_msg_class_no_open_in_4_hour%>",     // 開課四小時內不允許訂課
            "class_no_open_in_24_hour" : "<%=str_msg_class_no_open_in_24_hour%>",   // 24小時內臨時訂課僅開放一對多
            "class_no_open_in_5_min" : "<%=str_msg_class_no_open_in_5_min%>",       // 大會堂只允許開課前五分鐘可以訂課
            "sel_same_class" : "<%=str_msg_sel_same_class%>",                       // 選到同一堂課 (已在頁面上選過了)
            "reorder_class_fault" : "<%=str_msg_reorder_class_fault%>",             // 按下 重订此课程 按鈕後 執行失敗
            "no_add_class" : "<%=str_msg_no_add_class%>",                           // 按下 近期預定課程頁面 or 近期预订大会堂课程 "確認訂課" 按鈕 且並無新增任何課程到訂課列表
            "confirm_cancel_class" : "<%=str_msg_confirm_cancel_class%>",           // 按下 取消課程的按鈕
            "cancel_class_fault" : "<%=str_msg_cancel_class_fault%>",               // 按下 取消課程
            "contract_edate_invalid_0" : "<%=str_msg_contract_edate_invalid_0%>",   // 合約已
            "contract_edate_invalid_1" : "<%=str_msg_contract_edate_invalid_1%>",   // 合約于
            "contract_edate_invalid_2" : "<%=str_msg_contract_edate_invalid_2%>",   // 到期
            "contract_edate_invalid_3" : "<%=str_msg_contract_edate_invalid_3%>",   // 請至 ｢學習紀錄｣ 查詢詳細資訊
            "contract_edate_invalid_4" : "<%=str_msg_contract_edate_invalid_4%>",   // 合约剩余课时数不足
            "select_other_class_date" : "<%=str_msg_select_other_class_date%>",     // 请选择其他上课日期
            "reservation_class_fail_contact_cs" : "<%=str_msg_reservation_class_fail_contact_cs%>",     // 訂位失敗，請與客服人員連絡：02-3365-9999
            "reservation_after_contract_start" : "<%=getWord("PLEASE_RESERVATION") & str_service_start_date & getWord("AFTER_CONTRACT_START")%>", // 請預訂YYYY/MM/DD(合約開始)之後的課程。
            "this_class_already_order" : "<%=str_msg_this_class_already_order%>",     // 此堂课您已经预定过了
            "class_no_open_in_hour" : "<%=str_msg_class_no_open_in_hour & "，" & str_msg_class_no_open_2%>",     // 小时前没有开课，请选择其他上课日期或咨询时段
            "close_commend_lobby" : "<%=str_msg_close_commend_lobby%>",             // 按下 推薦大會堂的 關閉 (HTML Element)
            "open_commend_lobby" : "<%=str_msg_open_commend_lobby%>",               // 按下 推薦大會堂的 打開 (HTML Element)
            "btn_cancel_class" : "<%=str_msg_btn_cancel_class%>",                   // 取消此課程 按鈕 (HTML Element)
            "btn_reorder_class" : "<%=str_msg_btn_reorder_class%>",                 // 重订此课程 按鈕 (HTML Element)
            "a_have_order_class" : "<%=str_msg_a_have_order_class%>",               // <a>此次预订課程</a> (HTML Element)
            "a_have_cancel_class" : "<%=str_msg_a_have_cancel_class%>",             // <a>近期取消課程</a> (HTML Element)
            "a_have_order_lobby_class" : "<%=str_msg_a_have_order_lobby_class%>",   // <a>近期预订大会堂课程</a> (HTML Element)
            "a_have_cancel_lobby_class" : "<%=str_msg_a_have_cancel_lobby_class%>", // <a>近期取消大会堂课程</a> (HTML Element)
            "a_no_have_order_class" : "<%=str_msg_a_no_have_order_class%>",         // <a>尚未预订谘询课程</a> (HTML Element)
            "a_no_have_cancel_class" : "<%=str_msg_a_no_have_cancel_class%>",       // <a>近期尚未取消課程</a> (HTML Element)
            "font_already_cancel_class" : "<%=str_msg_font_already_cancel_class%>", // 按下 取消此課程按鈕後要顯示的 已取消 字樣 (HTML Element)
            "font_list_1" : "<%=str_msg_font_list_1%>",                             // 顯示在訂課列表上的文字 (HTML Element)
            "font_list_before_24_hour" : "<%=str_msg_font_list_before_24_hour%>",   // 顯示在訂課列表上的文字 表示客戶訂的課程距離開時時間多久 (HTML Element)
            "font_list_between_12_hour" : "<%=str_msg_font_list_between_12_hour%>", // 顯示在訂課列表上的文字 表示客戶訂的課程距離開時時間多久 (HTML Element)
            "font_list_between_4_hour" : "<%=str_msg_font_list_between_4_hour%>",   // 顯示在訂課列表上的文字 表示客戶訂的課程距離開時時間多久 (HTML Element)
            "font_no_cancel_in_4_hour" : "<%=str_msg_font_no_cancel_in_4_hour%>",   // 開課前四小時內無法取消 (HTML Element)
            "font_normal_class" : "一般课程/小班制课程",                 // 一般课程/小班(1-3人)课程 (HTML Element)
            "font_vip_class" : "<%=str_msg_sel_vip_option%>",                       // 一般课程/1对1课程 (HTML Element)
            "font_lobby_class" : "<%=str_msg_sel_lobby_option%>",                   // 大会堂课程 (HTML Element)
            "font_one_on_three_class" : "一般课程/小班制课程",           // 一般课程/小班(1-3人)课程 (HTML Element)
            "font_no_cancel_1" : "<%=str_msg_font_no_cancel_1%>",                   // 开课前 (HTML Element)
            "font_no_cancel_2" : "<%=str_msg_font_no_cancel_2%>",                    // 小时内无法取消 (HTML Element)
			"before_confirm_order_class" : "<%=str_msg_order_class_as%>"					//您新增/取消 预约上课时间如下
        }
    };
//-->
</script>