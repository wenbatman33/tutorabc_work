<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/class/lobby/CourseUtility.asp" -->
<% '20120320 阿捨新增 VIPABCJR客戶判斷 bolVIPABCJR %>
<!--#include virtual="/program/member/getJRClient.asp"-->
<%
    dim bolDebugMode : bolDebugMode = false
    dim bolTestCase : bolTestCase = false

    '由於asp的星期一代碼為"2"...以此類推到7 而星期日為"1" 所以如果不是星期天則根據當天星期幾參數進行位移
    Dim arr_weekday_name : arr_weekday_name = Array("", "星期日", _
                                                    "星期一", "星期二", _
                                                    "星期三", "星期四", _
                                                    "星期五", "星期六")
    Dim int_db_connect : int_db_connect = CONST_TUTORABC_RW_CONN    'DB Connection Code

    'Hardcode Start hour reservation class time
    Dim intStartHour : intStartHour = 6

    '開放能預訂幾天內的課程
    Dim int_reservation_range_day : int_reservation_range_day = CONST_ALLOW_RESERVATION_RANGE_DAY

    '一對一 VIP 訂位顯示的圖片路徑'
    Dim str_img_src_vip : str_img_src_vip = "/images/reservation_class/one_on_one.gif"
	'Special 一對一 訂位顯示的圖片路徑'
	Dim str_img_src_vip_sp : str_img_src_vip_sp = "/images/reservation_class/sp_one_on_one.gif"
    '一對三 訂位顯示的圖片路徑'
    Dim str_img_src_three : str_img_src_three = "/images/reservation_class/one_on_three.gif"
    '一對四 訂位顯示的圖片路徑'
    Dim str_img_src_four : str_img_src_four = "/images/reservation_class/one_on_four.gif"
    '一對六 訂位顯示的圖片路徑'
    Dim str_img_src_six : str_img_src_six = "/images/reservation_class/one_on_three.gif"
    '大會堂 訂位顯示的圖片路徑'
    Dim str_img_src_lobby : str_img_src_lobby = "/images/reservation_class/lobby.gif"


    '一對一 VIP 訂位顯示的圖片路徑'
    'Dim str_img_src_vip : str_img_src_vip = "/images/reservation_class/reserved_vip.gif"
    '一對三 訂位顯示的圖片路徑'
    'Dim str_img_src_three : str_img_src_three = "/images/reservation_class/one_on_three.gif"
    '一對六 訂位顯示的圖片路徑'
    'Dim str_img_src_six : str_img_src_six = "/images/reservation_class/reserved.gif"
    '大會堂 訂位顯示的圖片路徑'
    'Dim str_img_src_lobby : str_img_src_lobby = "/images/reservation_class/ss_icon.gif"


    Dim str_now_year : str_now_year = Year(now)
    Dim arr_class_datetime_desc(6)                                      '每格checkbox代表的年月日/星期
    Dim arr_class_datetime(6)                                           '每格checkbox代表的年月日
    Dim str_class_time, str_class_datetime                              '每格checkbox代表的時間
    Dim arr_off_time(6)                                                 '人事部門規定之不開放上課之時間
    Dim arr_off_time_desc(6)                                            '人事部門規定之不開放上課的描述
    Dim dat_monday, dat_sunday                                          '星期一的日期 星期日的日期
    Dim str_tb_head : str_tb_head = ""                                  'TABLE 表頭 表尾
    Dim str_tb_head_background_css : str_tb_head_background_css = ""    'TABLE 表頭 表尾 css class
    Dim str_tb_head_today_css : str_tb_head_today_css = ""              'TABLE 表頭 表尾 欄位css class
    Dim arr_row_background_css : arr_row_background_css = Array("", "") 'Row 的css class, [0]=odd row, [1]=even row
    Dim str_no_order_chk_id : str_no_order_chk_id = ""                  '尚未訂過課的時段 checkbox id
    Dim str_order_chk_id : str_order_chk_id = ""                        '已訂過課的時段 checkbox id
    Dim str_order_sn_hdn_id : str_order_sn_hdn_id = ""                  '已訂過課的時段 hidden input id 儲存client_attend_list.sn
    Dim str_order_date_hdn_id : str_order_date_hdn_id = ""              '已訂過課的時段 hidden input id 儲存client_attend_list.attend_date
    Dim str_order_time_hdn_id : str_order_time_hdn_id = ""              '已訂過課的時段 hidden input id 儲存client_attend_list.attend_sestime
    Dim str_order_ctype_hdn_id : str_order_ctype_hdn_id = ""            '已訂過課的時段 hidden input id 儲存課程類型
    Dim str_no_order_chk_value : str_no_order_chk_value = ""            '尚未訂過課的時段 checkbox value
    Dim str_order_chk_value : str_order_chk_value = ""                  '已訂過課的時段 checkbox value
    Dim arr_booking_info : arr_booking_info = null                      '有訂課的資訊
    Dim bol_is_output_checkbox : bol_is_output_checkbox = true          '是否顯示可供訂課或已經訂課的checkbox控制項在頁面上
    Dim bol_allow_cancel : bol_allow_cancel = true                      '是否允許取消課程
    Dim str_not_allow_reason : str_not_allow_reason = ""                '不允許取消課程的原因
    Dim str_block_sylte : str_block_sylte = ""                          '不允許訂課時欄位的背景色
    Dim str_now_use_product_sn : str_now_use_product_sn = ""            '離目前時間最近的一筆 客戶合約的產品序號
    Dim int_no_order_in_minute : int_no_order_in_minute = 0             '幾分鐘內不能訂課 (定時定額產品規則)
    Dim int_no_cancel_in_minute : int_no_cancel_in_minute = 0           '幾分鐘內不能取消課程 (定時定額產品規則)
    Dim arr_product_session_rule : arr_product_session_rule = null      'product_session_rule
    Dim intOrderClassType : intOrderClassType = 0                       '已訂課的課程類型
    Dim strClassHtml : strClassHtml = ""

    Dim str_title : str_title = ""
    Dim dat_add_monday, str_style_class_week, str_style_class_date, i, j, k, arr_lobby_session_info, intArrLobyIndex
    Dim str_service_start_date, str_service_end_date, int_order_class_type, str_sql, strLobbySn, strLobbyDateTime, strImgTitle
    Dim int_diff_day, int_diff_minute, strBookingDate
    'client email
    Dim strClientEmail : strClientEmail = getSession("client_email", CONST_DECODE_NO)
    'debug
    Dim strDebug : strDebug = getRequest("strDebug", CONST_DECODE_NO)
    'strDebug = "yes"
	'CS12041502-VIPABC訂課規則修改 tree 20120515 
    Dim strProductNewOrOld : strProductNewOrOld = ""                    '取得產品是否為舊(1)還是新(2) ，VIPABC 舊產品客戶訂課時間為8小時前(1對1扣2堂) 新產品客戶訂課時間為24小時前(1對1扣3堂)
    '客戶編號 client_sn
    Dim str_client_sn : str_client_sn = getSession("client_sn", CONST_DECODE_NO)
    if (isEmptyOrNull(str_client_sn)) then
        str_client_sn = getRequest("client", CONST_DECODE_NO)
    end if
    
    if (Not isEmptyOrNull(str_client_sn)) then
        '客戶資料 COM 物件
        Dim obj_member_opt : Set obj_member_opt = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
        Dim int_member_result : int_member_result = obj_member_opt.prepareData(str_client_sn, int_db_connect)
        if (int_member_result = CONST_FUNC_EXE_SUCCESS) then
            'int_chk_client_type = obj_member_opt.getData("client_type")
            'int_chk_client_level = obj_member_opt.getData(CONST_CUSTOMER_NOW_LEVEL)             '取得客戶的等級
            str_service_start_date = session("str_service_start_date_cache_data")
            if(isEmptyOrNull(str_service_start_date)) then
                str_service_start_date = obj_member_opt.getData(CONST_CONTRACT_SERVICE_BEGIN_DATE)           '取得目前服務開始日
                session("str_service_start_date_cache_data") = str_service_start_date
            end if
            str_service_end_date = session("str_service_end_date_cache_data")
            if(isEmptyOrNull(str_service_end_date)) then
                str_service_end_date = obj_member_opt.getData(CONST_CONTRACT_SERVICE_END_DATE)               '取得目前服務結束日
                session("str_service_end_date_cache_data") = str_service_end_date
            end if
            'str_now_use_contract_sn = obj_member_opt.getData(CONST_CONTRACT_SN)                 '目前電子合約序號
            'bol_client_have_valid_contract = obj_member_opt.getData(CONST_CUSTOMER_IS_HAVE_CONTRACT)        '取得客戶是否有合約
            str_now_use_product_sn = session("str_now_use_product_sn_cache_data")
            if(isEmptyOrNull(str_now_use_product_sn)) then
                str_now_use_product_sn = obj_member_opt.getData(CONST_CONTRACT_PRODUCT_SN)          '離目前時間最近的一筆 客戶合約的產品序號
                session("str_now_use_product_sn_cache_data") = str_now_use_product_sn
            end if
            'str_purchase_account_sn = obj_member_opt.getData(CONST_LAST_CONTRACT_ACCOUNT_SN)    '離目前時間最近的一筆的於PURCHASE的序號
            if ( Not isEmptyOrNull(str_now_use_product_sn) ) then
                if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
                    session("unlimited") = "true" 
                    strUnlimited = session("unlimited")
                end if
            end if
        end if
        Set obj_member_opt = Nothing
    end if

	'CS12041502-VIPABC訂課規則修改 tree 20120515 
    strProductNewOrOld = getProductNewOrOld(str_now_use_product_sn) '取得產品是否為舊(1)還是新(2) 產品編號小於等於433=舊

    '設定星期一的日期
    dat_monday = Trim(Request("dat_monday"))
    if (isEmptyOrNull(dat_monday)) then
        '根據今天的日期，找出這周禮拜一的日期
        'dat_monday = getMondayOfWeek()
        dat_monday = date
    end if

    '設定星期日的日期
    dat_sunday = DateAdd("d", 6, dat_monday)

    '產品類型
    Dim strProductType : strProductType = getRequest("product_type", CONST_DECODE_NO)
    dim safeStrProductType : safeStrProductType = scstr(strProductType)

    'cfg_product.point_type
    Dim strProductPointType : strProductPointType = getRequest("product_point_type", CONST_DECODE_NO)


    if (cstr(CONST_PRODUCT_QUOTA) = cstr(strProductType)) then
        if (Not IsArray(arr_product_session_rule)) then
            '取得 產品對應課程類型的扣堂規則 (預設值)
            arr_product_session_rule = getProductSessionRule(str_now_use_product_sn, 0, CONST_VIPABC_RW_CONN)
        end if
    end if

    '**** 取得課程類型 START ****
    '課程類型 一對一 or 一對三 or 大會堂 or 一對六
    Dim str_class_type : str_class_type = getRequest("class_type", CONST_DECODE_NO)
    if (Not isEmptyOrNull(str_class_type)) then
        str_class_type = Right("0" & str_class_type, 2)

        '大會堂 找出有開課的
        if (cstr(CONST_CLASS_TYPE_LOBBY) = str_class_type) then
            arr_lobby_session_info = null
            str_sql = "SELECT sn, session_date, session_time FROM lobby_session WHERE (valid = 1) " &_
                      "AND (session_date BETWEEN @sdate AND @edate)"
            arr_lobby_session_info = excuteSqlStatementRead(str_sql, Array(dat_monday, dat_sunday), int_db_connect)          
        end if 'end of if (cstr(CONST_CLASS_TYPE_LOBBY) = str_class_type) then

    else
        '一般產品 & 無限卡 & 套餐產品
        if (sCStr(CONST_PRODUCT_NORMAL) = safeStrProductType or "3" = safeStrProductType or "4" = safeStrProductType or "5" = safeStrProductType ) then
            '預設1對1
            str_class_type = Right("0" & sCStr(CONST_CLASS_TYPE_ONE_ON_ONE), 2) 

        '定時定額產品
        elseif (cstr(CONST_PRODUCT_QUOTA) = cstr(strProductType)) then
            '取得 產品對應課程類型的扣堂規則 (預設值)
            'arr_product_session_rule = getProductSessionRule(str_now_use_product_sn, 0, CONST_VIPABC_RW_CONN)
            if (IsArray(arr_product_session_rule)) then
                str_class_type = Right("0" & cstr(arr_product_session_rule(3, 0)), 2)                
            end if
        end if 'end of 產品判斷
    end if 'end of 課程類型判斷
    '**** 取得課程類型 END ****

    dim int_class_type
    int_class_type = sCInt(str_class_type)

    '**** 取得 幾分鐘內不能訂課 和 幾分鐘內不能取消課程 START ****
    '定時定額產品
    if (cstr(CONST_PRODUCT_QUOTA) = cstr(strProductType)) then
        if (Not IsArray(arr_product_session_rule)) then
            '取得 產品對應課程類型的扣堂規則 (預設值)
            'arr_product_session_rule = getProductSessionRule(str_now_use_product_sn, 0, CONST_VIPABC_RW_CONN)
        end if

        if (IsArray(arr_product_session_rule)) then
            '取得 幾分鐘內不能訂課 和 幾分鐘內不能取消課程
            for i = 0 to Ubound(arr_product_session_rule, 2)
                if (not isEmptyOrNull(arr_product_session_rule(3, i))) then
                    if (int_class_type = Int(arr_product_session_rule(3, i))) then
                        int_no_order_in_minute = arr_product_session_rule(1, i) '幾分鐘內不能訂課 (定時定額產品規則)
                        int_no_cancel_in_minute = arr_product_session_rule(2, i)'幾分鐘內不能取消課程 (定時定額產品規則)
                    end if
                end if
            next
        end if
    end if
    '**** 取得 幾分鐘內不能訂課 和 幾分鐘內不能取消課程 END   ****

    '**** 取得產品類型 T:台灣TutorABC產品；C：陸網ABC產品 START ****
    Dim arrProductInfo, strProductWebsite
    strProductWebsite = ""
    arrProductInfo = getProductInfo(str_now_use_product_sn, CONST_VIPABC_RW_CONN)
    if (isarray(arrProductInfo)) then
        if (not isEmptyOrNull(arrProductInfo(11, 0))) then
            strProductWebsite = arrProductInfo(11, 0)
        end if
    end if
    '**** 取得產品類型 T:台灣TutorABC產品；C：陸網ABC產品 END ****

    '**** 取得總共有幾堂課 START ****
    Dim intClassCount
    intClassCount = getOpenClassCount(CONST_CLASS_TIME45_TYPEA)
    
    '**** 取得總共有幾堂課 END   ****

    '取得有訂課的資訊
   
    '20130318 阿捨新增 VIPABCJR客戶判斷
    if ( true = bolVIPABCJR ) then
        arr_booking_info = getCustomerBookingInfoJR(str_client_sn, dat_monday, dat_sunday, int_db_connect)
    else
        arr_booking_info = getCustomerBookingInfo(str_client_sn, dat_monday, dat_sunday, int_db_connect)
    end if

   '取得是Special 1 on 1課程的
	strSpecialSessionDateTime = "#"
	Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
	sql = "SELECT CONVERT(VARCHAR,SessionDate,112)+RIGHT('0'+CONVERT(VARCHAR,SessionTime),2) SessionDateTime FROM Special1on1Session WITH (NOLOCK) WHERE Valid=1 AND ClientSn=@ClientSn AND SessionDate BETWEEN @Sdate AND @Edate"
	arrParam = Array(str_client_sn,dat_monday,dat_sunday)
	intQueryResult = rs.excuteSqlStatementEach(sql, arrParam, CONST_TUTORABC_RW_CONN)
	if not rs.eof then
		for i=0 to (rs.RecordCount)-1
			strSpecialSessionDateTime = strSpecialSessionDateTime & rs("SessionDateTime") & "#"
			rs.MoveNext
		next
	end if
	
'**** 2012-09-13增加 因訂過課就不會排除預約測試之前的時間  Johnny
set course_ = new CourseUtility
re_valArr = course_.getFirstSessionTestTime ( rs , str_client_sn)
if "N"=re_valArr(0) then
	if (true = isDate(re_valArr(1))) then
		int_diff_orderEnvTest_minute = DateDiff("n", now(), re_valArr(1))
	end if
end if		
set course_ = nothing
'**** 2012-09-13 End
rs.close
Set rs = nothing

'str_attend_date_condition = " AND (client_attend_list.attend_date BETWEEN '" & dat_monday & "' AND '" & dat_sunday & "') "            
'Call setCalenderArrayData(str_attend_date_condition)    

'Function getOtherNewSession(ByVal p_intClientSn, ByVal p_strSdate, ByVal p_strEdate, ByVal p_intDeBugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述      ：取得此客戶訂的新課程類型 包括10mins大會堂(98), 20mins大會堂(97)
'傳入參數  ：
'[p_intClientSn:Integer] client_sn
'[p_strSdate:String] 開始日期
'[p_strEdate:String] 結束日期
'[p_intDeBugMode:Integer] 除錯模式 開:1 關:0
'[p_intConnectType:Integer] 連線字串
'回傳值    ：[objCfgSession] 課程資訊
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) : ArthurWu 2013/05/24 Created
'*****************************************************************************************/
dim objNewSessionInfo : objNewSessionInfo = null
dat_monday = getFormatDateTime(dat_monday, 5)
dat_sunday = getFormatDateTime(dat_sunday, 5)
set objNewSessionInfo = getOtherNewSession(str_client_sn, dat_monday, dat_sunday, 0, CONST_VIPABC_R_CONN)                
%>

<table cellpadding="0" cellspacing="0" class="tab">   
    <%
        '計算星期一~日的日期 START
        str_tb_head = "<tr>"
        str_tb_head = str_tb_head & "<th class=""t1"" nowrap>时间/日期</th>"
        for i = 0 to 6
            '預設的CSS Style
            str_style_class_week = "class="""""
            str_style_class_date = "class="""""

            '判斷是否是今日
            dat_add_monday = DateAdd("d", i, dat_monday)
            if (dat_add_monday = Date) then
                str_style_class_week = "class=""in"""
                str_style_class_date = "class=""in"""
            end if

            '輸出 Month / Day
            str_tb_head = str_tb_head & "<th " & str_style_class_week&" nowrap>" &_
                                        Month(dat_add_monday) & "/" & Day(dat_add_monday)

            '輸出 Week Name
            str_tb_head = str_tb_head & arr_weekday_name(DatePart("w", dat_add_monday)) & "</th>"

            '紀錄日期/星期
            arr_class_datetime_desc(i) = Month(dat_add_monday) & "/" & Day(dat_add_monday) & arr_weekday_name(DatePart("w", dat_add_monday))
                    
            '每格checkbox代表的年月日
            arr_class_datetime(i) = Year(dat_add_monday) & "/" & Month(dat_add_monday) & "/" & Day(dat_add_monday)

            '**** 檢查訂課日期是否為人事部門規定之不開放上課日期，取得時間及描述 START ****
            arr_off_time(i) = Array()

            Call getCompanyHolidayInfo(Year(dat_add_monday) & "/" & Month(dat_add_monday) & "/" & Day(dat_add_monday), _
                                      arr_off_time_desc(i), _
                                      arr_off_time(i), int_db_connect)
            '**** 檢查訂課日期是否為人事部門規定之不開放上課日期，取得時間及描述 END ****
        next
        str_tb_head = str_tb_head & "</tr>"
        '計算星期一~日的日期 END

        '[TABLE 表頭] 日期 START
        Response.Write str_tb_head
        '[TABLE 表頭]日期 END

        if ( isHardCodeComplied("vip_program_member_reservation_class_include_ajax_reservation_class_chose_time", str_client_sn) ) then
            intStartHour  = getClassBeginHour(0, CONST_CLASS_TIME45_TYPEA)
        end if

        '課程時間 START
        str_class_time = ""
        for i = intStartHour to intClassCount
            dim classBeginMinute
            classBeginMinute = getClassBeginMinute(i, CONST_CLASS_TIME45_TYPEA)

            '上課時間 HH:mm
            str_class_time = Right("0" & getClassBeginHour(i, CONST_CLASS_TIME45_TYPEA), 2) & ":" & classBeginMinute

            Response.Write "<tr>"
            '開課時間
            Response.Write "<td class=""t1"" >" & str_class_time & "</td>"
            
            dim objClassRecord : objClassRecord = null
            dim strShortSessionSTime : strShortSessionSTime = ""
            dim strShortSessionETime : strShortSessionETime = ""
            dim intShortSessionLength : intShortSessionLength = 10
            dim intShortSessionType : intShortSessionType = 98
            dim strClassDateETime : strClassDateETime = ""
            dim bolHaving : bolHaving = false
            dim bolShortSession : bolShortSession = false
            dim strClassWholeDateSTime : strClassWholeDateSTime = ""
            dim strClassWholeDateETime : strClassWholeDateETime = ""
            dim bolTimed : bolTimed = false
            dim arrClassInfo : arrClassInfo = null
            dim strSessionWord : strSessionWord = ""
            '可供訂課的欄位
            for j = 0 to 6
                int_diff_minute = 0
                int_diff_day = 0
                strLobbySn = ""
                strLobbyDateTime = ""
                strImgTitle = ""
                bol_is_output_checkbox = true
                bol_allow_cancel = true
                str_block_sylte = ""
                str_title = arr_class_datetime_desc(j)  'td 的 title
                strClassHtml = ""

                '上課時間 YYYY/MM/DD HH:mm
                str_class_datetime = arr_class_datetime(j) & " " & str_class_time

                int_diff_minute = DateDiff("n", now(), str_class_datetime)  '取得和系統時間相差的分
                int_diff_day = DateDiff("d", now(), str_class_datetime)  '取得和系統時間相差的天

                '**** 因為7x24在 2010/07/30 10:00 上線，但是只允許客戶訂8月1日以後24小時的課 raychien 20100729 START ****
                'Anfernee 2013/4/2 這段執行不到，直接註解掉
                'strBookingDate = str_class_datetime
                'if ((DateDiff("d", strBookingDate, "2010/08/01") > 0) And _
                    '(0 = i or 1 = i or 2 = i or 3 = i or 4 = i or 5 = i)) then
                    'bol_is_output_checkbox = false
                'end if
                '**** 因為7x24在 2010/07/30 10:00 上線，但是只允許客戶訂8月1日以後24小時的課 raychien 20100729 END   ****

                '**** 只允許預定幾天內的課程 START ****
                if (true = bol_is_output_checkbox) then
                    if (int_diff_day >= int_reservation_range_day) then
                        bol_allow_cancel = false
                        bol_is_output_checkbox = false
                        str_title = str_title & "(只开放预订"&int_reservation_range_day&"天内的课程)"
                    end if
                end if
                '**** 只允許預定幾天內的課程 END ****

                '**** 判斷該時段是否為人事部門規定之不開放上課之時間 START ****
                'arr_off_time(j) 存的是不開放的時間 0~6 代表星期一到日所代表的日期
                for k = 0 to Ubound(arr_off_time(j))
                    if (Not isEmptyOrNull(arr_off_time(j)(k))) then
                        if (CInt(arr_off_time(j)(k)) = i) then
                            bol_is_output_checkbox = false
                            '為何不開放的原因
                            str_title = str_title & "(" & arr_off_time_desc(j) & ")"
                            exit for
                        end if
                    end if
                next
                '**** 判斷該時段是否為人事部門規定之不開放上課之時間 END ****

                '**** 排除掉已超過上課時間的 START ****
                bolTimed = false
                if (true = bol_is_output_checkbox) then
                    if (compareCurrentTime(str_class_datetime, "n") <> CONST_TIME_AFTER) then
                        bol_is_output_checkbox = false
                        str_title = str_title & "(上课时间已过)"
                        bolTimed = true
                    end if
                end if
                '**** 排除掉已超過上課時間的 END ****

                '**** 排除掉合約的服務開始日之前的 START ****
                if (true = bol_is_output_checkbox) then
                    if (compareTime(str_class_datetime, str_service_start_date, "d") = CONST_TIME_BEFORE) then
                        bol_is_output_checkbox = false
                        str_title = str_title & "(合约尚未开始)"
                    end if
                end if 
                '**** 排除掉合約的服務開始日之前的 END ****

                '**** 排除掉合約的服務結束日之後的 START ****
                if (true = bol_is_output_checkbox) then
                    if (compareTime(str_class_datetime, str_service_end_date, "d") = CONST_TIME_AFTER) then
                        bol_is_output_checkbox = false
                        str_title = str_title &  "(合约已到期)"
                    end if
                end if 
                '**** 排除掉合約的服務結束日之後的 END ****

                '**** 排除掉大會堂未開課的時段 START ****
                if (true = bol_is_output_checkbox) then
                    if (cstr(CONST_CLASS_TYPE_LOBBY) = str_class_type) then
                        bol_is_output_checkbox = false
                        if (isarray(arr_lobby_session_info)) then
                            if (Ubound(arr_lobby_session_info) >= 0) then
                                for intArrLobyIndex = 0 to Ubound(arr_lobby_session_info, 2)
                                    strLobbyDateTime = arr_lobby_session_info(1, intArrLobyIndex) & " " &_
                                                       arr_lobby_session_info(2, intArrLobyIndex) & ":" & classBeginMinute
                                    if (compareTime(str_class_datetime, strLobbyDateTime, "n") = CONST_TIME_EQUAL) then
                                        'lobby_session.sn
                                        strLobbySn = arr_lobby_session_info(0, intArrLobyIndex)
                                        bol_is_output_checkbox = true
                                        exit for
                                    end if
                                next
                            end if
                        end if
                    end if  'end of if (cstr(CONST_CLASS_TYPE_LOBBY) = str_class_type) then
                end if
                '**** 排除掉大會堂未開課的時段 END   ****

                '**** 排除掉不能訂課的時段 START ****
                if (true = bol_is_output_checkbox) then
                    '一般產品 & 無限卡 & 套餐產品
                    if (sCStr(CONST_PRODUCT_NORMAL) = safeStrProductType or "3" = safeStrProductType or "4" = safeStrProductType or "5" = safeStrProductType ) then

                        '若是大陸一般產品，一對一 和 一對三 舊產品是八小時內不開放訂課 新產品是24小時內不開放訂課
                        if ("C" = Ucase(strProductWebsite)) then
							'CS12041502-VIPABC訂課規則修改 tree 20120515 --開始--
                            '如果是舊產品訂課規則不可更改 舊合約客戶權益不可以受損 20120427 tree
                            if ("1" = strProductNewOrOld) then
                                if(CONST_CLASS_TYPE_LOBBY <> int_class_type And _
                                   int_diff_minute <= 480) then

                                    bol_is_output_checkbox = false
                                    str_title = str_title & "(不开放8小时内预定课程)"
                                end if
                            '如果是新產品訂課規則更改為24小時內不可訂課 20120427
                            else
                                if(CONST_CLASS_TYPE_LOBBY <> int_class_type And _
                                   int_diff_minute <= 1440) then

                                    bol_is_output_checkbox = false
                                    str_title = str_title & "(不开放24小时内预定课程)"
                                end if
                            end if
							'CS12041502-VIPABC訂課規則修改 tree 20120515 --結束--
                        '台灣產品
                        else
                            '1對1不開放24小時內訂課
                            if (CONST_CLASS_TYPE_ONE_ON_ONE = int_class_type And _
                                int_diff_minute <= 1440) then

                                bol_is_output_checkbox = false
                                str_title = str_title & "(1对1不开放24小时内预定课程)"
       
                            '1對6和1對3僅不開放4小時內訂課
                            elseif ((CONST_CLASS_TYPE_NORMAL = int_class_type or _
                                     CONST_CLASS_TYPE_ONE_ON_THREE = int_class_type or _
                                     CONST_CLASS_TYPE_ONE_ON_FOUR = int_class_type) And _
                                     int_diff_minute <= 240) then
          
                                bol_is_output_checkbox = false
                                'str_title = str_title & "(1对3仅不开放4小时内预定课程)"
                                str_title = str_title & "(小班制仅不开放4小时内预定课程)"
                            '大會堂只允許開課前五分鐘可以訂課
                            elseif (CONST_CLASS_TYPE_LOBBY = int_class_type And _
                                int_diff_minute <= 5) then
          
                                bol_is_output_checkbox = false
                                str_title = str_title & "(开课前五分钟无法预定课程)"
                            end if
                        end if

                    '定時定額產品
                    elseif (cstr(CONST_PRODUCT_QUOTA) = cstr(strProductType)) then
                        if (int_no_order_in_minute <> 0 And _
                            int_diff_minute <= int_no_order_in_minute) then

                            bol_is_output_checkbox = false
                            str_title = str_title & "(开课前" & (int_no_order_in_minute\60) & "小时无法预定课程)"
                        end if
                    end if 'end of 產品判斷
                end if
				'****  2012-09-14增加 排除預約測試之前的時間  Johnny
				if "N"=re_valArr(0) then
					if (true = isDate(re_valArr(1))) then
						if int_diff_minute < int_diff_orderEnvTest_minute +30 then  '+30分鐘後測試時間
							bol_is_output_checkbox = false
						end if
					end if
				end if				
				'****  2012-09-14 End
                '**** 排除掉不能訂課的時段 END   ****
             
                '尚未訂過課的時段 checkbox id
                str_no_order_chk_id = "chk_no_order_time_" & i & j
                '已訂過課的時段 checkbox id 
                str_order_chk_id = "chk_order_time_" & i & j
                '已訂過課的時段 hiddend input id 儲存client_attend_list.sn
                str_order_sn_hdn_id = "hdn_order_time_sn_" & i & j
                '已訂過課的時段 hiddend input id 儲存client_attend_list.attend_date
                str_order_date_hdn_id = "hdn_order_time_date_" & i & j
                '已訂過課的時段 hiddend input id 儲存client_attend_list.attend_sestime
                str_order_time_hdn_id = "hdn_order_time_time_" & i & j
                '已訂過課的時段 hiddend input id 儲存client_attend_list.attend_sestime課程類型
                str_order_ctype_hdn_id = "hdn_order_time_ctype_" & i & j


                '尚未訂過課的時段 checkbox value
                str_no_order_chk_value = str_class_type & getFormatDateTime(str_class_datetime, 11) & strLobbySn

                '設定td 的 background color'
                '不能訂課(區塊呈現灰色)
               if ( false = bol_is_output_checkbox ) then
                    str_block_sylte = "style=""background-color:Gray"""
                '臨時預約24小時之內課程（區塊內呈現橘色，名額有限）tree
                elseif (int_diff_minute <= 1440) then
                    '新產品才顯示黃色和訊息 '舊產品應該是8小時候隨時可以訂課 顯示客戶會有抱怨
                    if ( "2" = strProductNewOrOld ) then
                        str_block_sylte = "style=""background-color:#ff9728;"""
                        str_title = "临时预约24小时之内课程（名额有限）"
                     end if
                end if

                bolHaving = false
                strSessionWord = ""
                bolShortSession = false
                if ( not objNewSessionInfo.eof ) then
                    while not objNewSessionInfo.eof
                        ' : 30~ : 15
                        strShortSessionSTime = objNewSessionInfo("SessionStartTime")
                        strShortSessionETime = objNewSessionInfo("SessionEndTime")
                        strShortSessionSTime = getFormatDateTime(strShortSessionSTime, 1)
                        strShortSessionETime = getFormatDateTime(strShortSessionETime, 1)
                        intNewSessionType = objNewSessionInfo("attend_livesession_types")
                        if ( 98 = sCint(intNewSessionType) ) then
                            strSessionWord = "10mins大会堂"
                        elseif ( 97 = sCint(intNewSessionType) ) then
                            strSessionWord = "20mins大会堂"
                        end if

                        if ( true = bolDebugMode ) then
                            response.write "intNewSessionType : " & intNewSessionType &"<br/>"
                            response.write "strShortSessionSTime : " & strShortSessionSTime &"<br/>"
                            response.write "strShortSessionETime : " & strShortSessionETime &"<br/>"
                        end if

                        str_class_datetime = getFormatDateTime(str_class_datetime, 1) '1 = YYYY/MM/DD HH:MM    (HH 為24小時制)
                        strClassDateETime = DateAdd("n", 45, str_class_datetime)
                        strClassDateETime = getFormatDateTime(strClassDateETime, 1) '1 = YYYY/MM/DD HH:MM    (HH 為24小時制)
                        'Function isDateTimeMixed(ByVal p_strStartTime1, ByVal p_strEndTime1, ByVal p_strStartTime2, ByVal p_strEndTime2, ByVal p_intDebugMode)
                        if ( true = isDateTimeMixed(str_class_datetime, strClassDateETime, strShortSessionSTime, strShortSessionETime, 0) ) then
                             bol_is_output_checkbox = false
                            if ( true = bolTimed ) then '課程時間已過, 不改變短課程預定背景顏色 不秀出預訂課程資訊
                            else
                                str_title = str_title & "&#10;" & "(已预定 "& strShortSessionSTime & " ~ " & strShortSessionETime &" [" & strSessionWord & "]!)"
                                bolShortSession = true
                            end if
							bolHaving = true
                        end if

                        ' : 00~ : 59
                        strClassWholeDateSTime = arr_class_datetime(j) & " " & Right("0" & getClassBeginHour(i, CONST_CLASS_TIME45_TYPEA), 2) & ":00"
                        strClassWholeDateSTime = getFormatDateTime(strClassWholeDateSTime, 1)
                        strClassWholeDateETime = DateAdd("n", 59, strClassWholeDateSTime)
                        strClassWholeDateETime = getFormatDateTime(strClassWholeDateETime, 1)
                        if ( true = isDateTimeMixed(strClassWholeDateSTime, strClassWholeDateETime, strShortSessionSTime, strShortSessionETime, 0) ) then
                            if ( true = bolHaving OR true = bolTimed ) then '課程時間已過, 或實際課程區間已找出 不改變短課程預定背景顏色 不秀出預訂課程資訊
                            else
                                str_title = str_title & "&#10;" & "(已预定 "& strShortSessionSTime & " ~ " & strShortSessionETime &" [" & strSessionWord & "]!)"
                                bolShortSession = true
                            end if
                        end if
                    objNewSessionInfo.movenext
                    wend
                    objNewSessionInfo.movefirst
                end if

                if ( true = bolShortSession ) then
                    str_block_sylte = "style=""background-color:#ff9728;"""
                end if

                '**** 判斷該時段是否已經訂課 START ****
                if (IsArray(arr_booking_info)) then

                    '**** 判斷該時段是否允許取消已訂課的課程 START ****
                    '一般產品 & 無限卡 & 套餐產品
                    if (sCStr(CONST_PRODUCT_NORMAL) = safeStrProductType or "3" = safeStrProductType or "4" = safeStrProductType or "5" = safeStrProductType ) then
                        if (int_diff_minute <= 240) then
                            bol_allow_cancel = false
                            strImgTitle = "(开课前四小时无法取消课程)"
                        end if
                    '定時定額產品
                    elseif (cstr(CONST_PRODUCT_QUOTA) = cstr(strProductType)) then
                        if (int_no_cancel_in_minute <> 0 And _
                            int_diff_minute <= int_no_cancel_in_minute) then
                            bol_allow_cancel = false
                            strImgTitle = "(开课前" & (int_no_cancel_in_minute\60) & "小时无法取消课程)"
                        end if
                    end if
                    '**** 判斷該時段是否允許取消已訂課的課程 END   ****
                    
                    for k = 0 to ubound(arr_booking_info, 2)
                        if (compareTime(str_class_datetime, getFormatDateTime(arr_booking_info(1, k), 5) & " " & arr_booking_info(2, k) & ":" & classBeginMinute, "n") = CONST_TIME_EQUAL) then
                            str_block_sylte = ""
                            intOrderClassType = arr_booking_info(3, k)
                            if ( true = bolDebugMode ) then
                                response.Write "intOrderClassType: " & intOrderClassType & "<br/>"
                            end if
                            '輸出課程類型對應的圖片'
                            '大會堂
                            if ( sCInt(arr_booking_info(10, k)) > 0 and intOrderClassType = 4 ) then
                                int_order_class_type = CONST_CLASS_TYPE_LOBBY
                                strClassHtml = "<img src=""" & str_img_src_lobby & """ alt=""大会堂"&strImgTitle&""" />"
                            '一對一
                            elseif (cint(intOrderClassType) = CONST_CLASS_TYPE_ONE_ON_ONE or _
                                    cint(intOrderClassType) = 1) then
                                int_order_class_type = CONST_CLASS_TYPE_ONE_ON_ONE
								if instr(strSpecialSessionDateTime,"#"&getFormatDateTime(arr_booking_info(1, k), 7)&right("0"&arr_booking_info(2, k),2)&"#")>0 then
									strClassHtml = "<img src=""" & str_img_src_vip_sp & """  alt=""Special 1对1"&strImgTitle&"""/>"
								else
									strClassHtml = "<img src=""" & str_img_src_vip & """  alt=""1对1"&strImgTitle&"""/>"
								end if
                            '一對三
                            elseif (sCInt(intOrderClassType) = CONST_CLASS_TYPE_ONE_ON_THREE) then
                                int_order_class_type = CONST_CLASS_TYPE_ONE_ON_THREE
                                'strClassHtml = "<img src=""" & str_img_src_three & """  alt=""1对3"&strImgTitle&"""/>"
                                strClassHtml = "<img src=""" & str_img_src_three & """  alt=""小班制"&strImgTitle&"""/>"
                            elseif (sCInt(intOrderClassType) = CONST_CLASS_TYPE_ONE_ON_FOUR) then
                                int_order_class_type = CONST_CLASS_TYPE_ONE_ON_FOUR
                                'strClassHtml = "<img src=""" & str_img_src_three & """  alt=""1对3"&strImgTitle&"""/>"
                                strClassHtml = "<img src=""" & str_img_src_four & """  alt=""小班制"&strImgTitle&"""/>"
                            '一對六
                            elseif (cint(intOrderClassType) = CONST_CLASS_TYPE_NORMAL) then
                                '定時定額產品
                                'TODO: 目前定時定額產品一對多課程當成是一對三
                                if (cstr(CONST_PRODUCT_QUOTA) = cstr(strProductType)) then
                                    int_order_class_type = CONST_CLASS_TYPE_ONE_ON_THREE
                                    strClassHtml = "<img src=""" & str_img_src_three & """  alt=""1对3"&strImgTitle&"""/>"
                                else
                                    int_order_class_type = CONST_CLASS_TYPE_NORMAL
                                    strClassHtml = "<img src=""" & str_img_src_six & """  alt=""1对3"&strImgTitle&"""/>"
                                end if
                            end if

                            if ( true = bol_allow_cancel ) then
                                str_title = arr_class_datetime_desc(j)  'td 的 title
                                '輸出Checkbox
                                strClassHtml = strClassHtml & "<input type=""checkbox"" checked=""checked"" id=""" & str_order_chk_id & """ " & _
                                            " onclick=""onclickAlreadyOrderTime('"&str_order_chk_id&"', '"&str_order_sn_hdn_id&"', " & _
                                            "  '"&str_order_date_hdn_id&"', '"&str_order_time_hdn_id&"', '"&str_order_ctype_hdn_id&"', " & _
                                            "  '"&arr_booking_info(0, k)&"', '"&getFormatDateTime(arr_booking_info(1, k), 5)&"', " & _
                                            "  '"&arr_booking_info(2, k)&"', '"&int_order_class_type&"');""/>" & _
                                            "<input type=""hidden"" id=""" & str_order_sn_hdn_id & """ name=""total_cancel_class_sn"" value="""" />" & _
                                            "<input type=""hidden"" id=""" & str_order_date_hdn_id & """ name=""total_cancel_class_date"" value="""" />" & _
                                            "<input type=""hidden"" id=""" & str_order_time_hdn_id & """ name=""total_cancel_class_time"" value="""" />" & _
                                            "<input type=""hidden"" id=""" & str_order_ctype_hdn_id & """ name=""total_cancel_class_type"" value="""" />"
                            end if
                            bol_is_output_checkbox = false
                        end if
                    next
                end if
                '**** 判斷該時段是否已經訂課 END  ****

                '該時段尚未訂課
                if (true = bol_is_output_checkbox) then
                    '輸出Checkbox
                    strClassHtml = "<input type=""checkbox"" name=""hdn_total_reservation_class_dt"" " & _
                                " onclick=""orderCheckDialogueBox('"& str_no_order_chk_value &"','"&str_no_order_chk_id&"','"&strClientEmail&"','"&strDebug&"')""    " & _
                                " id=""" & str_no_order_chk_id & """  value=""" & str_no_order_chk_value & """ " & _
                                " />"
                end if

                Response.Write "<td title=""" & str_title & """ " & str_block_sylte & " >"
                Response.Write strClassHtml
                Response.Write "&nbsp;</td>"
            next

            Response.Write "</tr>"
        next
        '課程時間 END

        '[TABLE 表尾]日期 START
        Response.Write str_tb_head
        '[TABLE 表尾]日期 END
    %>
</table>
<input type="hidden" name="special1on1value" value="<%=strSpecialSessionDateTime%>"/>
<!--<input type="hidden" id="hdn_already_reservation" value="" />-->
<!--<input type="hidden" id="hdn_already_cancle" value="" />-->

<script type="text/javascript" language="javascript">

    /// <summary>
    /// 點擊已訂課的checkbox
    /// </summary>
    /// <param name="p_chk_id">checkbox id</param>
    /// <param name="p_hdn_sn_id">input hidden id</param>
    /// <param name="p_hdn_date_id">input hidden id</param>
    /// <param name="p_hdn_time_id">input hidden id</param>
    /// <param name="p_hdn_ctype_id">input hidden id</param>
    /// <param name="p_booking_sn">booking sn</param>
    /// <param name="p_booking_date">booking date(YYYY/MM/DD)</param>
    /// <param name="p_booking_time">booking time(24H)</param>
    /// <return>no return</return>
    /// <remarks>[RayChien] 2009/10/29 Created</remarks>
    function onclickAlreadyOrderTime(p_chk_id, p_hdn_sn_id, p_hdn_date_id,
                                     p_hdn_time_id, p_hdn_ctype_id, p_booking_sn,
                                     p_booking_date, p_booking_time, p_booking_type)
    {
        //var obj_chk = document.getElementById(p_chk_id);
        //(obj_chk.checked) ? obj_chk.checked = false : obj_chk.checked = true;
        
        var obj_chk = document.getElementById(p_chk_id);
        var obj_hdn_sn = document.getElementById(p_hdn_sn_id);
        var obj_hdn_date = document.getElementById(p_hdn_date_id);
        var obj_hdn_time = document.getElementById(p_hdn_time_id);
        var obj_hdn_ctype = document.getElementById(p_hdn_ctype_id);

        // 設定欲取消課程的相關資料
        if (false == obj_chk.checked)
        {
            obj_hdn_sn.value = p_booking_sn;
            obj_hdn_date.value = p_booking_date;
            obj_hdn_time.value = p_booking_time;
            obj_hdn_ctype.value = p_booking_type;
        }
        else
        {
            obj_hdn_sn.value = "";
            obj_hdn_date.value = "";
            obj_hdn_time.value = "";
            obj_hdn_ctype.value = "";
        }
        // 紀錄已經在頁面上點擊過可取消的課程時段
        //document.getElementById("hdn_already_cancle").value = "yes";
    }

    /// <summary>
    /// 紀錄已經在頁面上點擊過可預訂的課程時段
    /// </summary>
    /// <return>no return</return>
    /// <remarks>[RayChien] 2010/07/07 Created</remarks>
    //function recordAlreadyReservation()
    //{
        //document.getElementById("hdn_already_reservation").value = "yes";
    //}
</script>
<!--#include virtual="/program/member/reservation_class/include/orderCheckLobbySession.asp" -->