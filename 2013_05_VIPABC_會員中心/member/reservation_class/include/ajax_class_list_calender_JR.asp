<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<%
    sub setCalenderArrayData(p_str_attend_date_condition, p_int_type)
    '/****************************************************************************************
    '描述      ：設定要顯示的日期及訂課資訊
    '傳入參數  ：[p_str_sql:String] client_attend_list.attend_date where 的條件
    '            [p_int_type:String] 1:設定上個月 2:設定本月 3:設定下個月
    '回傳值    ：
    '牽涉數據表： 
    '備註      ：
    '歷程(作者/日期/原因) ：[Ray Chien] 2009/12/21 Created
    '*****************************************************************************************/
        Dim arr_attend_list_data, str_title, str_ball_style_class, str_sql, str_suit_lev, str_arr_tmp
        Dim str_cancel_class_datetime, str_lobby_session_topic

        'TODO: 大會堂的標題 判斷語系
        if (Session("language") = CONST_LANG_CN) then
            str_lobby_session_topic = " ISNULL(lobby_session.topic_cn, '') AS topic "
        else
            str_lobby_session_topic = " ISNULL(lobby_session.topic, '') AS topic "
        end if
		
        dim arrClassTypeVIPJR : arrClassTypeVIPJR = split(CONST_CLASS_TYPE_VIPJR,",") '1, 1, 4, 5, 6, 2, 3
	    dim arrClassTypeVIPABC : arrClassTypeVIPABC = split(CONST_CLASS_TYPE_VIPABC,",") '1, 7, 2, 3, 14, 4, 99
        dim intI : intI = 0
		'Special 1 on 1 課程查詢
		'special1on1SessionList = "#"
		's_session_date = replace(p_str_attend_date_condition,"client_attend_list.attend_date","Special1on1Session.SessionDate")
		'Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
		'sql = "SELECT CONVERT(VARCHAR(8), SessionDate, 112) + RIGHT('0'+ CONVERT(VARCHAR(2), SessionTime),2) SessionDateTime "
		'sql = sql & "FROM    Special1on1Session WITH ( NOLOCK ) "
		'sql = sql & "WHERE   Valid = 1 "
		'sql = sql & "        AND ClientSn = "&g_var_client_sn&" "
		'sql = sql & "        AND ClientPlatform = 2 "
		'sql = sql & s_session_date
		'intQueryResult = rs.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_RW_CONN)
		'if not rs.eof then
		'	for i=0 to rs.RecordCount-1
		'		special1on1SessionList = special1on1SessionList & rs("SessionDateTime")&"#"
		'		rs.MoveNext
		'	next
		'end if
		'rs.close
		'Set rs = nothing

        'str_sql = "SELECT CONVERT(varchar, client_attend_list.attend_date, 111) AS sdate, " &_
        '          "client_attend_list.attend_sestime, " &_
        '          "ISNULL(con_basic.basic_fname,'') + ' ' + ISNULL(con_basic.basic_lname,'') AS con_ename, " &_
        '          "client_attend_list.attend_livesession_types, " &_
        '          "ISNULL(client_attend_list.attend_level, 0) AS attend_level, " &_                          
        '          str_lobby_session_topic & ", " &_
        '          "ISNULL(lobby_session.lev, '') AS lobby_suit_level, " &_
        '          "client_attend_list.sn, " &_
		'  "client_attend_list.special_sn " &_
        '          "FROM client_attend_list " &_                    
        '          "LEFT JOIN lobby_session on client_attend_list.special_sn = lobby_session.sn " &_
        '          "LEFT JOIN con_basic on lobby_session.consultant = con_basic.con_sn " &_
        '          "WHERE (client_attend_list.client_sn = @client_sn) " &_
        '          "AND (client_attend_list.valid = 1) " & p_str_attend_date_condition & "  ORDER BY client_attend_list.attend_date DESC, client_attend_list.attend_sestime ASC  "
           
        str_sql = " SELECT CONVERT(VARCHAR, member.ClassRecord.ClassStartDate, 111) AS sdate, "
        str_sql = str_sql & " member.ClassRecord.ClassStartTime, "
        str_sql = str_sql & " ISNULL(con_basic.basic_fname,'') + ' ' + ISNULL(con_basic.basic_lname,'') AS con_ename, "
        str_sql = str_sql & " member.ClassRecord.ClassTypeId , "
        str_sql = str_sql & " ISNULL(member.ClassRecord.ProfileLearnLevel, 0) AS attend_level, "                   
        str_sql = str_sql & str_lobby_session_topic & ", "
        str_sql = str_sql & " ISNULL(lobby_session.lev, '') AS lobby_suit_level, "
        str_sql = str_sql & " member.ClassRecord.ClassRecordId, "
		str_sql = str_sql & " member.ClassRecord.LobbySessionId "
        str_sql = str_sql & " FROM member.ClassRecord "               
        str_sql = str_sql & " LEFT JOIN lobby_session ON member.ClassRecord.LobbySessionId = lobby_session.sn "
        str_sql = str_sql & " LEFT JOIN con_basic ON lobby_session.consultant = con_basic.con_sn "
        str_sql = str_sql & " WHERE ( member.ClassRecord.client_sn = @client_sn ) "
        str_sql = str_sql & " AND ( member.ClassRecord.ClassStatusId = 2 AND member.ClassRecord.ClassRecordValid = 1 ) " & p_str_attend_date_condition & "  ORDER BY member.ClassRecord.ClassStartDate DESC, member.ClassRecord.ClassStartTime  "
        arr_attend_list_data = excuteSqlStatementRead(str_sql, Array(g_var_client_sn), CONST_VIPABC_RW_CONN)
        'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
        if (isSelectQuerySuccess(arr_attend_list_data)) then
            if (Ubound(arr_attend_list_data) >= 0) then
                for i = 0 to Ubound(arr_attend_list_data, 2)
                    
                    '轉成新Table的值
	                for intI = 0 to Ubound(arrClassTypeVIPJR) 
	                    if ( sCInt(arr_attend_list_data(3, i)) = sCInt(arrClassTypeVIPJR(intI)) ) then
                            arr_attend_list_data(3, i) = arrClassTypeVIPABC(intI)
                            exit for
                        end if
	                next

                    '四小時內無法取消 不顯示四小時內的課程
                    str_cancel_class_datetime = arr_attend_list_data(0, i) & " " & Right("0"&arr_attend_list_data(1, i), 2) & ":30"
                    str_title = arr_attend_list_data(0, i) & chr(10)
                    str_title = str_title & arr_attend_list_data(1, i) & ":30" & chr(10)

                    '大會堂 ball_blue
                    'if (arr_attend_list_data(4, i) > 12) then
		            '20120905 阿捨新增 大會堂改用special_sn判斷
		            if (  0 < sCInt(arr_attend_list_data(8, i) )  ) then 
                        str_title = str_title & arr_attend_list_data(2, i) & chr(10)
                        str_title = str_title & arr_attend_list_data(5, i) & chr(10)

                        '大會堂適合的等級
                        str_suit_lev = Trim(arr_attend_list_data(6, i))
                        if (Not isEmptyOrNull(str_suit_lev)) then
			                if (Left(str_suit_lev, 1) = ",") then
                                str_suit_lev = Right(str_suit_lev, Len(str_suit_lev)-1)
                            end if
			                if (Right(str_suit_lev, 1) = ",") then 
                                str_suit_lev = Left(str_suit_lev, Len(str_suit_lev)-1)
                            end if
			                str_arr_tmp = Split(str_suit_lev, ",")
                            str_title = str_title & "LV" & str_arr_tmp(0) & "-" & str_arr_tmp(UBound(str_arr_tmp)) & chr(10)
                        end if

                        str_ball_style_class = "class=""ball_blue"""
                        str_title = str_title & str_msg_sel_lobby_option
                    '一對三 ball_red
                    '20120729 阿捨修正 在資料庫內一對多的課程類型為4 非3
                    elseif ( scint(arr_attend_list_data(3, i)) = 4 or scint(arr_attend_list_data(3, i)) = CONST_CLASS_TYPE_ONE_ON_FOUR ) then
                        str_ball_style_class = "class=""ball_red"""
                        str_title = str_title & str_msg_sel_normal_option
                    '一對一 ball_green
                    elseif (Int(arr_attend_list_data(3, i)) = CONST_CLASS_TYPE_ONE_ON_ONE) then
						if instr(special1on1SessionList,"#"&replace(arr_attend_list_data(0, i),"/","")&right("0"&arr_attend_list_data(1, i),2)&"#")>0 then
							str_ball_style_class = "class=""ball_purple"""
							str_title = str_title & replace(str_msg_sel_vip_option,"1对1","Special1对1")
						else
							str_ball_style_class = "class=""ball_green"""
							str_title = str_title & str_msg_sel_vip_option
						end if
                    end if

                    str_title = getReplaceInjectionWord(str_title)

                    'Prev
                    if (p_int_type = 1) then
                        if (DateDiff("n", now(), str_cancel_class_datetime) >= 240 and bol_client_have_valid_contract) then
                            arr_prev_td(Day(arr_attend_list_data(0, i))) = arr_prev_td(Day(arr_attend_list_data(0, i))) & "<p id=""cancel_class_info_"&arr_attend_list_data(7, i)&"""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span><a id=""cancel_class_link_"&arr_attend_list_data(7, i)&""" href=""javascript:void(0);"" onclick=""javascript:g_obj_reservation_class_ctrl.cancelClass('"&arr_attend_list_data(7, i)&"', '"&arr_attend_list_data(0, i)&"', '"&arr_attend_list_data(1, i)&"');"">"&str_msg_a_cancel_class&"</a><img id=""img_cancel_class_loading_"&arr_attend_list_data(7, i)&"""  width=""25"" height=""25"" src=""/lib/javascript/JQuery/images/ajax-loader.gif"" border=""0"" style=""display:none;""/></p>"
                        '四小時內無法取消 不顯示取消的按鈕 或 合約已到期
                        else
                            arr_prev_td(Day(arr_attend_list_data(0, i))) = arr_prev_td(Day(arr_attend_list_data(0, i))) & "<p id=""cancel_class_info_"&arr_attend_list_data(7, i)&"""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span></p>"
                        end if
                    'Now
                    elseif (p_int_type = 2) then
                        if (DateDiff("n", now(), str_cancel_class_datetime) >= 240 and bol_client_have_valid_contract) then
                            arr_now_td(Day(arr_attend_list_data(0, i))) = arr_now_td(Day(arr_attend_list_data(0, i))) & "<p id=""cancel_class_info_"&arr_attend_list_data(7, i)&"""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span><a id=""cancel_class_link_"&arr_attend_list_data(7, i)&""" href=""javascript:void(0);"" onclick=""javascript:g_obj_reservation_class_ctrl.cancelClass('"&arr_attend_list_data(7, i)&"', '"&arr_attend_list_data(0, i)&"', '"&arr_attend_list_data(1, i)&"');"">"&str_msg_a_cancel_class&"</a><img id=""img_cancel_class_loading_"&arr_attend_list_data(7, i)&"""  width=""25"" height=""25"" src=""/lib/javascript/JQuery/images/ajax-loader.gif"" border=""0"" style=""display:none;""/></p>"
                        '四小時內無法取消 不顯示取消的按鈕 或 合約已到期
                        else
                            arr_now_td(Day(arr_attend_list_data(0, i))) = arr_now_td(Day(arr_attend_list_data(0, i))) & "<p id=""cancel_class_info_"&arr_attend_list_data(7, i)&"""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span></p>"
                        end if
                    'Next
                    else
                        if (DateDiff("n", now(), str_cancel_class_datetime) >= 240 and bol_client_have_valid_contract) then
                            arr_next_td(Day(arr_attend_list_data(0, i))) = arr_next_td(Day(arr_attend_list_data(0, i))) & "<p id=""cancel_class_info_"&arr_attend_list_data(7, i)&"""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span><a id=""cancel_class_link_"&arr_attend_list_data(7, i)&""" href=""javascript:void(0);"" onclick=""javascript:g_obj_reservation_class_ctrl.cancelClass('"&arr_attend_list_data(7, i)&"', '"&arr_attend_list_data(0, i)&"', '"&arr_attend_list_data(1, i)&"');"">"&str_msg_a_cancel_class&"</a><img id=""img_cancel_class_loading_"&arr_attend_list_data(7, i)&"""  width=""25"" height=""25"" src=""/lib/javascript/JQuery/images/ajax-loader.gif"" border=""0"" style=""display:none;""/></p>"
                        '四小時內無法取消 不顯示取消的按鈕 或 合約已到期
                        else
                            arr_next_td(Day(arr_attend_list_data(0, i))) = arr_next_td(Day(arr_attend_list_data(0, i))) & "<p id=""cancel_class_info_"&arr_attend_list_data(7, i)&"""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span></p>"
                        end if
                    end if
                next
            end if
        end if
    end sub

    '萬年曆的製作原則是找出判斷平、閏年的法則，然後定出所求當年或當月的第一天是星期幾。
    Dim i, int_day_num
    Dim str_month_name : str_month_name = ""                        '月份的描述 (數字:10 or 英文:January or 中文:一)
    Dim int_td_num : int_td_num = 0                                 '宣告一個int_td_num變數，來做週末換行的動作 (int_td_num > 7 之後就是該換行了)
    Dim str_th_class_name : str_th_class_name = ""                  '星期幾 的 <th> css class name
    Dim str_td_class_name : str_td_class_name = ""                  '幾號的 的 <td> css class name
    Dim str_attend_date_condition : str_attend_date_condition = ""  'client_attend_list where 條件

    '選到的月份要用到的
    Dim int_dat_year, int_dat_month
    Dim arr_now_td(31)                                              '選到的月每個td要顯示的資訊
    Dim int_now_arr_start_index, int_now_arr_end_index              'arr_prev_td 陣列 跑for迴圈的起始和結束值

    '上個月要用到的
    Dim int_prev_year, int_prev_month
    Dim arr_prev_td(31)                                             '上個月每個td要顯示的資訊
    Dim int_prev_arr_start_index, int_prev_arr_end_index            'arr_prev_td 陣列 跑for迴圈的起始和結束值

    '下個月要用到的
    Dim int_next_year, int_next_month
    Dim arr_next_td(31)                                             '下個月每個td要顯示的資訊
    Dim int_next_arr_start_index, int_next_arr_end_index            'arr_prev_td 陣列 跑for迴圈的起始和結束值

    '變數int_dat_year為萬年曆運算所需的年，變數int_dat_month為萬年曆運算所需的月
    '自動抓取電腦的系統時間做出當月的月曆
    int_dat_year = Request("dat_year")
    int_dat_month = Request("dat_month")

    '*************** 設定預設值 START ***************
    '目前的年度
    if (isEmptyOrNull(int_dat_year)) then
        int_dat_year = Year(now) '年度的格式為 2003 四位數
    else
        int_dat_year = Int(int_dat_year)
    end if

    '目前的月份
    if (isEmptyOrNull(int_dat_month)) then
        int_dat_month = Month(now) '月份的格式為 6 一位數
    else
        int_dat_month = Int(int_dat_month)
    end if

    '設定月份的描述
    str_month_name = int_dat_month
    'select case int_dat_month
        'case 1
            'str_month_name = "January"
        'case 2
            'str_month_name = "February"
        'case 3
            'str_month_name = "March"
        'case 4
            'str_month_name = "April"
        'case 5
            'str_month_name = "May"
        'case 6
            'str_month_name = "June"
        'case 7
            'str_month_name = "July"
        'case 8
            'str_month_name = "August"
        'case 9
            'str_month_name = "September"
        'case 10
            'str_month_name = "October"
        'case 11
            'str_month_name = "November"
        'case 12
            'str_month_name = "December"
    'end select

    '宣告每個月份的天數，而預設二月為平年的天數28天
    Dim int_arr_month_day(12)
    int_arr_month_day(0) = 0
    int_arr_month_day(1) = 31
    int_arr_month_day(2) = 28
    int_arr_month_day(3) = 31
    int_arr_month_day(4) = 30
    int_arr_month_day(5) = 31
    int_arr_month_day(6) = 30
    int_arr_month_day(7) = 31
    int_arr_month_day(8) = 31
    int_arr_month_day(9) = 30
    int_arr_month_day(10) = 31
    int_arr_month_day(11) = 30
    int_arr_month_day(12) = 31
   
    'ASP Weekday 傳回的是 1(星期日),2(星期一),3(星期二),4(星期三),5(星期四),6(星期五),7(星期六)
    Dim str_arr_week_name(8)
    str_arr_week_name(0) = ""
    str_arr_week_name(1) = str_msg_font_sunday
    str_arr_week_name(2) = str_msg_font_monday
    str_arr_week_name(3) = str_msg_font_tuesday
    str_arr_week_name(4) = str_msg_font_wednesday
    str_arr_week_name(5) = str_msg_font_thursday
    str_arr_week_name(6) = str_msg_font_friday
    str_arr_week_name(7) = str_msg_font_saturday

    '判斷是否為閏年或平年。閏年的話，二月則為29天；平年二月則為28天(預設值)。
    if ((int_dat_year mod 4 = 0) and not(int_dat_year mod 100 = 0)) or (int_dat_year mod 400 = 0) then
        int_arr_month_day(2) = 29
    end if

    Dim int_dat_year_pass, int_dat_year_first
    '為判斷元旦而宣告一變數 int_dat_year_pass 從西元紀元始至去年共過了幾年
    '變數int_dat_year_first 表示元旦是屬於星期幾  ps:在vb中 \ 是求除後整數，而 / 是求除後商數可含小數
    int_dat_year_pass = int_dat_year - 1
    int_dat_year_first = (int_dat_year + int_dat_year_pass \ 4 - int_dat_year_pass \ 100 + int_dat_year_pass \ 400) mod 7 
    
    Dim int_day_pass
    '變數 intdaypass 表示今年元旦起至上個月底已過了多少天
    int_day_pass = 0
    for i = 0 to (int_dat_month - 1)
        int_day_pass = int_day_pass + int_arr_month_day(i)
    next

    Dim int_dat_month_of_week_first
    '變數 int_dat_month_of_week_first 表示要顯示月份的當月第一天是星期幾
    int_dat_month_of_week_first = (int_dat_year_first + (int_day_pass mod 7)) mod 7
	
	'20120501 Beson 因mod 7 的結果是0~6 所以0是7而不是0
	if int_dat_month_of_week_first = 0 then
		int_dat_month_of_week_first = 7
	end if

     if ( int_dat_year = 2012 and int_dat_month = 4 ) then
        int_dat_month_of_week_first = 7
    end if

    '*************** 設定預設值 END ***************
'response.Write "int_dat_year_first=" & int_dat_year_first & " -- int_day_pass=" & int_day_pass & " -- int_dat_month_of_week_first=" & int_dat_month_of_week_first
'response.Write int_dat_year_first&" = ("&dat_year &"+"& int_dat_year_pass &"\"& 4 &"-"& int_dat_year_pass &"\"& 100 &"+"& int_dat_year_pass &"\"& 400&") mod"& 7 
'response.Write int_dat_month_of_week_first

%>
<!--課程列表start-->
<div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="calendar_bd">
        <!-- 星期幾 標題 START -->
        <tr>
            <%
                for i = 2 to Ubound(str_arr_week_name)-1
                    str_th_class_name = ""
                    '判斷是否是今日
                    if (Weekday(now) = i and (int_dat_month = Month(now)) and (int_dat_year = Year(now))) then
                        str_th_class_name = "class=""bg_color_brown"""
                    end if

                    '星期一 to 星期六
                    Response.Write "<th "&str_th_class_name&">" & str_arr_week_name(i) & "</th>"
                next

                str_th_class_name = ""
                '判斷是否是今日
                if (Weekday(now) = i) then
                    str_th_class_name = "class=""bg_color_brown"""
                end if
                '星期日
                Response.Write "<th "&str_th_class_name&">" & str_arr_week_name(1) & "</th>"
            %>
        </tr>
        <!-- 星期幾 標題 END -->

        <!-- Day Number START -->
        <%
            Response.Write "<tr>"

            '判斷上一個月份是幾月
            int_prev_month = int_dat_month - 1
            if (int_prev_month = 0) then
                int_prev_month = 12
            end if

            '判斷上一年是幾年
            int_prev_year = int_dat_year
            if (int_dat_month = 1) then
                int_prev_year = int_dat_year - 1
            end if
            '找出上一個月份有訂的課程
            if (int_dat_month_of_week_first > 1) then
                int_prev_arr_start_index = int_arr_month_day(int_prev_month) - ((int_dat_month_of_week_first - 1)-1)
                int_prev_arr_end_index = int_arr_month_day(int_prev_month)

                if (Not isCalenderYearMonthBiggerThanNow(int_prev_year, int_prev_month)) then
                    '目前顯示的月份 1號之前的日期 START
                    for i = int_prev_arr_start_index to int_prev_arr_end_index
                        '輸出前一個月份 MM/DD
                        Response.Write "<td>" & int_prev_month & "/" & i & "</td>"
                        int_td_num = int_td_num + 1
                    next
                    '目前顯示的月份 1號之前的日期 END
                else
                    for i = 1 to (int_dat_month_of_week_first - 1)
                        'client_attend_list.attend_date 的 where 條件
                        if (str_attend_date_condition = "") then
                            str_attend_date_condition = "'" & int_prev_year & "/" & Right("0" & int_prev_month, 2) & "/" & Right("0" & (int_arr_month_day(int_prev_month) - (i-1)), 2) & "'"
                        else
                            str_attend_date_condition = str_attend_date_condition & ",'" & int_prev_year & "/" & Right("0" & int_prev_month, 2) & "/" & Right("0" & (int_arr_month_day(int_prev_month) - (i-1)), 2) & "'"
                        end if
                    next                    

                    'client_attend_list.attend_date 的 where 條件
                    '只撈上個月指定日期有訂的課程
                    if (str_attend_date_condition <> "") then
                        str_attend_date_condition = " AND ( member.ClassRecord.ClassStartDate IN (" & str_attend_date_condition & ")) "
                    end if
                    
                    Call setCalenderArrayData(str_attend_date_condition, 1)
                        
                    '目前顯示的月份 1號之前的日期 START
                    for i = int_prev_arr_start_index to int_prev_arr_end_index
                        if (arr_prev_td(i) = "" or Not isCalenderDateBiggerThanNow(int_prev_year, int_prev_month, i)) then
                            '輸出前一個月份 MM/DD
                            Response.Write "<td>" & int_prev_month & "/" & i & "</td>"
                        else
                            Response.Write "<td>" & int_prev_month & "/" & i & arr_prev_td(i) & "</td>"
                        end if
                        int_td_num = int_td_num + 1
                    next
                    '目前顯示的月份 1號之前的日期 END
                end if
            end if
            

            '顯示選到的月份的日期 START
            str_attend_date_condition = ""
            str_attend_date_condition = " AND ( member.ClassRecord.ClassStartDate BETWEEN '" & int_dat_year & "/" & Right("0" & int_dat_month, 2) & "/01' AND '" & int_dat_year & "/" & Right("0" & int_dat_month, 2) & "/" & int_arr_month_day(int_dat_month) & "') "
            Call setCalenderArrayData(str_attend_date_condition, 2)
                            
            for int_day_num = 1 to int_arr_month_day(int_dat_month)
                str_td_class_name = ""

                'int_td_num > 7 之後就是到星期日 表示該換行了
                if (int_td_num >= 7) then
                    Response.Write "</tr><tr>"
                    int_td_num = 0
                end if

                '判斷是否為今天的日期
                if (Right(CStr("0"&int_day_num),2) = CStr(Right("0"&Day(now),2)) and (int_dat_month = Month(now)) and (int_dat_year = Year(now))) then
                    str_td_class_name = "class=""bg_color_green"""
                end if
    
                if (arr_now_td(int_day_num) = "" or Not isCalenderDateBiggerThanNow(int_dat_year, int_dat_month, int_day_num)) then
                    '輸出Day Number
                    Response.Write "<td "&str_td_class_name&">" & int_day_num & "</td>"
                else
                    Response.Write "<td "&str_td_class_name&">" & int_day_num & arr_now_td(int_day_num) & "</td>"
                end if
                int_td_num = int_td_num + 1
            next
            '顯示選到的月份的日期 END


            '判斷下一個月份是幾月
            int_next_month = int_dat_month + 1
            if (int_next_month > 12) then
                int_next_month = 1
            end if

            '判斷下一年是幾年
            int_next_year = int_dat_year
            if (int_dat_month = 12) then
                int_next_year = int_dat_year + 1
            end if

            '目前顯示的月份 30或31號之後的日期 START
            str_attend_date_condition = ""
            if (int_td_num < 7) then  
                int_next_arr_start_index = 1
                int_next_arr_end_index = 7 - int_td_num
                if (Not isCalenderYearMonthBiggerThanNow(int_next_year, int_next_month)) then
                    for i = int_next_arr_start_index to int_next_arr_end_index                        
                        '輸出下一個月份 MM/DD
                        Response.Write "<td>" & int_next_month & "/" & i & "</td>"
                    next
                else
                    for i = 1 to 7 - int_td_num
                        'client_attend_list.attend_date 的 where 條件
                        if (str_attend_date_condition = "") then
                            str_attend_date_condition = "'" & int_next_year & "/" & Right("0" & int_next_month, 2) & "/" & Right("0" & i, 2) & "'"
                        else
                            str_attend_date_condition = str_attend_date_condition & ",'" & int_next_year & "/" & Right("0" & int_next_month, 2) & "/" & Right("0" & i, 2) & "'"
                        end if
                    next              

                    'client_attend_list.attend_date 的 where 條件
                    '只撈下一年指定日期有訂的課程
                    if (str_attend_date_condition <> "") then
                        str_attend_date_condition = " AND ( member.ClassRecord.ClassStartDate IN (" & str_attend_date_condition & ")) "
                    end if

                    Call setCalenderArrayData(str_attend_date_condition, 3)
                        
                    for i = int_next_arr_start_index to int_next_arr_end_index
                        if (arr_next_td(i) = "" or Not isCalenderDateBiggerThanNow(int_next_year, int_next_month, i)) then
                            '輸出下一個月份 MM/DD
                            Response.Write "<td>" & int_next_month & "/" & i & "</td>"
                        else
                            Response.Write "<td>" & int_next_month & "/" & i & arr_next_td(i) & "</td>"
                        end if
                    next
                end if
            end if
            '目前顯示的月份 30或31號之後的日期 END

            Response.Write "</tr>"

        %>
        <!-- Day Number END -->
    </table>
 </div>
<!--課程列表end-->