<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<%
    sub setCalenderArrayData(p_str_attend_date_condition)
    '/****************************************************************************************
    '描述      ：設定要顯示的日期及訂課資訊
    '傳入參數  ：[p_str_sql:String] client_attend_list.attend_date where 的條件
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

        'str_sql = "SELECT CONVERT(varchar, client_attend_list.attend_date, 111) AS sdate, " &_
        '          "client_attend_list.attend_sestime, " &_
        '          "ISNULL(con_basic.basic_fname,'') + ' ' + ISNULL(con_basic.basic_lname,'') AS con_ename, " &_
        '          "client_attend_list.attend_livesession_types, " &_
        '          "ISNULL(client_attend_list.attend_level, 0) AS attend_level, " &_                          
        '          str_lobby_session_topic & ", " &_
        '          "ISNULL(lobby_session.lev, '') AS lobby_suit_level, " &_
        '          "client_attend_list.sn " &_
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

        if (isSelectQuerySuccess(arr_attend_list_data)) then
            if (Ubound(arr_attend_list_data) >= 0) then
                for i = 0 to Ubound(arr_attend_list_data, 2)
                    str_cancel_class_datetime = arr_attend_list_data(0, i) & " " & Right("0"&arr_attend_list_data(1, i), 2) & ":30"
                    str_title = arr_attend_list_data(0, i) & chr(10)
                    str_title = str_title & arr_attend_list_data(1, i) & ":30" & chr(10)

                    '轉成新Table的值
	                for intI = 0 to Ubound(arrClassTypeVIPJR) 
	                    if ( sCInt(arr_attend_list_data(3, i)) = sCInt(arrClassTypeVIPJR(intI)) ) then
                            arr_attend_list_data(3, i) = arrClassTypeVIPABC(intI)
                            exit for
                        end if
	                next

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
                            str_title = str_title & "LV" & str_arr_tmp(0) & "-" & str_arr_tmp(UBound(str_arr_tmp))
                        end if

                        str_ball_style_class = "class=""ball_blue"""
                        str_title = str_title & str_msg_sel_lobby_option
                    '一對三 ball_red
                    elseif ( Int(arr_attend_list_data(3, i)) = CONST_CLASS_TYPE_ONE_ON_THREE or Int(arr_attend_list_data(3, i)) = CONST_CLASS_TYPE_ONE_ON_FOUR ) then
                        str_ball_style_class = "class=""ball_red"""
                        str_title = str_title & str_msg_sel_normal_option
                    '一對一 ball_green
                    elseif (Int(arr_attend_list_data(3, i)) = CONST_CLASS_TYPE_ONE_ON_ONE) then
                        str_ball_style_class = "class=""ball_green"""
                        str_title = str_title & str_msg_sel_vip_option
                    end if

                    str_title = getReplaceInjectionWord(str_title)
                    
                    if (DateDiff("n", now(), str_cancel_class_datetime) >= 240 and bol_client_have_valid_contract) then
                        arr_now_td(Day(arr_attend_list_data(0, i))) = arr_now_td(Day(arr_attend_list_data(0, i))) & "<div id=""cancel_class_info_"&arr_attend_list_data(7, i)&""" class="""" style=""float:left; margin-left:10px""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span><a id=""cancel_class_link_"&arr_attend_list_data(7, i)&""" href=""javascript:void(0);"" onclick=""javascript:g_obj_reservation_class_ctrl.cancelClass('"&arr_attend_list_data(7, i)&"', '"&arr_attend_list_data(0, i)&"', '"&arr_attend_list_data(1, i)&"');"">"&str_msg_a_cancel_class&"</a><img id=""img_cancel_class_loading_"&arr_attend_list_data(7, i)&""" src=""/lib/javascript/JQuery/images/ajax-loader.gif"" width=""25"" height=""25"" border=""0"" style=""display:none;""/></div>"
                    
                    '四小時內無法取消 不顯示取消的按鈕 或 合約已到期
                    else
                        arr_now_td(Day(arr_attend_list_data(0, i))) = arr_now_td(Day(arr_attend_list_data(0, i))) & "<div id=""cancel_class_info_"&arr_attend_list_data(7, i)&""" class="""" style=""float:left; margin-left:10px""><span "&str_ball_style_class&" title=""" & str_title & """>●</span><span class=""color_1"" title=""" & str_title & """>" & arr_attend_list_data(1, i) & ":30</span></div>"
                    end if

                next
            end if
        end if
    end sub

    Dim arr_weekday_name : arr_weekday_name = Array(str_msg_font_monday, str_msg_font_tuesday, str_msg_font_wednesday, str_msg_font_thursday, str_msg_font_friday, str_msg_font_saturday, str_msg_font_sunday)
    Dim dat_monday, dat_sunday  '星期一的日期 星期日的日期
    Dim dat_today               '今天的日期
    Dim dat_weekday_today       '今天是星期幾 1:星期天 2:星期一 3:星期二 4:星期三 5:星期四 6:星期五 7:星期六
    Dim arr_now_td(31)          '每個td要顯示的資訊 (一週內)
    Dim str_attend_date_condition : str_attend_date_condition = ""  'client_attend_list where 條件
    Dim dat_add_monday, str_style_class_week, str_style_class_date, i
    

    '取得今天的日期
    dat_today = Date

    '取得今天是星期幾
    dat_weekday_today = DatePart("w", dat_today)

    '設定星期一的日期
    dat_monday = Trim(Request("dat_monday"))
    if (dat_monday = "") then
        '根據今天的日期，找出禮拜一的相對日期
        '由於asp的星期一代碼為"2"...以此類推到7 而星期日為"1" 所以如果不是星期天則根據當天星期幾參數進行位移
        if (dat_weekday_today <> 1) then
	        dat_monday = DateAdd("d", 2-dat_weekday_today, dat_today)

        '如果為週末 直接向前移六天
        else    
	        dat_monday = DateAdd("d", -6, dat_today)
        end if
    end if

    '設定星期日的日期
    dat_sunday = DateAdd("d", 6, dat_monday)

    str_attend_date_condition = " AND (member.ClassRecord.ClassStartDate BETWEEN '" & dat_monday & "' AND '" & dat_sunday & "') "            

    Call setCalenderArrayData(str_attend_date_condition)       
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="calendar_bd">
    <%
        for i = 0 to 6
            Response.Write "<tr>"

            str_style_class_week = "class=""shortbar"""
            str_style_class_date = "class=""logobar"""
            
            '判斷是否是今日
            dat_add_monday = DateAdd("d", i, dat_monday)
            if (dat_add_monday = Date) then
                str_style_class_week = "class=""bg_color_brown"""
                str_style_class_date = "class=""bg_color_green"""
            end if

            '輸出 Week Name
            Response.Write "<th "&str_style_class_week&">" & arr_weekday_name(i) & "</th>"


            '輸出 Month / Day
            Response.Write "<td "&str_style_class_date&">"
            Response.Write "<div class=""floatL_a"" title="""&dat_add_monday&""" >"
            Response.Write Month(dat_add_monday) & "/" & Day(dat_add_monday)
            Response.Write "</div>"

            '輸出訂課資訊
            if (arr_now_td(Day(dat_add_monday)) <> "" and isCalenderDateBiggerThanNow(Year(dat_add_monday), Month(dat_add_monday), Day(dat_add_monday))) then
                Response.Write arr_now_td(Day(dat_add_monday))
            end if

            Response.Write "</td>"

            Response.Write "</tr>"
        next
    %>
</table>