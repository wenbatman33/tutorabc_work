<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<%
    '大會堂開課資訊搜尋

    '被 /program/member/reservation_class/lobby_near_order.asp ajax 載入

    Dim str_topic_opt, str_consultant_opt                                                       '有開課的大會堂主題, 有開課的大會堂顧問
    Dim str_lobby_condition_datetime : str_lobby_condition_datetime = ""                        'lobby_session.session_date & lobby_session.session_time where 條件  
    Dim str_now_browse_prog_url : str_now_browse_prog_url = Trim(unEscape(Request("page_url"))) '所在頁面的url
    Dim str_from_other_page_con_sn : str_from_other_page_con_sn = Trim(Request("from_other_page_con_sn"))     '搜尋頁面和推薦大會堂也可以訂課 若有值則表示有訂課
    Dim str_sql_lobby_session, arr_lobby_data, arr_lobby_data_index, int_time_index, str_arr_lobby_type_name
    
    str_topic_opt = ""
    str_consultant_opt = ""


    ''設定 lobby_session 時間條件 (大會堂只允許訂24時候的課)
    'g_str_lobby_condition_datetime = " AND ((CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE()+1, 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_
	                                 '"      OR (CONVERT(varchar, lobby_session.session_date, 111) > CONVERT(varchar, GETDATE(), 111))) "

    '設定 lobby_session 時間條件 (大會堂允許開課前5分鐘都可以訂)
    str_lobby_condition_datetime = " AND ((CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE(), 111) AND lobby_session.session_time >= '"&Hour(now)&"' ) " &_
	                                 "      OR (CONVERT(varchar, lobby_session.session_date, 111) > CONVERT(varchar, GETDATE(), 111))) "

    '大會堂主題系列編號和名稱
    str_sql_lobby_session = "SELECT * FROM lobby_session_type"
    arr_lobby_data = excuteSqlStatementReadQuick(str_sql_lobby_session, CONST_VIPABC_RW_CONN)
    if (isSelectQuerySuccess(arr_lobby_data)) then
        if (Ubound(arr_lobby_data) >= 0) then
            for arr_lobby_data_index = 0 to Ubound(arr_lobby_data, 2)
                if (Not isEmptyOrNull(arr_lobby_data(1, arr_lobby_data_index))) then
                    str_arr_lobby_type_name = Split(arr_lobby_data(1, arr_lobby_data_index), "-")
                    str_topic_opt = str_topic_opt & "<option value=""" & arr_lobby_data(0, arr_lobby_data_index) & """>" & str_arr_lobby_type_name(1) & "</option>"
                end if
            next
        else
            'TODO: 沒資料時的預設值
        end if
    else
        'TODO: 發生錯誤時的預設值
    end if


    '--有開課的大會堂及顧問的資料(>=今天)
 	str_sql_lobby_session = " SELECT "
    '--大會堂的編號 0
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.sn, "
    '--開課日期 1 
    str_sql_lobby_session = str_sql_lobby_session & " CONVERT(varchar, lobby_session.session_date, 111) AS sdate, "
    '--開課時間 2
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.session_time AS stime, "
    '--顧問名字 3
  	str_sql_lobby_session = str_sql_lobby_session & " isnull(con_basic.basic_fname, '') + ' ' + isnull(con_basic.basic_lname,'') AS ename, "
    '--顧問編號 4 
    str_sql_lobby_session = str_sql_lobby_session & " con_basic.con_sn, "
    '--大會堂的標題 5
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.topic, "
    '--大會堂內容簡介 6 
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.introduction, "
    '--大會堂session_sn 7
  	str_sql_lobby_session = str_sql_lobby_session & " CONVERT(varchar, lobby_session.session_date,112) + CAST(lobby_session.session_time AS varchar(2)) + CAST(lobby_session.sn AS varchar(5)) AS dcode, "
    '--大會堂適合的等級 8
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.lev, "
    '--大會堂人數上限 9
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.limit, "
    '--大會堂的主題系列編號 10
    str_sql_lobby_session = str_sql_lobby_session & " ISNULL(lobby_session.lobby_type, '') AS lobby_type, "
    '--大會堂的主題系列名稱 11
    str_sql_lobby_session = str_sql_lobby_session & " ISNULL(lobby_session_type.type_name, '') AS type_name "
  	str_sql_lobby_session = str_sql_lobby_session & " FROM lobby_session "
  	str_sql_lobby_session = str_sql_lobby_session & " LEFT JOIN con_basic ON lobby_session.consultant = con_basic.con_sn "   
    str_sql_lobby_session = str_sql_lobby_session & " LEFT JOIN lobby_session_type ON lobby_session.lobby_type = lobby_session_type.sn "
  	str_sql_lobby_session = str_sql_lobby_session & " WHERE (lobby_session.valid = 1)  " & str_lobby_condition_datetime
    str_sql_lobby_session = str_sql_lobby_session & " ORDER BY lobby_session.session_date ASC "

    arr_lobby_data = excuteSqlStatementReadQuick(str_sql_lobby_session, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_lobby_data)) then
        if (Ubound(arr_lobby_data) >= 0) then
            for arr_lobby_data_index = 0 to Ubound(arr_lobby_data, 2)
                '顧問名字
                str_consultant_opt = str_consultant_opt & "<option value=""" & arr_lobby_data(4, arr_lobby_data_index) & """>" & arr_lobby_data(3, arr_lobby_data_index) & "</option>"
            next
        else
            'TODO: 沒資料時的預設值
        end if
    else
        'TODO: 發生錯誤時的預設值
    end if
%>
<!--搜尋大會堂start-->
<div class="top_class">
    <h1><%=str_msg_font_lobby_class_head%><!--快速选择大会堂--></h1>
    <ul>
        <li>                        
            <select id="sel_search_lobby_topic" name="sel_search_lobby_topic" class="class_select" style="width:120px;" >
                <option value="" selected="selected"><%=str_msg_sel_lobby_topic%><!--请选择主题系列--></option>
                <%=str_topic_opt%>
            </select>
        </li>
        <li>
            <select id="sel_search_lobby_consultant" name="sel_search_lobby_consultant" class="class_select" >
                <option value="" selected="selected"><%=str_msg_sel_lobby_consultant%><!--请选择顾问--></option>
                <%=str_consultant_opt%>
            </select>
        </li>
        <li>
            <%=getHtmlSelectLevel(str_msg_sel_suit_level, "", "")%><!--適用程度-->
        </li>
        <li class="noline" >
            <select id="sel_search_lobby_time" name="sel_search_lobby_time" class="class_select" style="width:120px;">
                <option value="" selected="selected"><%=str_msg_sel_lobby_time%><!--请选择课程時段--></option>
                <%
                    for int_time_index = 6 to 23
                        Response.Write "<option value="""&int_time_index&""">" & int_time_index & ":30</option>"
                    next
                %>
            </select>
        </li>
        <li class="noline">
            <input type="button" id="btn_search_lobby" value="+ <%=str_msg_btn_search_lobby_class%>" class="btn_class" onclick="g_obj_reservation_lobby_class_ctrl.searchLobbyClassInfo();" /><!--确认-->
        </li>
    </ul>
</div>
<!--搜尋大會堂end-->
<!--上版位課程start-->
<div class="top_class2n"></div>
<%
    if (str_now_browse_prog_url = "/program/member/reservation_class/lobby_near_order.asp") then
%>
<div id="div_near_order_class" class="top_class_N3">
<%
    if (isEmptyOrNull(str_from_other_page_con_sn)) then
        Response.Write str_msg_a_no_have_order_class '您此次尚未预订谘询课程
    else
        Response.Write str_msg_a_have_order_class    '此次预订課程
    end if
%><!--您此次尚未预订谘询课程--></div>
<a class="top_class_N4" href="/program/member/reservation_class/lobby_near_cancel.asp"><%=str_msg_a_have_cancel_lobby_class%><!--近期取消大会堂课程--></a>
<%
    elseif (str_now_browse_prog_url = "/program/member/reservation_class/lobby_near_cancel.asp") then
%>
<a class="top_class_N4" href="/program/member/reservation_class/lobby_near_order.asp"><%=str_msg_a_have_order_lobby_class%><!--近期预订大会堂课程--></a>
<div id="div_near_cancel_class" class="top_class_N3">
<%=str_msg_a_no_have_cancel_class%><!--您近期尚未取消課程--></div>
<%end if%>
<div class="clear"></div>
<!--上版位課程end-->