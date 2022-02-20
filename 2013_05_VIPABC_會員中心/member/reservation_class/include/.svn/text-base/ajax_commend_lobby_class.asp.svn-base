<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<%
    '被 /program/member/reservation_class/normal_and_vip.asp ajax 動態載入
    '被 /program/member/reservation_class/normal_and_vip_near_order.asp ajax 動態載入
    '被 /program/member/reservation_class/normal_and_vip_near_cancel.asp ajax 動態載入

    Dim str_sql_lobby_session, str_lobby_condition_datetime, arr_lobby_data, i
    Dim str_con_img_url, str_suit_lev, str_arr_tmp
    
    str_con_img_url = ""
    '設定 lobby_session 時間條件 (大會堂允許開課前5分鐘都可以訂)
    str_lobby_condition_datetime = " AND ((CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE(), 111) AND lobby_session.session_time >= '"&Hour(now)&"' ) " &_
	                                 "      OR (CONVERT(varchar, lobby_session.session_date, 111) > CONVERT(varchar, GETDATE(), 111))) "


    '--有開課的大會堂及顧問的資料(>=今天)
 	str_sql_lobby_session = " SELECT TOP 2 "
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
    'TODO: 判斷語系
    if (Session("language") = CONST_LANG_CN) then
        str_sql_lobby_session = str_sql_lobby_session & " lobby_session.topic_cn, "
    else
        str_sql_lobby_session = str_sql_lobby_session & " lobby_session.topic, "
    end if

    '--大會堂內容簡介 6 
    'TODO: 判斷語系
    if (Session("language") = CONST_LANG_CN) then
        str_sql_lobby_session = str_sql_lobby_session & " lobby_session.introduction_cn, "
    else
        str_sql_lobby_session = str_sql_lobby_session & " lobby_session.introduction, "
    end if

    '--大會堂session_sn 7
  	str_sql_lobby_session = str_sql_lobby_session & " CONVERT(varchar, lobby_session.session_date,112) + CAST(lobby_session.session_time AS varchar(2)) + CAST(lobby_session.sn AS varchar(5)) AS dcode, "
    '--大會堂適合的等級 8
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.lev, "
    '--大會堂人數上限 9
    str_sql_lobby_session = str_sql_lobby_session & " lobby_session.limit, "
    '--大會堂的主題系列編號 10
    str_sql_lobby_session = str_sql_lobby_session & " ISNULL(lobby_session.lobby_type, '') AS lobby_type, "
    '--大會堂的主題系列名稱 11
    str_sql_lobby_session = str_sql_lobby_session & " ISNULL(lobby_session_type.type_name, '') AS type_name, "
    '顧問的大頭照 hex code 12
    str_sql_lobby_session = str_sql_lobby_session & " ISNULL(con_basic.img, '') AS con_img, "
    '--顧問名字 13 fname
  	str_sql_lobby_session = str_sql_lobby_session & " isnull(con_basic.basic_fname, '') AS con_fname, "
    '--顧問名字 14 lname
  	str_sql_lobby_session = str_sql_lobby_session & " isnull(con_basic.basic_lname,'') AS con_lname "
  	str_sql_lobby_session = str_sql_lobby_session & " FROM lobby_session "
  	str_sql_lobby_session = str_sql_lobby_session & " LEFT JOIN con_basic ON lobby_session.consultant = con_basic.con_sn "   
    str_sql_lobby_session = str_sql_lobby_session & " LEFT JOIN lobby_session_type ON lobby_session.lobby_type = lobby_session_type.sn "
  	str_sql_lobby_session = str_sql_lobby_session & " WHERE (lobby_session.valid = 1)  " & str_lobby_condition_datetime
    str_sql_lobby_session = str_sql_lobby_session & " ORDER BY lobby_session.session_date ASC, lobby_session.session_time ASC "

%>
<script type="text/javascript">
<!--
    // 顯示或關閉推薦大會堂
    function showCommendLobbyClass()
    {
        if ($("#div_commend_lobby_class_info_section").css("display") == "none")
        {
            $("#div_commend_lobby_class_info_section").show("fast");
            $("#a_open_and_close_commend_lobby_class").text(g_obj_class_data.arr_msg.close_commend_lobby);    // 關閉
        }
        else
        {
            $("#div_commend_lobby_class_info_section").hide();
            $("#a_open_and_close_commend_lobby_class").text(g_obj_class_data.arr_msg.open_commend_lobby);     // 打開
        }
    }
//-->
</script>
<style type="text/css">
<!--
    .cmd_con_introduction a:link    {text-decoration:none;}
    .cmd_con_introduction a:visited {text-decoration:none;}
    .cmd_con_introduction a:active  {text-decoration:underline;}
    .cmd_con_introduction a:hover   {text-decoration:underline;}
-->
</style>
<!--標題start-->
<div class="main_memtitle_line_03">
<div class="main_memtitle_line_left_03"><%=str_msg_font_commend_lobby%><!--推薦大會堂--></div>
<div class="main_memtitle_line_right_03">
<a id="a_open_and_close_commend_lobby_class" class="close_btn01" href="javascript:void(0);" onclick="showCommendLobbyClass();"><%=str_msg_close_commend_lobby%></a></div><!--关闭-->
<div class="clear">
</div>
</div>
<!--標題end-->
<!--課程列表start-->
<div id="div_commend_lobby_class_info_section">
<% 
    arr_lobby_data = excuteSqlStatementReadQuick(str_sql_lobby_session, CONST_VIPABC_RW_CONN)

    if (isSelectQuerySuccess(arr_lobby_data)) then
        if (Ubound(arr_lobby_data) >= 0) then
            for i = 0 to Ubound(arr_lobby_data, 2)
%>
<ul class="class_list">    
    <li>
        <%
            '顧問的大頭照
            if (Not isEmptyOrNull(arr_lobby_data(12, i))) then
                str_con_img_url = "/lib/functions/functions_show_consultant_image.asp?con_sn=" & arr_lobby_data(4, i) & "&website=" & CONST_VIPABC_WEBSITE
            else
                '若找不到顧問的大頭照 則show出預設的圖片
                str_con_img_url = "/images/no_face_img.jpg"
            end if
        %>
        <!--<img src="<%=str_con_img_url%>" width="70" style="cursor:pointer;" border="0" title="<%=getWord("RESERVATION_LOBBY_CLASS")%>"
             onclick="javascript:$('#hdn_commend_lobby_date').val('<%=Trim(arr_lobby_data(1, i))%>'); 
                        $('#hdn_commend_lobby_time').val('<%=Trim(arr_lobby_data(2, i))%>');
                        $('#hdn_commend_con_ename').val('<%=Escape(Trim(arr_lobby_data(3, i)))%>');
                        $('#hdn_commend_lobby_topic').val('<%=Escape(Trim(arr_lobby_data(5, i)))%>');
                        $('#hdn_commend_lobby_sn').val('<%=Trim(arr_lobby_data(0, i))%>');
                        $('#frm_commend_lobby_reservation').submit();"/>-->
            <a href="/program/member/reservation_class/lobby_near_order.asp"><img src="<%=str_con_img_url%>" width="70" style="cursor:pointer;" border="0" title="<%=getWord("RESERVATION_LOBBY_CLASS")%>"/></a>
        <h2>
            <%
                '開課日期
                Response.Write "<span>" & Trim(arr_lobby_data(1, i)) & "</span><br/>"

                '開課時間
                Response.Write "<span>" & Trim(arr_lobby_data(2, i)) & ":30" & "</span><br/>"

                '顧問名字
                Response.Write "<span>" & Trim(arr_lobby_data(13, i)) & "<br/>" & Trim(arr_lobby_data(14, i)) & "</span>"
            %>
        </h2>
        <h3>
            <p>
            <%
                '大會堂的標題
                Response.Write "<span>" & Trim(arr_lobby_data(5, i)) & "</span>"
            %>
            </p>
                <%'大會堂內容簡介 %>
                <span class="cmd_con_introduction">
                    <!--<a href="javascript:void(0);" title="<%=getWord("RESERVATION_LOBBY_CLASS")%>" 
                        onclick="javascript:$('#hdn_commend_lobby_date').val('<%=Trim(arr_lobby_data(1, i))%>'); 
                            $('#hdn_commend_lobby_time').val('<%=Trim(arr_lobby_data(2, i))%>');
                            $('#hdn_commend_con_ename').val('<%=Escape(Trim(arr_lobby_data(3, i)))%>');
                            $('#hdn_commend_lobby_topic').val('<%=Escape(Trim(arr_lobby_data(5, i)))%>');
                            $('#hdn_commend_lobby_sn').val('<%=Trim(arr_lobby_data(0, i))%>');
                            $('#frm_commend_lobby_reservation').submit(); return false;">
                        <%=Trim(arr_lobby_data(6, i))%>
                    </a>-->
                    <a href="/program/member/reservation_class/lobby_near_order.asp" title="<%=getWord("RESERVATION_LOBBY_CLASS")%>" ><%=Trim(arr_lobby_data(6, i))%></a>
                </span>
        </h3>
        <h4>
            <%
                '大會堂適合的等級
                str_suit_lev = Trim(arr_lobby_data(8, i))
                if (Not isEmptyOrNull(str_suit_lev)) then
				    if (Left(str_suit_lev, 1) = ",") then
                        str_suit_lev = Right(str_suit_lev, Len(str_suit_lev)-1)
                    end if
				    if (Right(str_suit_lev, 1) = ",") then 
                        str_suit_lev = Left(str_suit_lev, Len(str_suit_lev)-1)
                    end if
				    str_arr_tmp = Split(str_suit_lev, ",")
				    Response.Write str_msg_font_suit_level_eng & str_arr_tmp(0) & "-" & str_arr_tmp(UBound(str_arr_tmp)) & "<br/>"  '適用lv
                end if                       
            %>
            <%=str_msg_font_cut_word%><!--扣--> <span class="color_1">1</span> <%=str_msg_font_list_1%><!--课时数-->
        </h4>
    </li>
</ul>
<%          next
        end if
    end if 
%>
    <!--<form id="frm_commend_lobby_reservation" method="post" action="/program/member/reservation_class/lobby_near_order.asp">
        <input type="hidden" id="hdn_commend_lobby_date" name="lobby_date" />
        <input type="hidden" id="hdn_commend_lobby_time" name="lobby_time" />
        <input type="hidden" id="hdn_commend_con_ename" name="con_ename" />
        <input type="hidden" id="hdn_commend_lobby_topic" name="lobby_topic" />
        <input type="hidden" id="hdn_commend_lobby_sn" name="lobby_sn" />
    </form>-->
</div>
<!--課程列表end-->