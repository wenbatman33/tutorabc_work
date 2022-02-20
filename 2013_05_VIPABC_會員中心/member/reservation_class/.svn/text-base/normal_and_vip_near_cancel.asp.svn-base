<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include file="include/include_prepare_and_check_data.inc"-->
<%
    '使用新的課程前導頁
    Response.Redirect "class_choose.asp"
    Response.End

    '一對一和一對多近期取消

    Dim str_attend_condition_datetime : str_attend_condition_datetime = ""  'client_attend_list.attend_date & client_attend_list.attend_sestime where 條件    
    Dim str_sql, arr_attend_list_data, var_arr_param_value
    Dim int_client_sn, i, arr_cancel_normal_data

    arr_cancel_normal_data = ""

    '取得客戶編號
    if (g_var_client_sn <> "") then
        int_client_sn = Session("client_sn")
    end if


    '設定 client_attend_list 時間條件 (一對一和一對多只允許訂24hour以後的課)
    str_attend_condition_datetime = " AND ((CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE()+1, 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
	                                "      OR (CONVERT(varchar, client_attend_list.attend_date, 111) > CONVERT(varchar, GETDATE(), 111))) "

    '--找出客戶近期取消的一對一和一對多
    str_sql = "SELECT  "
    '--課程日期 0
    str_sql = str_sql & "CONVERT(varchar, client_attend_list.attend_date, 111) AS sdate, "
    '--課程時間 1
    str_sql = str_sql & "client_attend_list.attend_sestime, "
    '--扣的堂數 2,
    str_sql = str_sql & "client_attend_list.svalue, "
    '--一對一或一對多 3,
    str_sql = str_sql & "client_attend_list.attend_livesession_types "
    str_sql = str_sql & "FROM client_attend_list "
    str_sql = str_sql & "WHERE (client_attend_list.client_sn = @client_sn) AND (client_attend_list.valid = 0) AND (client_attend_list.special_sn is null) AND (ISNULL(client_attend_list.attend_level,0) <= 12)" & str_attend_condition_datetime
    str_sql = str_sql & "ORDER BY client_attend_list.attend_date, client_attend_list.attend_sestime "

    var_arr_param_value = Array(int_client_sn)
    arr_attend_list_data = excuteSqlStatementRead(str_sql, var_arr_param_value, CONST_VIPABC_RW_CONN)
    if (isSelectQuerySuccess(arr_attend_list_data)) then
        if (Ubound(arr_attend_list_data) >= 0) then
            for i = 0 to Ubound(arr_attend_list_data, 2)
                arr_cancel_normal_data = arr_cancel_normal_data & ",{class_type:" & """" & arr_attend_list_data(3, i) & """" & "," & "sdate:" & """" & arr_attend_list_data(0, i) & """" & "," & "stime:" & """" & Right("0"&arr_attend_list_data(1, i),2) & """" & "," & "svalue:" & """" & arr_attend_list_data(2, i) & """" & "}"           
            next
            arr_cancel_normal_data = Right(arr_cancel_normal_data, Len(arr_cancel_normal_data)-1)            
        else
            'TODO: 沒資料時的預設值
        end if
    else
        'TODO: 發生錯誤時的預設值
    end if

    '輸出 javascript Array
    Response.Write "<script type=""text/javascript"">"
    Response.write " var arr_cancel_normal_data_info = [" & arr_cancel_normal_data & "]; "    '已取消的一對一和一對多資料
    Response.Write "</script>"
%>
<script type="text/javascript" src="/program/member/reservation_class/javascript/reservation_class.js"></script>
<script type="text/javascript">
<!--
    $(document).ready(function(){
        var int_add_result, str_order_class_datetime, i;
        var str_contract_end_date = "<%=str_service_end_date%>";    // 合約結束日

        // 合約已到期
        if (str_contract_end_date == "")
        {
            // 合約已到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
            alert(g_obj_class_data.arr_msg.contract_edate_invalid_0 + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);            
        }
        else
        {
            // 若近期有取消課程 則顯示在課程列表
            if (arr_cancel_normal_data_info.length > 0)
            {
                $("#div_near_cancel_class").text(g_obj_class_data.arr_msg.a_have_cancel_class); // 近期取消課程
                $("#div_no_cancel_class_for_normal").hide();
                $("#ul_reservation_class_info").show();
                $("#div_reservation_class_result").show();
                for (i = 0; i < arr_cancel_normal_data_info.length; i++)
                {
                    int_add_result = g_obj_reservation_class_ctrl.addNewClassInfo(arr_cancel_normal_data_info[i].class_type, arr_cancel_normal_data_info[i].sdate, arr_cancel_normal_data_info[i].stime);

                    str_order_class_datetime = arr_cancel_normal_data_info[i].sdate.replace(new RegExp("/","gm"), "") + arr_cancel_normal_data_info[i].stime; //  訂課日期時間 YYYYMMDDHHmm

                    // 觸發點擊事件 轉換為 "重訂此課程"
                    $("#btn_"+str_order_class_datetime).click();
                }
            }
        }

        // Ajax 載入 推薦大會堂
        $("#div_commend_lobby_class_info").load("/program/member/reservation_class/include/ajax_commend_lobby_class.asp", {ndate : new Date()});
    });
//-->
</script>
<!--內容start-->
<div class="main_membox">
    <!--上版位課程start-->
    <div class="top_class2n"></div>
    <a class="top_class_N4" href="/program/member/reservation_class/normal_and_vip_near_order.asp"><%=str_msg_a_have_order_class%></a><!--此次预订課程-->
    <div id="div_near_cancel_class" class="top_class_N3">
    <%=str_msg_a_no_have_cancel_class%><!--您近期尚未取消課程--></div>
    <div class="clear"></div>
    <!--2010.01.18新增:找不到或沒有資料的共用訊息框-->
    <div id="div_no_cancel_class_for_normal" class="wowbox">
        <div class="wow">
        <div style="margin-top:10px; margin-bottom:10px;"><%=getWord("NO_CANCEL_CLASS")%><!--无取消课程--></div>
        </div>
    </div>
    <!--2010.01.18新增:找不到或沒有資料的共用訊息框-->       
    <!--上版位課程end-->
    <!--課程列表start-->
    <ul id="ul_reservation_class_info" class="class_list2" style="display:none;">
        <li>
            <h1>&nbsp;</h1>
            <h2><%=str_msg_font_date%></h2><!--日期-->
            <h3><%=str_msg_font_time%></h3><!--时间-->
            <h4><%=str_msg_font_lobby_class_list_head_3%></h4><!--预计扣除课时数-->
            <h6>&nbsp;</h6>
        </li>
    </ul>
    <div id="div_reservation_class_result" class="list_txt" style="display:none;">
        <form id="form_reservation_class_module_cancel" method="post" action="/program/member/reservation_class/reservation_class_exe.asp" onsubmit="javascript:return g_obj_reservation_class_ctrl.submitReservationClass();" >
            <%'hdn_reservation_class_count 隱藏值 是為了要紀錄已新增了幾筆課程%>
            <input type="hidden" id="hdn_reservation_class_count" name="hdn_reservation_class_count" value="0" />
            <%'hdn_prev_select_class_date 隱藏值 是為了要紀錄前一次選擇的日期%>
            <input type="hidden" id="hdn_prev_select_class_date" value="" />        
            <%'hdn_total_reservation_class_dt 隱藏值 客戶已預訂的課程日期和時間 YYYYMMDDHH%>
            <input type="hidden" id="hdn_total_reservation_class_dt" name="hdn_total_reservation_class_dt" />
            <%'hdn_order_or_reorder_type: 1=預定 2=重新預定%>
            <input type="hidden" id="hdn_order_or_reorder_type" name="hdn_order_or_reorder_type" value="2" />
            <p>
            <span class="ball_red">●</span><%=str_msg_font_normal_class%></p><!--一般课( 1-6人 )-->
            <p class="color_green">
            <span class="ball_green">●</span><%=str_msg_font_vip_class%></p><!--一般课( 1對1 )-->
            <p>
                <%=str_msg_font_order_result_1%><!--您此次预订课时共--> <span id="span_reservation_class_count"></span> <%=str_msg_font_order_result_2%><!--次--><span class="m-left20"><span class="color_5"><%=str_msg_font_lobby_class_list_head_3%><!--预计扣除课时数--></span>
                <span id="span_reservation_class_total_subtract_session"></span> <span class="color_5"><%=str_msg_font_order_result_3%><!--可用课时数--></span> <span id="span_client_total_session"></span></span>
            </p>
            <br class="clear" />
            <div class="t-right">
            <input type="submit" id="btn_submit_reservation_class" value="+ <%=str_msg_btn_confirm_order%>" class="btn_1" /></div><!--确认订課-->
        </form>
    </div>
    <!--課程列表end-->
    <!--推薦大會堂 start-->
    <div id="div_commend_lobby_class_info"></div>
    <!--推薦大會堂 end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->