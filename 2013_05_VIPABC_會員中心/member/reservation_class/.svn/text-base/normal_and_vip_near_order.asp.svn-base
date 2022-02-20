<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include file="include/include_prepare_and_check_data.inc"-->
<% 
    '使用新的課程前導頁
    Response.Redirect "class_choose.asp"
    Response.End
%>
<script type="text/javascript">
<!--
    $(document).ready(function(){
        var str_now_page_url = '<%=Request.ServerVariables("PATH_INFO")%>';
        var str_class_type = '<%=Trim(Request("sel_class_type")) %>';
        var str_class_date = '<%=Trim(Request("txt_class_date")) %>';
        var str_class_time = '<%=Trim(Request("sel_class_time"))%>';
        var str_class_remind_msg = '<%=Trim(Request("hdn_show_reservation_class_remind_msg"))%>';
        var str_contract_end_date = "<%=str_service_end_date%>";    // 合約結束日

        if (str_contract_end_date == "")
        {
            // 合約已到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
            alert(g_obj_class_data.arr_msg.contract_edate_invalid_0 + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);      
            $("#remind_contract_start_msg").text("<%=getWord("WELCOME_JOIN_VIPABC")%>" + "! " + g_obj_class_data.arr_msg.contract_edate_invalid_0 + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);          
            $("#ul_reservation_class_info").hide();
            $("#div_reservation_class_result").hide();
        }
        else
        {
            // Ajax 載入 
            $("#div_add_normal_vip_class").load("/program/member/reservation_class/include/ajax_reservation_normal_and_vip_class.asp", 
                                                {
                                                    ndate : new Date(), 
                                                    page_url: escape(str_now_page_url), 
                                                    sel_class_type : str_class_type,
                                                    txt_class_date : str_class_date,
                                                    sel_class_time : str_class_time,
                                                    hdn_show_reservation_class_remind_msg : str_class_remind_msg
                                                });

            // Ajax 載入 推薦大會堂
            $("#div_commend_lobby_class_info").load("/program/member/reservation_class/include/ajax_commend_lobby_class.asp", {ndate : new Date()});
        }
    });
//-->
</script>
<!--內容start-->
<div class="main_membox">
    <!--新增一般/VIP 課程 START-->
    <div id="div_add_normal_vip_class"></div>
    <!--新增一般/VIP 課程 END-->
    <!--2010.01.18新增:找不到或沒有資料的共用訊息框-->
    <div id="div_no_reservation_class_remind_msg" class="wowbox">
        <div class="wow">
        <p id="remind_contract_start_msg" style="margin-top:10px; margin-bottom:10px;"><%=getWord("WELCOME_JOIN_VIPABC") & "! " & getWord("YOU_CAN_RESERVATION_NOW") & "<font color='red'>" & str_service_start_date & "</font>" & getWord("AFTER_CONTRACT_START")%></p><!--欢迎您加入VIPABC行列! 您现在可预订YYYY/MM/DD(合约开始)之後的课程。-->
        <div id="div_no_reservation_class_for_normal" style="margin-top:10px; margin-bottom:10px;"><%=getWord("NO_RESERVATION_CLASS")%><!--尚未预订课程--></div>
        </div>
    </div>
    <!--2010.01.18新增:找不到或沒有資料的共用訊息框-->
    <!--課程列表start-->
    <ul id="ul_reservation_class_info" class="class_list2">
        <li id="">
            <h1>&nbsp;</h1>
            <h2><%=str_msg_font_date%></h2><!--日期-->
            <h3><%=str_msg_font_time%></h3><!--时间-->
            <h4><%=str_msg_font_lobby_class_list_head_3%></h4><!--预计扣除课时数-->
        </li>
    </ul>
    <div id="div_reservation_class_result" class="list_txt">
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
            <form id="form_reservation_class_module_order" method="post" action="/program/member/reservation_class/reservation_class_exe.asp" onsubmit="javascript:return g_obj_reservation_class_ctrl.submitReservationClass();" >
                <%'hdn_total_reservation_class_dt 隱藏值 客戶已預訂的課程日期和時間 YYYYMMDDHH%>
                <input type="hidden" id="hdn_total_reservation_class_dt" name="hdn_total_reservation_class_dt" value="" />
                <%'hdn_order_or_reorder_type: 1=預定 2=重新預定%>
                <input type="hidden" id="hdn_order_or_reorder_type" name="hdn_order_or_reorder_type" value="1" />
                <input type="submit" id="btn_submit_reservation_class" value="+ <%=str_msg_btn_confirm_order%>" class="btn_1" /><!--确认订課-->
            </form>
        </div>
    </div>
    <!--課程列表end-->
    <!--推薦大會堂 start-->
    <div id="div_commend_lobby_class_info"></div>
    <!--推薦大會堂 end--> 
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->