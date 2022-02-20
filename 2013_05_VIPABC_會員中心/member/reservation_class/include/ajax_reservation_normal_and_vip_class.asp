<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<%
    '一對一和一對多訂課 

    '被 /program/member/reservation_clas/normal_and_vip.asp ajax 載入
    '被 /program/member/reservation_class/normal_and_vip_near_order.asp ajax 載入

    'TODO: 選到已經預定的課程的處理?

    '所在頁面的url
    Dim str_now_browse_prog_url : str_now_browse_prog_url = Trim(unEscape(Request("page_url")))
%>
<script type="text/javascript" src="/program/member/reservation_class/javascript/reservation_class.js"></script>
<script type="text/javascript">
<!--
    // TODO: 為什麼跑了兩次 ready?
    $(document).ready(function(){
        var str_class_type = "<%=Trim(Request("sel_class_type"))%>";
        var str_class_date = "<%=Trim(Request("txt_class_date"))%>";
        var str_class_time = "<%=Trim(Request("sel_class_time"))%>";
        var str_contract_start_date = "<%=getFormatDateTime(str_service_start_date, 7)%>";   // 合約開始日
        var str_contract_end_date = "<%=getFormatDateTime(str_service_end_date, 7)%>";    // 合約結束日
        var dat_min_date, dat_max_date;
        var int_reservation_class_cycle_day = 30;   // 可預約幾天內的課
        var bol_first_is_today = true;

        var dat_contract_end_date;
        if (str_contract_end_date != "")
        {
            dat_contract_end_date = new Date(str_contract_end_date.substr(0, 4), 
                                            parseInt(str_contract_end_date.substr(4, 2), 10) - 1, 
                                            str_contract_end_date.substr(6, 2));
        }

        var dat_contract_start_date = new Date(str_contract_start_date.substr(0, 4), 
                                               parseInt(str_contract_start_date.substr(4, 2), 10) - 1, 
                                               str_contract_start_date.substr(6, 2));
        var dat_contract_start_date_added = new Date(str_contract_start_date.substr(0, 4), 
                                               parseInt(str_contract_start_date.substr(4, 2), 10) - 1, 
                                               str_contract_start_date.substr(6, 2));
        dat_contract_start_date_added.setDate(dat_contract_start_date_added.getDate() + int_reservation_class_cycle_day);    // 加上30天

        var str_now_sys_datetime = "<%=getFormatDateTime(now, 7) %>";   // 系統時間
        var dat_now_sys_datetime = new Date(str_now_sys_datetime.substr(0, 4), 
                                               parseInt(str_now_sys_datetime.substr(4, 2), 10) - 1, 
                                               str_now_sys_datetime.substr(6, 2));
        var dat_now_sys_datetime_added = new Date(str_now_sys_datetime.substr(0, 4), 
                                               parseInt(str_now_sys_datetime.substr(4, 2), 10) - 1, 
                                               str_now_sys_datetime.substr(6, 2));
        dat_now_sys_datetime_added.setDate(dat_now_sys_datetime_added.getDate() + int_reservation_class_cycle_day);    // 加上30天

        // 可選擇合約開始日起一個月內的時間
        if (dat_contract_start_date > dat_now_sys_datetime)
        {
            bol_first_is_today = false;
            $("#txt_class_date").Watermark(g_obj_class_data.arr_msg.no_sel_class_date).datepicker({ 
                                           showOtherMonths: true, 
                                           minDate: new Date(dat_contract_start_date.getFullYear(), dat_contract_start_date.getMonth(), dat_contract_start_date.getDate()), 
                                           maxDate: new Date(dat_contract_start_date_added.getFullYear(), dat_contract_start_date_added.getMonth(), dat_contract_start_date_added.getDate()), 
                                           rangeSelect: false, 
                                           changeFirstDay: false, 
                                           dateFormat:"yy/mm/dd"}).keydown(function(){return false;});
        }
        // 是 首次上課客戶且有預約First Session，預約的日期和時間
        else if (g_obj_class_data.str_last_first_session_datetime != "-1" && g_obj_class_data.str_last_first_session_datetime != "-2")
        {
            var dat_first_session = new Date(g_obj_class_data.str_last_first_session_datetime.substr(0, 4), 
                                               parseInt(g_obj_class_data.str_last_first_session_datetime.substr(4, 2), 10) - 1, 
                                               g_obj_class_data.str_last_first_session_datetime.substr(6, 2));

            if (dat_first_session > dat_now_sys_datetime)
            {
                bol_first_is_today = false;
                var dat_first_session_added = new Date(g_obj_class_data.str_last_first_session_datetime.substr(0, 4), 
                                                   parseInt(g_obj_class_data.str_last_first_session_datetime.substr(4, 2), 10) - 1, 
                                                   g_obj_class_data.str_last_first_session_datetime.substr(6, 2));
                dat_first_session_added.setDate(dat_first_session_added.getDate() + int_reservation_class_cycle_day);    // 加上30天

                dat_min_date = dat_first_session;
                dat_max_date = dat_first_session_added;

                $("#txt_class_date").Watermark(g_obj_class_data.arr_msg.no_sel_class_date).datepicker({ 
                                               showOtherMonths: true, 
                                               minDate: new Date(dat_min_date.getFullYear(), dat_min_date.getMonth(), dat_min_date.getDate()), 
                                               maxDate: new Date(dat_max_date.getFullYear(), dat_max_date.getMonth(), dat_max_date.getDate()), 
                                               rangeSelect: false, 
                                               changeFirstDay: false, dateFormat:"yy/mm/dd"}).keydown(function(){return false;});
            }
        }

        // 可選擇今日起一個月內的時間
        if (bol_first_is_today)
        {
            dat_min_date = dat_now_sys_datetime;
            dat_max_date = dat_now_sys_datetime_added;

            // 若系統日期>合約結束日 則 alert 合约于YYYY/MM/DD到期请，至 ｢学习纪录｣ 查询详细资讯。
            // 合約已到期
            if (str_contract_end_date == "" || dat_now_sys_datetime > dat_contract_end_date)
            {
                // 合約已到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
                alert(g_obj_class_data.arr_msg.contract_edate_invalid_0 + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
                $("#txt_class_date").attr("disabled", "disabled");  // disabled 日期
                $("#btn_add_reservation_class").attr("disabled", "disabled");  // disabled 新增按鈕
            }
            // 系統日期+30天 > 合約結束日
            else if (dat_contract_end_date < dat_now_sys_datetime_added)
            {
                dat_min_date = dat_now_sys_datetime;
                dat_max_date = dat_contract_end_date;
            }

            $("#txt_class_date").Watermark(g_obj_class_data.arr_msg.no_sel_class_date).datepicker({ 
                                           showOtherMonths: true, 
                                           minDate: new Date(dat_min_date.getFullYear(), dat_min_date.getMonth(), dat_min_date.getDate()), 
                                           maxDate: new Date(dat_max_date.getFullYear(), dat_max_date.getMonth(), dat_max_date.getDate()), 
                                           rangeSelect: false, 
                                           changeFirstDay: false, dateFormat:"yy/mm/dd"}).keydown(function(){return false;});
        }        

        // 註冊點擊事件 當點擊日曆時觸發
        $("#ui-datepicker-for-datepicker-div").bind("click", {clear_old_date : 0}, g_obj_reservation_class_ctrl.getHtmlSelectClassTime);

        // 註冊課程類型事件 onchange 觸發
        $("#sel_class_type").bind("change", {clear_old_date : 1}, g_obj_reservation_class_ctrl.getHtmlSelectClassTime);

        // 檢查是否有預訂課程 改變顯示字樣
        if ($("#hdn_reservation_class_count").val() == "0")
        {
            if (document.getElementById("div_near_order_class") != null)
            {
                $("#div_near_order_class").text(g_obj_class_data.arr_msg.a_no_have_order_class);
            }
            $("#ul_reservation_class_info").hide();     // 隱藏 訂課資訊列表
            $("#div_reservation_class_result").hide();  // 隱藏 訂課結果及確認送出按鈕
        }

        // 若是從預訂課程頁面 訂課後轉址到 此次預訂課程頁面，則會自動新增一筆訂課資訊
        if (str_class_type != "")
        {
            g_obj_reservation_class_ctrl.addNewClassInfo(str_class_type, str_class_date, str_class_time);
        }
    });
//-->
</script>
<!--上版位課程start-->
<div class="top_class">
    <form id="form_reservation_class_module" method="post" action="/program/member/reservation_class/normal_and_vip_near_order.asp" >
        <%'hdn_now_browse_program_name 隱藏值 是為了要判斷頁面是否已經在此次預訂課程頁面 若是的話點擊後就不需再轉址到那 %>
        <input type="hidden" id="hdn_now_browse_program_name" value="<%=str_now_browse_prog_url%>" />
        <%'hdn_show_reservation_class_remind_msg 隱藏值 是為了要判斷是否是首次點擊+新增 若有值的話點擊後就不需再顯示提醒訊息 %>
        <input type="hidden" id="hdn_show_reservation_class_remind_msg" name="hdn_show_reservation_class_remind_msg" value="<%=Trim(Request("hdn_show_reservation_class_remind_msg"))%>" />
        <%'hdn_reservation_class_count 隱藏值 是為了要紀錄已新增了幾筆課程%>
        <input type="hidden" id="hdn_reservation_class_count" name="hdn_reservation_class_count" value="0" />
        <%'hdn_prev_select_class_date 隱藏值 是為了要紀錄前一次選擇的日期%>
        <input type="hidden" id="hdn_prev_select_class_date" value="" />
        <h1><%=str_msg_font_normal_class_head%></h1><!--请选择课程类型 / 时间-->
        <ul>
            <li>
                <select id="sel_class_type" name="sel_class_type" class="class_select" style="width:166px;">
                <%
                    '取得 產品對應課程類型的扣堂規則
                    Dim arr_product_session_rule, i
                    arr_product_session_rule = getProductSessionRule(str_now_use_product_sn, 0, CONST_VIPABC_RW_CONN)
                    if (IsArray(arr_product_session_rule)) then
                        for i = 0 to Ubound(arr_product_session_rule, 2)
                            '一般课程/小班(1-3人)课程
                            if (CStr(CONST_CLASS_TYPE_ONE_ON_THREE) = CStr(arr_product_session_rule(3, i))) then
                                Response.Write "<option value="""&Right("0"&CONST_CLASS_TYPE_ONE_ON_THREE, 2)&""" selected=""selected"">" & str_msg_sel_normal_option & "</option>"

                            '一般课程/1对1课程
                            elseif (CStr(CONST_CLASS_TYPE_ONE_ON_ONE) = CStr(arr_product_session_rule(3, i))) then
                                Response.Write "<option value="""&Right("0"&CONST_CLASS_TYPE_ONE_ON_ONE, 2)&""">" & str_msg_sel_vip_option & "</option>"
                            end if
                        next
                    end if
                %>                              
                </select>
            </li>
            <li>
                <input type="text" id="txt_class_date" name="txt_class_date" size="10" style="width:100px;"/></li>    
            <li id="li_reservation_class_time" class="noline">
                <select id="sel_class_time" name="sel_class_time" class="class_select" style="width:160px;" disabled="disabled" >
                    <option value=""><%=str_msg_no_sel_class_date%></option><!--请选择上课日期-->
                </select>
            </li>
            <li class="noline">
                <input type="button" id="btn_add_reservation_class" value="+ <%=str_msg_btn_add_normal_class%>" class="btn_class" onclick="return g_obj_reservation_class_ctrl.addNewReservationClass();"/><!--新增-->
            </li>
        </ul>
    </form>
</div>
<%
    '檢查目前瀏覽頁面 不是 此次預訂課程 和 近期取消課程
    if (str_now_browse_prog_url <> "/program/member/reservation_class/normal_and_vip_near_order.asp" and str_now_browse_prog_url <> "/program/member/reservation_class/normal_and_vip_near_cancel.asp") then
%>
<a class="top_class_N2" href="/program/member/reservation_class/normal_and_vip_near_order.asp"><%=str_msg_a_have_order_class%><!--此次预订課程--></a>
<a class="top_class_N2" href="/program/member/reservation_class/normal_and_vip_near_cancel.asp"><%=str_msg_a_have_cancel_class%><!--近期取消課程--></a>
<%
    elseif (str_now_browse_prog_url = "/program/member/reservation_class/normal_and_vip_near_order.asp") then
%>
<div id="div_near_order_class" class="top_class_N3">
<%=str_msg_a_no_have_order_class%><!--尚未预订谘询课程--></div>
<a class="top_class_N2" href="/program/member/reservation_class/normal_and_vip_near_cancel.asp"><%=str_msg_a_have_cancel_class%><!--近期取消課程--></a>
<%
    elseif (str_now_browse_prog_url = "/program/member/reservation_class_cancel_order.asp") then
%>
<a class="top_class_N2" href="/program/member/reservation_class/normal_and_vip_near_order.asp"><%=str_msg_a_have_order_class%><!--此次预订課程--></a>
<div id="div_near_cancel_class" class="top_class_N3">
<%=str_msg_a_no_have_cancel_class%><!--近期尚未取消課程--></div>
<%
    end if
%>
<div class="clear"></div>
<!--上版位課程end-->