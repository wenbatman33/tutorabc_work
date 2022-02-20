<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include file="include/include_prepare_and_check_data.inc"-->
<%
    '停用頁面的快取暫存
    Call setPageNoCache()

    Dim str_btn_url

    '客戶剩餘課時數
    Dim flt_request_client_total_session : flt_request_client_total_session = Trim(Request("client_total_session"))

    '總共要扣除的課時數
    Dim flt_total_subtract_session : flt_total_subtract_session = Trim(Request("total_subtract_session"))

    if (isEmptyOrNull(flt_request_client_total_session)) then
        flt_request_client_total_session = 0.0
    else
        flt_request_client_total_session = CDbl(flt_request_client_total_session)
    end if

    if (isEmptyOrNull(flt_total_subtract_session)) then
        flt_total_subtract_session = 0.0
    else
        flt_total_subtract_session = CDbl(flt_total_subtract_session)
    end if

    if (Trim(Request("class_type")) = CStr(CONST_CLASS_TYPE_LOBBY)) then
        str_btn_url = "/program/member/reservation_class/lobby_near_order.asp"
    else
        str_btn_url = "/program/member/reservation_class/normal_and_vip.asp"
    end if
%>
<!--內容start-->
<div class="okbox">
    <div class="ok">
        <%
            '成功預定 或 重新預定的訊息
            if (Trim(Request("total_order_count")) <> "0") then
                Response.Write str_msg_font_order_success_1 & "<br/>" '您已成功预订
                Response.Write "<span class=""bold"">" & Replace(Trim(unEscape(Request("success_msg"))), "$br$", "<br>") & "</span>" & str_msg_font_order_success_2 & Trim(Request("total_order_count")) & str_msg_font_order_success_3   '課程共 XX 堂
            end if

            '失敗的訊息
            if (Not isEmptyOrNull(Trim(unEscape(Request("fault_msg"))))) then
                Response.Write "<br/>" & "<span class=""bold"">" & Replace(Trim(unEscape(Request("fault_msg"))), "$br$", "<br>") & "</span>"
            end if
        %>
        <%=str_msg_font_order_result_4%><!--共计扣除课时数--> <span class="bold"><%=flt_total_subtract_session%></span> <%=str_msg_font_order_result_3%><!--可用课时数--> <span class="bold"><%=Round(flt_request_client_total_session - flt_total_subtract_session, 2)%></span>
    </div>
    <div class="btn">
        <input type="button"name="button" id="button" value="+ <%=str_msg_btn_return_to_order_class%>" class="btn_1" onclick="javascript:location.href='<%=str_btn_url%>';" /><!--返回预订课程-->
        <input type="button" name="button" id="button" value="+ <%=str_msg_btn_watch_class%>" class=" m-left5 btn_1" onclick="javascript:location.href='/program/member/reservation_class/watch_class.asp';" /><!--查看课表-->
    </div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->