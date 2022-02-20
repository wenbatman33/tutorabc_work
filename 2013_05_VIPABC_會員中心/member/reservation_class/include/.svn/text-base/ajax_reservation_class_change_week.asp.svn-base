<!--#include virtual="/lib/include/global.inc"-->
<%
    Dim dat_monday          '星期一的日期
    Dim strClassType : strClassType = getRequest("class_type", CONST_DECODE_NO)
    Dim strProductType : strProductType = getRequest("product_type", CONST_DECODE_NO)
    'cfg_product.point_type
    Dim strProductPointType : strProductPointType = getRequest("product_point_type", CONST_DECODE_NO)

    '設定星期一的日期
    dat_monday = Trim(Request("dat_monday"))
    if (isEmptyOrNull(dat_monday)) then
        '根據今天的日期，找出這周禮拜一的日期
        'dat_monday = getMondayOfWeek()
        dat_monday = date
    end if
%>
<div class="time">
    <span class="year"><%=Year(dat_monday)%></span>
    <span class="word">年</span>
    <span class="year"><%=Month(dat_monday)%></span>
    <span class="word">月</span>
</div>
<div class="next">
    <a href="javascript:void(0);" onclick="javascript:g_obj_class.refreshWeek('<%=DateAdd("d", -7, dat_monday)%>', '<%=strClassType%>', '<%=strProductType%>');"><img src="/images/images_cn/add/back_btn.gif" width="18" height="11" alt="Prev Week" /></a>
    <a href="javascript:void(0);" onclick="javascript:g_obj_class.refreshWeek('<%=DateAdd("d", 7, dat_monday)%>', '<%=strClassType%>', '<%=strProductType%>');"><img src="/images/images_cn/add/next_btn.gif" width="18" height="11" alt="Next Week"/></a>
    <span class="today"><a href="javascript:void(0);" onclick="javascript:g_obj_class.refreshWeek('');">today</a></span>
</div>