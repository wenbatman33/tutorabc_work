<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<%
    Dim dat_monday          '星期一的日期
    Dim dat_today           '今天的日期
    Dim dat_weekday_today   '今天是星期幾 1:星期天 2:星期一 3:星期二 4:星期三 5:星期四 6:星期五 7:星期六

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

%>
<div class="date">
    <%=Year(dat_monday)%><span class="fontSize_12"><%=str_msg_font_year%><!--年--></span><%=Month(dat_monday)%><span class="fontSize_12"><%=str_msg_font_month%><!--月--></span>
</div>
<div class="date_btn">
    <a href="javascript:void(0);" onclick="javascript:g_obj_class_calender.refreshWeek('<%=DateAdd("d", -7, dat_monday)%>');"><img src="/images/images_cn/add/back_btn.gif" width="18" height="11" /></a>
    <a href="javascript:void(0);" onclick="javascript:g_obj_class_calender.refreshWeek('<%=DateAdd("d", 7, dat_monday)%>');"><img src="/images/images_cn/add/next_btn.gif" width="18" height="11" /></a>
    <a class="s_gray" href="javascript:void(0);" onclick="javascript:g_obj_class_calender.refreshWeek('');">today</a>
</div>