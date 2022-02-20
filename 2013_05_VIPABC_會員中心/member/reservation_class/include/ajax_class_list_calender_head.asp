<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<%
    Dim dat_year, dat_month, int_next_year, int_prev_year, int_next_month, int_prev_month
    Dim dat_now_year : dat_now_year = Year(now)
    Dim dat_now_month : dat_now_month = Month(now)

    '變數dat_year為萬年曆運算所需的年，變數dat_month為萬年曆運算所需的月
    '自動抓取電腦的系統時間做出當月的月曆
    dat_year = Request("dat_year")
    dat_month = Request("dat_month")

    '*************** 設定預設值 START ***************
    '目前的年度
    if (isEmptyOrNull(dat_year)) then
        dat_year = Year(now) '年度的格式為 2003 四位數
    else
        dat_year = Int(dat_year)
    end if

    '目前的月份
    if (isEmptyOrNull(dat_month)) then
        dat_month = Month(now) '月份的格式為 6 一位數
    else
        dat_month = Int(dat_month)
    end if
    '*************** 設定預設值 END ***************

    '----- 查詢上個月的按鈕 ------------------
    '如果是 1 月份
    if (dat_month - 1 = 0) then
        int_prev_year = dat_year - 1
        int_prev_month = 12
    else
        int_prev_year = dat_year
        int_prev_month = dat_month - 1
    end if

    '---- 查詢下個月的按鈕--------------------
	'如果是 12月的話 
	if (dat_month + 1 = 13) then
        int_next_year = dat_year + 1
        int_next_month = 1
	else
        int_next_year = dat_year
        int_next_month = dat_month + 1
	end if
%>
<div class="date">
    <%=dat_year%><span class="fontSize_12"><%=str_msg_font_year%><!--年--></span><%=dat_month%><span class="fontSize_12"><%=str_msg_font_month%><!--月--></span>
</div>
<div class="date_btn">
    <a href="javascript:void(0);" onclick="javascript:g_obj_class_calender.refreshCalender('<%=int_prev_year%>', '<%=int_prev_month%>');"><img src="/images/images_cn/add/back_btn.gif" width="18" height="11" /></a>
    <a href="javascript:void(0);" onclick="javascript:g_obj_class_calender.refreshCalender('<%=int_next_year%>', '<%=int_next_month%>');"><img src="/images/images_cn/add/next_btn.gif" width="18" height="11" /></a>
    <a class="s_gray" href="javascript:void(0);" onclick="javascript:g_obj_class_calender.refreshCalender('<%=dat_now_year%>', '<%=dat_now_month%>');">today</a>
</div>