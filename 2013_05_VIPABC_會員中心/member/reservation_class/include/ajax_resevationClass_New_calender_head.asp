<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<%
Function IsLeapYear(p_year)
	If ((p_year mod 4 = 0) and not(p_year mod 100 = 0)) or (p_year mod 400 = 0) Then
		IsLeapYear = True
	Else
		IsLeapYear = False
	End If
End Function

    Dim dat_year, dat_month, dat_day, int_next_year, int_prev_year, int_next_month, int_prev_month, int_next_day, int_prev_day
    
    '變數dat_year為萬年曆運算所需的年，變數dat_month為萬年曆運算所需的月
    '自動抓取電腦的系統時間做出當月的月曆
    dat_year = Request("dat_year")
    dat_month = Request("dat_month")
    dat_day = Request("dat_day")

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
	
    '目前的日期
    if (isEmptyOrNull(dat_day)) then
        dat_day = Day(now)
    else
        dat_day = Int(dat_day)
    end if
    '*************** 設定預設值 END ***************

	
    '宣告每個月份的天數，而預設二月為平年的天數28天
    Dim int_arr_month_day(12)
    int_arr_month_day(0) = 0
    int_arr_month_day(1) = 31
    int_arr_month_day(2) = 28
    int_arr_month_day(3) = 31
    int_arr_month_day(4) = 30
    int_arr_month_day(5) = 31
    int_arr_month_day(6) = 30
    int_arr_month_day(7) = 31
    int_arr_month_day(8) = 31
    int_arr_month_day(9) = 30
    int_arr_month_day(10) = 31
    int_arr_month_day(11) = 30
    int_arr_month_day(12) = 31
   
   
    '----- 查詢上一周的按鈕 ------------------
	If (dat_day < 7) Then	'如果是7號之前
		If (dat_month - 1 = 0) Then	'如果是一月份
			int_prev_year = dat_year - 1
			int_prev_month = 12
		Else
			int_prev_year = dat_year
			int_prev_month = dat_month - 1
		End If
	
		If (IsLeapYear(int_prev_year) = True) Then
			int_arr_month_day(2) = 29
		End If
		
		int_prev_day = dat_day + int_arr_month_day(dat_month-1) - 7
	Else
		int_prev_year = dat_year
        int_prev_month = dat_month
		int_prev_day = dat_day - 7
	End If
	
	
    '---- 查詢下一周的按鈕--------------------
	If (IsLeapYear(dat_year) = True) Then
		int_arr_month_day(2) = 29
	End If
	
	If ((int_arr_month_day(dat_month)-7) < dat_day) Then	'如果在月底前七天內
		If (dat_month + 1 = 13) Then	'如果是 12月
			int_next_year = dat_year + 1
			int_next_month = 1
		Else
			int_next_year = dat_year
			int_next_month = dat_month + 1
		End If
		
		int_next_day = dat_day - int_arr_month_day(dat_month) + 7
	Else
		int_next_year = dat_year
		int_next_month = dat_month
		int_next_day = dat_day + 7
	End If
%>
<div class="time">
	<span class="year"><%=dat_year%></span><span class="word"><%=str_msg_font_year%><!--年--></span>
	<span class="year"><%=dat_month%></span><span class="word"><%=str_msg_font_month%><!--月--></span>
</div>
<div class="next">
<%
	'TODO: 如果 int_prev_year int_prev_month int_prev_day 在合約開始日期之前, 則onclick開啟alert訊息 -->
	
%>
	<a href="javascript:void(0);" onclick="javascript:g_obj_resevation_class_calender.refreshCalender('<%=int_prev_year%>', '<%=int_prev_month%>', '<%=int_prev_day%>'); return false;">
<%
	'TODO
%>
		<img src="/images/images_cn/add/back_btn.gif" width="18" height="11" /></a>
	<a href="javascript:void(0);" onclick="javascript:g_obj_resevation_class_calender.refreshCalender('<%=int_next_year%>', '<%=int_next_month%>', '<%=int_next_day%>'); return false;">
		<img src="/images/images_cn/add/next_btn.gif" width="18" height="11" /></a>
	<span class="today"><a href="javascript:void(0);" onclick="javascript:g_obj_resevation_class_calender.refreshCalender('<%=Year(Now())%>', '<%=Month(Now())%>', '<%=Day(Now())%>'); return false;">
		today
	</a></span>
</div>