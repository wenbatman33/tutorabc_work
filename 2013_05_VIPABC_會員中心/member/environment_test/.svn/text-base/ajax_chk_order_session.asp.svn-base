<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_global.asp"-->
<%
'==========判斷這個人是否已經在這個時段預約過，如果預約過回傳已預約的日期 ==start========
Dim date_order_test_session_time '預約測試的時間
date_order_test_session_time = getAlreadyOrderTestTime(request("client_sn"),CONST_TUTORABC_RW_CONN)
if (Not IsArray(date_order_test_session_time)) then
	response.write ""
else
	response.write  date_order_test_session_time(0)
end if
'==========判斷這個人是否已經在這個時段預約過，如果預約過回傳已預約的日期 ==end==========
%>