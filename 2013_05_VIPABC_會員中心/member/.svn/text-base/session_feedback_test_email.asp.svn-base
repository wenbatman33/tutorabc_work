<!--#include virtual="/lib/include/global.inc"-->


<%

'function sendEmail(ByVal p_var_arr_email_from_paras, ByVal p_var_arr_email_to_paras, ByVal p_var_arr_set_paras, ByVal p_var_arr_opt_paras, ByVal p_int_website)

Dim tmpArray(3)
Dim tmpArray2(0)
Dim tmpArray3(1)
Dim tmpArray4(2)

tmpArray(0) = CONST_MAIL_ADDRESS_SERVICES_ABC '寄件者的電子信箱
tmpArray(1) = "ABOO CHEN" '寄件者的name
tmpArray(2) = "subject test"
tmpArray(3) = "test Group sendmail at vipabc"

tmpArray2(0) = "rd@tutorabc.com"

tmpArray3(0) = CONST_MAIL_SERVER_EDM
tmpArray3(1) = "utf-8"

tmpArray4(0) = false
tmpArray4(1) = false
tmpArray4(2) = false

sendEmail tmpArray, tmpArray2, tmpArray3, tmpArray4, CONST_VIPABC_WEBSITE

response.write "senddone! at" & now()
%>