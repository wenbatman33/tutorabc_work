<!--#include virtual="/lib/functions/functions_member.asp"-->
<% 
'判斷是否為VIPABCJR客戶
dim intVIPABCJR : intVIPABCJR = session("VIPABCJR")
dim bolVIPABCJR : bolVIPABCJR = ""
dim strVIPABCJR : strVIPABCJR = ""

if ( true = bolTestCase ) then
    bolVIPABCJR = true
end if

if ( 1 = sCInt(intVIPABCJR) ) then
	bolVIPABCJR = true
end if

if ( isEmptyOrNull(g_var_client_sn) ) then
    g_var_client_sn = session("client_sn")
end if

if ( isEmptyOrNull(bolVIPABCJR) ) then
	bolVIPABCJR = isVIPABCJRClient(g_var_client_sn, CONST_VIPABC_RW_CONN)
end if

if ( true = bolVIPABCJR ) then
    strVIPABCJR = "_JR"
end if
%>