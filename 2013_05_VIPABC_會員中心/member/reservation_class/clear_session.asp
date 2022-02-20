<!--#include virtual="/lib/include/global.inc"-->
<% 
dim strCheckDivOpen : strCheckDivOpen = getRequest("strCheckDivOpen", CONST_DECODE_ESCAPE)

'dim strSession : strSession =  getSession(strCheckDivOpen, CONST_DECODE_ESCAPE)
Session(strCheckDivOpen) = "false"
%>
