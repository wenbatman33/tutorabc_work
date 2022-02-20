<!--#include virtual="/lib/include/global.inc" -->
<%
'aboo  強制清除所有的Session值
Session.abandon
Response.Cookies("client_sn") = ""
Response.Cookies("client_sn").Expires = DateAdd( "d", -1, Now() )

'session_feedback_success.asp
Call alertGo(getWord("logout_success"), "/", CONST_ALERT_THEN_REDIRECT )
%>