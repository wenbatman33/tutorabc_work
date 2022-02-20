<!--#include virtual="/lib/include/global.inc"-->
<%   
Dim strLobbySn : strLobbySn = getRequest("lobby_sn", CONST_DECODE_NO) '大會堂sn	
if ( Not isEmptyOrNull(strLobbySn) ) then
    strSql = " SELECT svalue FROM lobby_session WHERE sn = " & strLobbySn & ""
    'response.write lobby_sn
    g_arr_result = excuteSqlStatementReadQuick(strSql, CONST_TUTORABC_R_CONN)
    if ( isSelectQuerySuccess (g_arr_result) ) then
	    for g_int_row = 0 To Ubound(g_arr_result,1) 
		    Response.write g_arr_result(0,0)
	    next
    else
	    '錯誤訊息
	    'Response.Write("錯誤訊息" & g_arr_result & "<br>")
    end if
else
    'response.write "error"
end if
%>