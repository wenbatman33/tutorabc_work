<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
'取得系統時間
dim bolDebugMode : bolDebugMode = false
dim strSysNowDateTime : strSysNowDateTime = getFormatDateTime(now(), 2)
dim intContractSn : intContractSn = getSession("contract_sn", CONST_DECODE_NO)
dim strGetSessionSn : strGetSessionSn = getSessionRoomTime() '欲找的session_sn
dim strTodaySessionHour : strTodaySessionHour = "" '今日上課 時
dim strTodaySessionDate : strTodaySessionDate = "" '今日上課 年月日
dim intSpecialSn : intSpecialSn = ""
dim intGoinClassMinute : intGoinClassMinute = "" '可進教室時間
dim intEndClassMinute : intEndClassMinute = CONST_DEFAULT_FINISH_CLASS_MINUTE '課程結束時間
dim strGoinClassDateTime : strGoinClassDateTime = ""
dim strFinishClassDateTime : strFinishClassDateTime = ""
dim intNowHour : intNowHour = hour(now)

dim arrParam : arrParam = null '傳入sql的array
dim strSql : strSql = "" 'sql
dim arrSqlResult : arrSqlResult = "" 'sql結果
dim intSqlRow : intSqlRow = 0
	
'------------------ 判斷合約未確認需導至確認頁 ------------- start -----------
'Session("contract_sn")-->有合約但尚未確認合約才會有值
'20101119 阿捨修正 續約客戶邏輯 當有未點擊同意合約時 且 合約服務日已到 才需導至合約確認頁
if ( not isEmptyOrNull(intContractSn)) then  
	arrParam = Array(g_var_client_sn, intContractSn)
    '修正續約客戶
	strSql = " SELECT contract_sn "
	strSql = strSql & " FROM client_purchase(nolock) "
	strSql = strSql & " WHERE (client_sn = @client_sn) AND (contract_sn = @contract_sn) AND (Valid = 1) "
	strSql = strSql & " AND datediff(day,product_sdate,GETDATE())>=0 AND datediff(day, GETDATE(),product_edate)>=0 "
	strSql = strSql & " ORDER BY contract_sn "
	arrSqlResult = excuteSqlStatementRead(strSql, arrParam, CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arrSqlResult)) then
		if (Ubound(arrSqlResult) >= 0) then
			'確認合約
			Call alertGo("", "contract_audit.asp", CONST_NOT_ALERT_THEN_REDIRECT)
			response.end
		end if
	end if
end if
'------------------ 判斷合約未確認需導至確認頁 ------------- end -------------

arrSqlResult = getClientAttendListJR(g_var_client_sn, strGetSessionSn)
if (isSelectQuerySuccess(arrSqlResult)) then
	if (Ubound(arrSqlResult) >= 0) then
		for intSqlRow = 0 to Ubound(arrSqlResult, 2) '只捉二筆
			if ( intNowHour = 0 ) then
                if ( intSqlRow = 0 and left(strGetSessionSn, 8) = left(arrSqlResult(2, intSqlRow), 8) ) then '捉第一筆
				    strTodaySessionHour = arrSqlResult(1, intSqlRow)
                    strTodaySessionDate = arrSqlResult(0, intSqlRow)
                    intSpecialSn = arrSqlResult(4, intSqlRow)
			    end if
            end if
            if ( intSqlRow = 0 and left(strGetSessionSn, 8) = left(arrSqlResult(2, intSqlRow), 8) ) then '捉第一筆
				strTodaySessionHour = arrSqlResult(1, intSqlRow)
                strTodaySessionDate = arrSqlResult(0, intSqlRow)
                intSpecialSn = arrSqlResult(4, intSqlRow)
			end if
		next
	else
        '導去上課頁面
		Call alertGo("", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT)
		response.end
	end if
end if

		
'大會堂20進教室，其他為26分以後才能進教室
if ( Not IsEmptyOrNull(intSpecialSn) ) then
    intGoinClassMinute = CONST_DEFAULT_FREELOBBY_GOIN_CLASS_MINUTE
else
    intGoinClassMinute = CONST_DEFAULT_GOIN_CLASS_MINUTE	
end if

if( Not isEmptyOrNull(strTodaySessionHour) ) then 
    if ( strTodaySessionHour = 23 ) then '臨界點, 要跨天
        strTodaySessionFinishHour = 0
        strTodaySessionFinishDate = DateAdd("d", 1, strTodaySessionDate)
    else
        strTodaySessionFinishHour = strTodaySessionHour + 1
        strTodaySessionFinishDate = strTodaySessionDate
    end if

    strGoinClassDateTime = strTodaySessionDate & " " & strTodaySessionHour & ":" & intGoinClassMinute & ":00"
    strFinishClassDateTime = strTodaySessionFinishDate & " " & strTodaySessionFinishHour & ":" & intEndClassMinute & ":00"

    if ( (DateDiff("s", strSysNowDateTime, strGoinClassDateTime) > 0) OR (DateDiff("s", strSysNowDateTime, strFinishClassDateTime) < 0) ) then
        '導去上課頁面
	    Call alertGo("", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT)
	    response.end
    end if
else
    Call alertGo("", "class.asp", CONST_NOT_ALERT_THEN_REDIRECT)
	response.end
end if

dim TypeLib, NewGUID
Set TypeLib = CreateObject("Scriptlet.TypeLib")
NewGUID = TypeLib.Guid
Set TypeLib = Nothing
%>	
<div class="temp_contant">
    <div class="main_membox">
	    <iframe name="I1" style="text-align: center" src="check_joinnet_room_JR.asp?language=<%=request("language")%>&rand=<%=NewGUID%>" width="100%" height="250" marginwidth="1" marginheight="1" scrolling="no" border="0" align="center" frameborder="0" ></iframe>
    </div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>