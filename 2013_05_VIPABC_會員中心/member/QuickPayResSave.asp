<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp" -->

<html>
<head>
</head>
<meta http-equiv="Content-Language" content="utf-8">
<body>
<%
'http://www.vipabc.com.cn/program/member/QuickPayResSave.asp?charset=UTF-8&cupReserved=&exchangeDate=&exchangeRate=&merAbbr=%e4%b8%8a%e6%b5%b7%e9%ba%a6%e5%a5%87%e6%95%99%e8%82%b2%e4%bf%a1%e6%81%af%e5%92%a8%e8%af%a2%e6%9c%89%e9%99%90%e5%85%ac%e5%8f%b8%ef%bc%88VIPABC%ef%bc%89&merId=104310052005355&orderAmount=1000000&orderCurrency=156&orderNumber=00035930&qid=201208231523384332632&respCode=00&respMsg=Success!&respTime=20120823152545&settleAmount=1000000&settleCurrency=156&settleDate=0823&traceNumber=433263&traceTime=0823152338&transType=01&version=1.0.0&signMethod=MD5&signature=547d2d1db18dba3db683281855b74d64

dim charset : charset = trim(request("charset"))
dim cupReserved : cupReserved = trim(request("cupReserved"))
dim exchangeDate : exchangeDate = trim(request("exchangeDate"))
dim exchangeRate : exchangeRate = trim(request("exchangeRate"))
dim merAbbr : merAbbr = trim(request("merAbbr"))
dim merId : merId = trim(request("merId"))
dim orderAmount : orderAmount = trim(request("orderAmount"))
dim orderCurrency : orderCurrency = trim(request("orderCurrency"))
dim orderNumber : orderNumber = trim(request("orderNumber"))
dim qid : qid = trim(request("qid"))
dim respCode : respCode = trim(request("respCode"))
dim respMsg : respMsg = trim(request("respMsg"))
dim respTime : respTime = trim(request("respTime"))
dim settleAmount : settleAmount = trim(request("settleAmount"))
dim settleCurrency : settleCurrency = trim(request("settleCurrency"))
dim settleDate : settleDate = trim(request("settleDate"))
dim signMethod : signMethod = trim(request("signMethod"))
dim signature : signature = trim(request("signature"))
dim traceNumber : traceNumber = trim(request("traceNumber"))
dim traceTime : traceTime = trim(request("traceTime"))
dim transType : transType = trim(request("transType"))
dim version : version = trim(request("version"))
dim sql : sql = ""
dim arr_result : arr_result = 0 
dim update_result : update_result = 0  
dim update_result2 : update_result2 = 0 
dim paymentdetailsn : paymentdetailsn = ""
dim clientsn : clientsn = "" 

'charset[0] = "UTF-8"
'cupReserved[1] = "{cardNumber=621234************1&cardType=02}"
'exchangeDate[2] = ""
'exchangeRate[3] = ""
'merAbbr[4] = ""
'merId[5] = ""
'orderAmount[6] = "3050"
'orderCurrency[7] = "156"
'orderNumber[8] = "12030910"
'qid[9] = "201203091034220539962"
'respCode[10] = "00"
'respMsg[11] = "支付成功"
'respTime[12] = "20120309103910"
'settleAmount[13] = "3050"
'settleCurrency[14] = "156"
'settleDate[15] = "1201"
'signMethod[16] = "MD5"
'signature[17] = ""
'traceNumber[18] = "053996"
'traceTime[19] = "0309103422"
'transType[20] = "01"
'version[21] = "1.0.0"

if orderAmount <> "" then
    orderAmount = orderAmount / 100
end if
if settleAmount <> "" then
    settleAmount = settleAmount / 100
end if


dim emailcheck : emailcheck = "N"
dim word : word = ""

'組回傳字串start**************************************
if (respCode = "00") then
    word = "0@1008"
else
    word = "1@fail"
end if
word = word & "@_" & charset
word = word & "@_" & cupReserved
word = word & "@_" & exchangeDate
word = word & "@_" & exchangeRate
word = word & "@_" & merAbbr
word = word & "@_" & merId
word = word & "@_" & orderAmount
word = word & "@_" & orderCurrency
word = word & "@_" & orderNumber
word = word & "@_" & qid
word = word & "@_" & respCode
word = word & "@_" & respMsg
word = word & "@_" & respTime
word = word & "@_" & settleAmount
word = word & "@_" & settleCurrency
word = word & "@_" & settleDate
word = word & "@_" & signMethod
word = word & "@_" & signature
word = word & "@_" & traceNumber
word = word & "@_" & traceTime
word = word & "@_" & transType
word = word & "@_" & version
'response.write word & "<br />"
'組回傳字串end**************************************

dim intordernumber : intordernumber = 0
if (ordernumber <> "") then
    do while (left(ordernumber, 1) = "0")
           ordernumber = mid(ordernumber, 2)
    loop
end if

'response.write ordernumber & "<br />"

'尋找該筆資料start**************************************
sql = "SELECT Retcode, out_date, payment_detail_sn, Account_sn, Ordernumber FROM hitrust_record_all WHERE (sn =@sn)"
arr_result = excuteSqlStatementRead(sql, Array(ordernumber), CONST_VIPABC_RW_CONN)
dim Retcode : Retcode = ""
dim out_date : out_date = ""
dim contract_sn : contract_sn = ""

if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
		Retcode = arr_result(0,0)
		out_date = arr_result(1,0)
		paymentdetailsn = arr_result(2,0)
		clientsn = arr_result(3,0)
		contract_sn = arr_result(4,0)
		'response.write "Retcode = " & len(Retcode) & "<br />"
        if (isnull(Retcode)=true and respCode = "00") then
			 '----------------- 更新刷卡紀錄----------------------- Begin ------------------
             update_result = excuteSqlStatementWrite("Update hitrust_record_all SET Retcode = @Retcode, out_date = @out_date where sn = @sn", Array(word, now(), ordernumber), CONST_VIPABC_RW_CONN)
             'if (update_result < 0) then 
                 'Response.Write("error message = " & g_str_run_time_err_msg & "<br>")
             'end if
             set update_result = nothing
             '----------------- 更新刷卡紀錄----------------------- end ------------------
			 '----------------- 更新合約狀態isbuy=1----------------------- Begin ------------------
             update_result2 = excuteSqlStatementWrite("Update client_payment_detail SET isbuy = 1 where sn = @sn", Array(paymentdetailsn), CONST_VIPABC_RW_CONN)
             'if (update_result2 < 0) then 
                 'Response.Write("error message = " & g_str_run_time_err_msg & "<br>")
             'end if
             set update_result2 = nothing
             '----------------- 更新合約狀態isbuy=1----------------------- end ------------------

             emailcheck = "Y"
        end if
    end if
end if
set sql = nothing
set arr_result = nothing
'尋找該筆資料end**************************************

'寄送email start**************************************
if (emailcheck = "Y") then
    Call SendCardPaymentMail_test(clientsn, paymentdetailsn, ordernumber, "1001", "银联快捷(QuickPay)-北京")
    '接著更新合約開通
	Response.Redirect("http://www.vipabc.com/program/member/buy_access.asp?hdn_contract_sn=" & contract_sn)
end if			 
'寄送email end**************************************** 
%>
    <table align="center" border="1">
         <tr><td align="left">charset&nbsp;&nbsp;:</td><td align="left"><%=Charset%>&nbsp;</td></tr>
         <tr><td align="left">cupReserved&nbsp;:&nbsp;</td><td align="left"><%=CupReserved%>&nbsp;</td></tr>
         <tr><td align="left">exchangeDate&nbsp;:&nbsp;</td><td align="left"><%=ExchangeDate%>&nbsp;</td></tr>
         <tr><td align="left">exchangeRate&nbsp;&nbsp;:</td><td align="left"><%=ExchangeRate%>&nbsp;</td></tr>
         <tr><td align="left">merAbbr&nbsp;:&nbsp;</td><td align="left"><%=MerAbbr%>&nbsp;</td></tr>
         <tr><td align="left">merId&nbsp;:&nbsp;</td><td align="left"><%=MerId%>&nbsp;</td></tr>
         <tr><td align="left">orderAmount&nbsp;&nbsp;:</td><td align="left"><%=OrderAmount%>&nbsp;</td></tr>
         <tr><td align="left">orderCurrency&nbsp;:&nbsp;</td><td align="left"><%=OrderCurrency%>&nbsp;</td></tr>
         <tr><td align="left">orderNumber&nbsp;:&nbsp;</td><td align="left"><%=OrderNumber%>&nbsp;</td></tr>
         <tr><td align="left">qid&nbsp;&nbsp;:</td><td align="left"><%=Qid%>&nbsp;</td></tr>
         <tr><td align="left">respCode&nbsp;:&nbsp;</td><td align="left"><%=RespCode%>&nbsp;</td></tr>
         <tr><td align="left">respMsg&nbsp;:&nbsp;</td><td align="left"><%=RespMsg%>&nbsp;</td></tr>
         <tr><td align="left">respTime&nbsp;&nbsp;:</td><td align="left"><%=RespTime%>&nbsp;</td></tr>
         <tr><td align="left">settleAmount&nbsp;:&nbsp;</td><td align="left"><%=SettleAmount%>&nbsp;</td></tr>
         <tr><td align="left">settleCurrency&nbsp;:&nbsp;</td><td align="left"><%=SettleCurrency%>&nbsp;</td></tr>
         <tr><td align="left">settleDate&nbsp;&nbsp;:</td><td align="left"><%=SettleDate%>&nbsp;</td></tr>
         <tr><td align="left">signMethod&nbsp;:&nbsp;</td><td align="left"><%=SignMethod%>&nbsp;</td></tr>
         <tr><td align="left">signature&nbsp;:&nbsp;</td><td align="left"><%=Signature%>&nbsp;</td></tr>
         <tr><td align="left">traceNumber&nbsp;:&nbsp;</td><td align="left"><%=TraceNumber%>&nbsp;</td></tr>
         <tr><td align="left">traceTime&nbsp;&nbsp;:</td><td align="left"><%=TraceTime%>&nbsp;</td></tr>
         <tr><td align="left">traceType&nbsp;:&nbsp;</td><td align="left"><%=TransType%>&nbsp;</td></tr>
         <tr><td align="left">version&nbsp;:&nbsp;</td><td align="left"><%=Version%>&nbsp;</td></tr>
   </table>
    
   <table align="center" border="1">
         <tr><td align="left">emailcheck&nbsp;&nbsp;:</td><td align="left"><%=emailcheck%>&nbsp;</td></tr>
   </table>
</body>
</html>
