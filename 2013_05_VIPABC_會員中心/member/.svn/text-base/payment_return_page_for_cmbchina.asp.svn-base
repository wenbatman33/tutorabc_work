<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="CMBNETPAYMENT" content="China Merchants Bank">
    <title>招商银行</title>
</head>
<body>
<%
on error resume next 

'取值Y(成功)或N(失败)
Succeed = trim(request("Succeed"))

'商户号，6位长数字，由银行在商户开户时确定；
CoNo = trim(request("CoNo"))

'定单号(由支付命令送来)
BillNo = trim(request("BillNo"))

'实际支付金额(由支付命令送来)
Amount = trim(request("Amount"))

'交易日期(主机交易日期)
strDate = trim(request("Date"))

'银行通知用户的支付结果消息。信息的前38个字符格式为：4位分行号＋6位商户号＋8位银行接受交易的日期＋20位银行流水号；可以利用交易日期＋银行流水号对该定单进行结帐处理；
Msg = trim(request("Msg"))

MerchantPara = trim(request("MerchantPara"))
Signature = trim(request("Signature"))

dim strMerId : strMerId = CONST_CREDIT_CARD_CMBCHINA_PAY_SUCCEED_CODE '商店代碼
dim strOrderNo : strOrderNo = BillNo '訂單編號
dim strTransDate : strTransDate = left(strDate,8)
dim intAmount : intAmount = Amount
dim strCurrencyCode : strCurrencyCode = "156"
dim strTransType : strTransType = MerchantPara
dim strStatusResult : strStatusResult = Succeed
dim strGateId : strGateId = CoNo
dim intDetailSn : intDetailSn = "" 'dsn
dim intRequestDetailSn : intRequestDetailSn = MerchantPara 'dsn
dim strScore : strScore = "CMBChina" '資料來源
dim intContractSn : intContractSn = "" '合約編號
dim strReturnValue : strReturnValue = "" '寫入hitrust_record_all中
dim strOrgValue : strOrgValue = "" '儲存原始回傳刷卡資料
dim intHitrustSn : intHitrustSn = "" 'hitrust_record_all.sn
dim intStatus : intStatus = ""
dim strResultStat : strResultStat = ""

'塞資料的相關設定值
dim strSql : strSql = ""
dim arrParam : arrParam = null '傳值陣列
dim arrSqlResult : arrSqlResult = null '接回來的陣列值

'Email的部份
dim strEmailFlag : strEmailFlag = ""
dim strChangeFlag : strChangeFlag = "" '是否有開通成功
dim arrEmailFromParas
dim arrEmailToParas
dim arrSetParas
dim arrOptParas
dim strMailSubject : strMailSubject = ""
dim exec_result 
dim strMailUrl '信件的url

'log紀錄部分
dim objLogOpt : objLogOpt = null
dim strFileName : strFileName = "" 'log檔名 
dim arrColumn : arrColumn = "" '資料欄位名稱
dim arrData : arrData = "" 'log資料
dim intCount : intCount = 0 '欄位筆數
dim strData : strData = "" 'log資料串
dim strUserIP : strUserIP = ""

'TestCaseConfig設定
dim bolTestCase : bolTestCase = false
dim intTestMode : intTestMode = 3
dim bolFindClientSn : bolFindClientSn = true

'==================== 判斷成功/失敗 Start ====================
if ( strStatusResult = "Y" ) then
    strResultStat = "0"
    intStatus = CONST_CREDIT_CARD_CMBCHINA_PAY_SUCCEED_CODE
elseif ( strStatusResult = "N" ) then
    strResultStat = "1@_" & Msg
else
    strResultStat = "1@_" & Msg
end if
'==================== 判斷成功/失敗 End ====================

if ( Not isEmptyOrNull(strOrderNo) ) then 
    intHitrustSn = CLng(Right(strOrderNo,5))
end if 

if ( IsNumeric(intAmount) ) then
	'intAmount  = intAmount / CONST_CHINA_TO_TAIWAN_EXCHANGE_RATE
    intAmount  = replace(intAmount,".00","")
end if 

strReturnValue = strResultStat & "@" & strMerId & "@_" & strOrderNo & "@_" &strTransDate & "@_" & intAmount & "@_" & strCurrencyCode & "@_" & strTransType & "@_" & strGateId & "@_VIPABC" 

'==================== 寫入刷卡記錄 Start ====================
dim intBolLog : intBolLog = ""
arrSqlResult = excuteSqlWriteHitrustRecord(intHitrustSn, strScore, strReturnValue)
'Response.Write("錯誤訊息sql....:" & g_strSql_statement_for_debug & "<br>")
if ( isSelectQuerySuccess(arrSqlResult) ) then
	if ( Ubound(arrSqlResult) >= 0 ) then '有資料
		intContractSn = arrSqlResult(0) '合約編號
		intDetailSn = arrSqlResult(1) 'dsn
        intBolLog = arrSqlResult(2) '是否為重復筆數
	end if 
else
end if 
'==================== 寫入刷卡記錄 End ====================

'==================== TestCase Start ====================
if ( true = bolTestCase ) then
    '1 : 客戶編號遺漏(補回客戶編號功能關閉)
    '2 : 合約編號遺漏
    '3 : 客戶編號 & 合約編號遺漏
    '4 : 補回客戶編號功能開啟
    select case (intTestMode)
        case 1
            g_var_client_sn = ""
            bolFindClientSn = false
        case 2
            g_var_client_sn = "223543"
            intContractSn = ""
            bolFindClientSn = true
        case 3
            g_var_client_sn = ""
            intContractSn = ""
            bolFindClientSn = false
        case 4
            g_var_client_sn = ""
            bolFindClientSn = true
    end select
end if
'==================== TestCase End ====================

'if ( g_var_client_sn = "223543" ) then
%>
<!--#include virtual="/program/member/choosePaymentAccount.asp" -->
<%
'end if

dim strPlamName : strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_CMBCHINA_NAME

if ( true = bolBJChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_BJ_CMBCHINA_NAME
elseif ( true = bolSHChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_SH_CMBCHINA_NAME
end if

'==================== Log紀錄:[start] Start ====================
Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")
if ( true = bolTestCase ) then
    strFileName = "log_SendCard_4_test-start"
else
    strFileName = "log_SendCard_4-start"
end if
strUserIP = ""
strUserIP = getIpAddress()
strData = ""
arrColumn = array("HitrustRecordAllSn", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "PayType", "InputTime", "ScoreSn", "ClientName", "SucceedCode", "ClientIP")
arrData = array(intHitrustSn, intDetailSn, g_var_client_sn, intContractSn, intAmount, strPlamName, now(), strScore, str_user_name, intStatus, strUserIP)
for intCount = 0 to UBound(arrColumn)
    if ( intCount > 0 ) then
        strData = strData & ","
    end if
    strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
next
strData = strData & vbCrLF
Call objLogOpt.WriteLog(strFileName, strData)
Set objLogOpt = nothing
'==================== Log紀錄:[start] End ====================

'==================== 補回g_var_client_sn Start ====================
if ( true = bolFindClientSn ) then
    dim intCookiesSn : intCookiesSn = Request.Cookies("client_sn")
    if ( isEmptyOrNull(g_var_client_sn) ) then
        'cookie補回session機制 
        if ( Not isEmptyOrNull(intCookiesSn) ) then
            session("client_sn") = intCookiesSn
            g_var_client_sn = session("client_sn")
        else
            '根據資料庫補回
            arrSqlResult = getClientCardPaymentDataNew(intContractSn, intDetailSn, CONST_VIPABC_RW_CONN)
            'Response.Write("錯誤訊息sql....:" & g_strSql_statement_for_debug & "<br>")
	        if ( isSelectQuerySuccess(arrSqlResult) ) then
		        if ( Ubound(arrSqlResult) >= 0 ) then '有資料
			        g_var_client_sn = arrSqlResult(0,0)
		        else
		        end if
	        else
	        end if
        end if
    end if
end if
'==================== 補回g_var_client_sn End ====================

'==================== 寄系統信 Start ====================
if ( 0 = intBolLog ) then '非重覆封包
    if ( Not isEmptyOrNull(intDetailSn) ) then
        strChangeFlag = changeCardPaymentDate(intContractSn, intDetailSn)    
        if ( Not isEmptyOrNull(intContractSn) ) then
            strEmailFlag = SendCardPaymentMailNew(intContractSn, intDetailSn, intHitrustSn, intStatus, strPlamName, CONST_VIPABC_RW_CONN)
        elseif ( Not isEmptyOrNull(g_var_client_sn) ) then
            strEmailFlag = SendCardPaymentMail_test(g_var_client_sn, intDetailSn, intHitrustSn, intStatus, strPlamName)
        else
            '==================== 寄信功能 Start ====================
            dim strClientCName : strClientCName = "银行回传数据残缺，请至银行后台查询此笔刷卡"
            strMailSubject = "VIPABC 客户(" & strClientCName & ")在线刷卡讯息"
            strMailUrl = "http://" & CONST_VIPABC_WEBHOST_NAME & "/program/mail_template/card_payment/card_payment_success_test.asp"
            strMailUrl = strMailUrl & "?str_pay_platform=" & escape(strPlamName) & "&mail_client_name=" & escape(strClientCName) 
            strMailUrl = strMailUrl & "&mail_porderno=" & strOrderNo & "&mail_contract_sn=" & intContractSn & "&mail_dsn=" & intDetailSn
            strMailUrl = strMailUrl & "&mail_staff_ename=" & strSalesEName & "&mail_product_name=" & escape(strProductName) 
            strMailUrl = strMailUrl & "&mail_pay_money="& intAmount & "&mail_status=" & intStatus & "&mail_pay_time=" & escape(strDate)
	        if ( true = bolTestCase ) then
                strSalesEmail = "arthurwu@tutorabc.com,beautitsou@tutorabc.com," & strSalesLearderEmail
            else
                strSalesEmail = strSalesEmail & "," & CONST_EMAIL_CREDIT_CARD_PAYMENT & strSalesLearderEmail
            end if 
	        arrEmailFromParas = Array(CONST_MAIL_ADDRESS_SERVICES_VIPABC, "", strMailSubject, "", "", strMailUrl)
	        arrEmailToParas = Array(strSalesEmail, strSalesEName)
	        arrSetParas = Array(CONST_MAIL_SERVER_CLIENT, CONST_EMAIL_CHARSET)
	        arrOptParas = Array(true)
	
	        exec_result = sendEmail(arrEmailFromParas, arrEmailToParas, arrSetParas, arrOptParas, CONST_TUTORABC_WEBSITE)

	        if (exec_result <> CONST_FUNC_EXE_SUCCESS) then
	            'E-mail 錯誤訊息
	        else
	        end if	
            '==================== 寄信功能 End ====================
        end if
    else
        '==================== Log紀錄:[fail] Start ====================
        Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")
        if ( true = bolTestCase ) then
            strFileName = "log_SendCard_4_test-fail"
        else
            strFileName = "log_SendCard_4-fail"
        end if
        strUserIP = ""
        strUserIP = getIpAddress()
        strData = ""
        arrColumn = array("HitrustRecordAllSn", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "PayType", "InputTime", "ScoreSn", "ClientName", "SucceedCode", "ClientIP")
        arrData = array(intHitrustSn, intDetailSn, g_var_client_sn, intContractSn, intAmount, strPlamName, now(), strScore, str_user_name, intStatus, strUserIP)
        for intCount = 0 to UBound(arrColumn)
            if ( intCount > 0 ) then
                strData = strData & ","
            end if
            strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
        next
        strData = strData & vbCrLF
        Call objLogOpt.WriteLog(strFileName, strData)
        Set objLogOpt = nothing
        '==================== Log紀錄:[fail] End ====================
    end if	
end if
'==================== 寄系統信 End ====================

'==================== Log紀錄:[end] Start ====================
Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")
if ( true = bolTestCase ) then
    strFileName = "log_SendCard_4_test-end"
else
    strFileName = "log_SendCard_4-end"
end if
strUserIP = ""
strUserIP = getIpAddress()
strData = ""
arrColumn = array("HitrustRecordAllSn", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "PayType", "InputTime", "ScoreSn", "ClientName", "SucceedCode", "ClientIP")
arrData = array(intHitrustSn, intDetailSn, g_var_client_sn, intContractSn, intAmount, strPlamName, now(), strScore, str_user_name, intStatus, strUserIP)
for intCount = 0 to UBound(arrColumn)
    if ( intCount > 0 ) then
        strData = strData & ","
    end if
    strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
next

strData = strData & vbCrLF
Call objLogOpt.WriteLog(strFileName, strData)
Set objLogOpt = nothing
'==================== Log紀錄:[end] End ====================
%>
    <table border="0" cellspacing="5" cellpadding="5" width="100%">
        <tr>
            <td>
                <img id="img_ajax_loader" src="/lib/javascript/JQuery/images/ajax-loader.gif" border="0" />
                <a style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 15px; text-decoration: none;">刷卡处理中… 请勿关闭此页面<br />(易造成刷卡失败)</a>
            </td>
        </tr>
    </table>
    <% 
    if ( CONST_CREDIT_CARD_CMBCHINA_PAY_SUCCEED_CODE = intStatus ) then
        'response.redirect "buy_access.asp?hdn_contract_sn=" & intContractSn & "&hdn_dsn_sn=" & intDetailSn & "&succ_code=" & intStatus & "&store_code=" & escape(strPlamName)
    %>
    <form name="frm_payment_return" id="frm_payment_return" action="buy_access.asp" method="post">
    <input name="hdn_contract_sn" type="hidden" value="<%=intContractSn%>" />
    <input name="hdn_dsn_sn" type="hidden" value="<%=intDetailSn%>" />
    <input name="succ_code" type="hidden" value="<%=intStatus%>" />
    <input name="store_code" type="hidden" value="<%=strPlamName%>" />
    <script type="text/javascript" language="javascript">
        document.frm_payment_return.submit();
    </script>
    </form>
    <% else 
        'response.redirect "buy_fail.asp?hdn_contract_sn=" & intContractSn & "&hdn_dsn_sn=" & intDetailSn
    %>
    <form name="frm_payment_return" id="frm_payment_return" action="buy_fail.asp" method="post">
    <input name="hdn_contract_sn" type="hidden" value="<%=intContractSn%>" />
    <input name="hdn_dsn_sn" type="hidden" value="<%=intDetailSn%>" />
    <script type="text/javascript" language="javascript">
        document.frm_payment_return.submit();
    </script>
    </form>
    <% end if %>
</body>
</html>
