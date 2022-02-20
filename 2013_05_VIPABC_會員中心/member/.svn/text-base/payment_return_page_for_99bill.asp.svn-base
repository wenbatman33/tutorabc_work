<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="Content-Language" content="utf-8">
    <title>快钱付款网关</title>
</head>
<body>
<%

'获取人民币网关账户号
merchantAcctId = trim(request("merchantAcctId"))

'设置人民币网关密钥
''区分大小写
key = CONST_CREDIT_CARD_99BILL_KEY '測試
'key="J4CNSQHLM7JAYRRB" '正式

'获取网关版本.固定值
''快钱会根据版本号来调用对应的接口处理程序。
''本代码版本号固定为v2.0
version = trim(request("version"))

'获取语言种类.固定选择值。
''只能选择1、2、3
''1代表中文；2代表英文
''默认值为1
language = trim(request("language"))

'签名类型.固定值
''1代表MD5签名
''当前版本固定为1
signType = trim(request("signType"))

'获取支付方式
''值为：10、11、12、13、14
''00：组合支付（网关支付页面显示快钱支持的各种支付方式，推荐使用）10：银行卡支付（网关支付页面只显示银行卡支付）.11：电话银行支付（网关支付页面只显示电话支付）.12：快钱账户支付（网关支付页面只显示快钱账户支付）.13：线下支付（网关支付页面只显示线下支付方式）.14：B2B支付（网关支付页面只显示B2B支付，但需要向快钱申请开通才能使用）
payType = trim(request("payType"))

'获取银行代码
''参见银行代码列表
bankId = trim(request("bankId"))

'获取商户订单号
orderId = trim(request("orderId"))

'获取订单提交时间
''获取商户提交订单时的时间.14位数字。年[4位]月[2位]日[2位]时[2位]分[2位]秒[2位]
''如：20080101010101
orderTime = trim(request("orderTime"))

'获取原始订单金额
''订单提交到快钱时的金额，单位为分。
''比方2 ，代表0.02元
orderAmount = trim(request("orderAmount"))

'获取快钱交易号
''获取该交易在快钱的交易号
dealId = trim(request("dealId"))

'获取银行交易号
''如果使用银行卡支付时，在银行的交易号。如不是通过银行支付，则为空
bankDealId = trim(request("bankDealId"))

'获取在快钱交易时间
''14位数字。年[4位]月[2位]日[2位]时[2位]分[2位]秒[2位]
''如；20080101010101
dealTime = trim(request("dealTime"))

'获取实际支付金额
''单位为分
''比方 2 ，代表0.02元
payAmount = trim(request("payAmount"))

'获取交易手续费
''单位为分
''比方 2 ，代表0.02元
fee = trim(request("fee"))

'获取扩展字段1
ext1 = trim(request("ext1"))

'获取扩展字段2
ext2 = trim(request("ext2"))

'获取处理结果
''10代表 成功11代表 失败
payResult = trim(request("payResult"))

'获取错误代码
''详细见文档错误代码列表
errCode = trim(request("errCode"))

'获取加密签名串
signMsg = trim(request("signMsg"))

dim str_merid : str_merid = CONST_CREDIT_CARD_99BILL_PAY_SUCCEED_CODE '商店代碼
dim str_orderno : str_orderno = orderId	'訂單編號
dim str_transdate : str_transdate = LEFT(dealTime,8)
dim str_amount : str_amount = payAmount
dim str_currencycode : str_currencycode = "156"
dim str_transtype : str_transtype = payType
dim strStatusResult : strStatusResult = payResult

if ( isEmptyOrNull(strStatusResult) ) then
    strStatusResult = 0
end if

if ( strStatusResult = 10 ) then
    str_Rstatus = "0"
    str_status = CONST_CREDIT_CARD_99BILL_PAY_SUCCEED_CODE
elseif ( strStatusResult = 11 ) then
    str_Rstatus = "1@_" & errCode
else
    str_Rstatus = "1@_" & errCode
end if

dim str_GateId : str_GateId = dealId

if ( Not isEmptyOrNull(str_orderno) ) then
dim str_hitrust_sn : str_hitrust_sn = CLng( Right(str_orderno,5))  'hitrust_record_all.sn
end if

dim str_dsn_sn : str_dsn_sn = "" 'dsn
dim str_score : str_score = "99Bill"	'資料來源
dim str_contract_sn : str_contract_sn = "" 		'合約編號
dim str_return_value : str_return_value = ""	'寫入hitrust_record_all中
dim str_org_value : str_org_value = "" '儲存原始回傳刷卡資料
dim str_client_sn :  str_client_sn = "" 'client_sn

'dim str_message : str_message = Request("Message") '應答碼的中文含義
'塞資料的相關設定值
dim str_sql : str_sql = ""
dim str_addnew_sql : str_addnew_sql  = ""
dim var_arr 		'傳excuteSqlStatementRead的陣列
dim arr_result		'接回來的陣列值
dim str_tablecol 	'table欄位
dim str_tablevalue 	'table值

'mail的部份
dim str_email_flag
dim str_change_flag	'是否有開通成功

'log紀錄
dim objLogOpt : objLogOpt = null
dim strFileName : strFileName = "" 'log檔名 
dim arrColumn : arrColumn = "" '資料欄位名稱
dim arrData : arrData = "" 'log資料
dim intCount : intCount = 0 '欄位筆數
dim strData : strData = "" 'log資料串
dim strUserIP : strUserIP = ""

if ( IsNumeric(str_amount)) then
    str_amount  = str_amount / CONST_CHINA_TO_TAIWAN_EXCHANGE_RATE
end if 

'if ( g_var_client_sn = "223543" ) then
    intHitrustSn = str_hitrust_sn
    intContractSn = str_contract_sn   
%>
<!--#include virtual="/program/member/choosePaymentAccount.asp" -->
<%
'end if

dim strPlamName : strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_99BILL_NAME

if ( true = bolBJChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_BJ_99BILL_NAME
elseif ( true = bolSHChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_SH_99BILL_NAME
end if

str_return_value = str_Rstatus & "@" & str_merid & "@_" & str_orderno & "@_" &str_transdate & "@_" & str_amount & "@_" & str_currencycode & "@_" & str_transtype & "@_" & str_GateId & "@_VIPABC" 

'log紀錄:log_SendCard_4-start
Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")

strFileName = "log_SendCard_4-start"
strUserIP = ""
strUserIP = getIpAddress()
strData = ""
arrColumn = array("HitrustRecordAllSn", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "PayType", "InputTime", "ScoreSn", "ClientName", "SucceedCode", "ClientIP")
arrData = array(str_hitrust_sn, str_dsn_sn, g_var_client_sn, str_contract_sn, str_amount, strPlamName, now(), str_score, str_user_name, str_status, strUserIP)
for intCount = 0 to UBound(arrColumn)
    if ( intCount > 0 ) then
        strData = strData & ","
    end if
    strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
next

strData = strData & vbCrLF
Call objLogOpt.WriteLog(strFileName, strData)
Set objLogOpt = nothing

'cookie補強session機制
if ( isEmptyOrNull(g_var_client_sn) ) then
    session("client_sn") = Request.Cookies("client_sn")
    g_var_client_sn = session("client_sn")
end if

'==================== 捉回client_sn Start ====================
'如果還有session就抓session值
str_client_sn = g_var_client_sn
'如果沒有就去資料庫撈資料
if ( isEmptyOrNull(str_client_sn) ) then	
    str_sql = " SELECT TOP 1 payment_detail_sn FROM hitrust_record_all WHERE sn = @sn "
    var_arr = array(str_hitrust_sn)
    arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
    'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
    if ( isSelectQuerySuccess(arr_result) ) then
        '如果有查尋hitrust_record_all資料
        if ( Ubound(arr_result) >= 0 ) then 
    	    '回傳detail_sn
		    str_dsn_sn = arr_result(0,0)
		    '根據detail_sn查出client_sn
		    str_sub_sql = " SELECT TOP 1 client_sn FROM contract_pay WHERE dsn = @dsn "	
            var_sub_arr = array(str_dsn_sn)
            arr_sub_result = excuteSqlStatementRead(str_sub_sql, var_sub_arr, CONST_VIPABC_RW_CONN)
            'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
            if ( isSelectQuerySuccess(arr_sub_result) ) then
	            '有資料
	            if ( Ubound(arr_sub_result) >= 0 ) then 
		            str_client_sn = arr_sub_result(0,0)			
	            else
		            '錯誤訊息
	            end if
            end if
	    else
		    '錯誤訊息
	    end if 
    end if 
end if
'==================== 捉回client_sn End ====================

'==================== 寫刷卡記錄到 hitrust_record_all Start ====================
'str_score：是誰Call這一頁：pageRet：成功 / 另一個：失敗 (需另開一個欄位來存放)  str_org_value
dim intBolLog : intBolLog = ""
arr_result = excuteSqlWriteHitrustRecord(str_hitrust_sn, str_score, str_return_value)
if ( isSelectQuerySuccess(arr_result) ) then
    if ( Ubound(arr_result) >= 0 ) then '有資料
	    str_contract_sn = arr_result(0) '合約編號
	    str_dsn_sn = arr_result(1) 'dsn
        intBolLog = arr_result(2) '是否為重復筆數
    end if 
    else
end if 
'==================== 寫刷卡記錄到 hitrust_record_all End ====================


'==================== 付款成功 Start ====================
if ( str_status = CONST_CREDIT_CARD_99BILL_PAY_SUCCEED_CODE ) then
    str_change_flag = changeCardPaymentDate(str_contract_sn, str_dsn_sn)
else	
end if 

if ( 0 = intBolLog ) then '非重覆封包
    if ( str_status = CONST_CREDIT_CARD_99BILL_PAY_SUCCEED_CODE AND Not isEmptyOrNull(str_client_sn) AND Not isEmptyOrNull(str_dsn_sn) ) then
	    'str_email_flag = SendCardPaymentMail(str_client_sn, str_dsn_sn, str_status, CONST_PAYMENT_PLATFORM_VIPABC_99BILL_NAME)
        str_email_flag = SendCardPaymentMail_test(str_client_sn, str_dsn_sn, str_hitrust_sn, str_status, strPlamName)
    else
        '寄失敗信
        'str_email_flag = SendCardPaymentMail_test(str_contract_sn, str_dsn_sn, str_status, CONST_PAYMENT_PLATFORM_VIPABC_99BILL_NAME, CONST_VIPABC_RW_CONN)
    
        'log紀錄:log_SendCard_4-fail
        Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")

        strFileName = "log_SendCard_4-fail"
        strUserIP = ""
        strUserIP = getIpAddress()
        strData = ""
        arrColumn = array("HitrustRecordAllSn", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "PayType", "InputTime", "ScoreSn", "ClientName", "SucceedCode", "ClientIP")
        arrData = array(str_hitrust_sn, str_dsn_sn, g_var_client_sn, str_contract_sn, str_amount, strPlamName, now(), str_score, str_user_name, str_status, strUserIP)
        for intCount = 0 to UBound(arrColumn)
            if ( intCount > 0 ) then
                strData = strData & ","
            end if
            strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
        next

        strData = strData & vbCrLF
        Call objLogOpt.WriteLog(strFileName, strData)
        Set objLogOpt = nothing
    end if
end if

'log紀錄:log_SendCard_4-end
Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")

strFileName = "log_SendCard_4-end"
strUserIP = ""
strUserIP = getIpAddress()
strData = ""
arrColumn = array("HitrustRecordAllSn", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "PayType", "InputTime", "ScoreSn", "ClientName", "SucceedCode", "ClientIP")
arrData = array(str_hitrust_sn, str_dsn_sn, g_var_client_sn, str_contract_sn, str_amount, strPlamName, now(), str_score, str_user_name, str_status, strUserIP)
for intCount = 0 to UBound(arrColumn)
    if ( intCount > 0 ) then
        strData = strData & ","
    end if
    strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
next

strData = strData & vbCrLF
Call objLogOpt.WriteLog(strFileName, strData)
Set objLogOpt = nothing

%>
<table border="0" cellspacing="5" cellpadding="5" width="100%">
    <tr>
        <td>
            <img id="img_ajax_loader" src="/lib/javascript/JQuery/images/ajax-loader.gif" border="0" />
            <a style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 15px; text-decoration: none;">刷卡处理中… 请勿关闭此页面<br />(易造成刷卡失败)</a>
        </td>
    </tr>
</table>
<% if ( str_status = CONST_CREDIT_CARD_99BILL_PAY_SUCCEED_CODE ) then 
    response.Redirect "buy_access.asp?hdn_contract_sn=" & str_contract_sn & "&hdn_dsn_sn=" & str_dsn_sn & "&succ_code=" & str_status & "&store_code=" & escape(strPlamName)
%>
<!--
<form name="frm_payment_return" id="frm_payment_return" action="buy_access.asp" method="post">
<input name="hdn_contract_sn" type="hidden" value="<%=str_contract_sn%>" />
<input name="hdn_dsn_sn" type="hidden" value="<%=str_dsn_sn%>" />
<input name="sendmail" type="hidden" value="" />
<input name="succ_code" type="hidden" value="<%=str_status%>" />
<input name="store_code" type="hidden" value="<%=escape(strPlamName)%>" />
<script type="text/javascript" language="javascript">
    document.frm_payment_return.submit();
</script>
</form>
-->
<% else 
    response.Redirect "buy_fail.asp?hdn_contract_sn=" & str_contract_sn & "&hdn_dsn_sn=" & str_dsn_sn
%>
<!--
<form name="frm_payment_return" id="frm_payment_return" action="buy_fail.asp" method="post">
<input name="hdn_contract_sn" type="hidden" value="<%=str_contract_sn%>" />
<input name="hdn_dsn_sn" type="hidden" value="<%=str_dsn_sn%>" />
<script type="text/javascript" language="javascript">
    document.frm_payment_return.submit();
</script>
</form>
-->
<% end if %>
</body>
</html>
