<!--#include virtual="/lib/include/global.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>
<%
'dim strTransBeijingDomain : strTransBeijingDomain = CONST_TRANS_BEIJING_DOMAIN
dim strBankPostUrl : strBankPostUrl = ""

'if ( "1" = strTransBeijingDomain ) then
    'strBankPostUrl = "http://www.vipabc.com/program/member/cmbchina/cmbchina_payment.asp"
'else
    strBankPostUrl = "https://netpay.cmbchina.com/netpayment/BaseHttp.dll?PrePayC2"
'end if

Dim str_client_cname  : str_client_cname = request("txt_cname") '客戶中文姓名
Dim int_product_sn : int_product_sn =  session("product_sn") '購買產品的序號(cfg_product.sn)

Dim str_client_address : str_client_address = Request("txt_client_address") '客戶住址
Dim str_client_tel_code : str_client_tel_code = Request("txt_tel_code") '手機號碼前四碼
Dim str_client_tel : str_client_tel  = Request("txt_tel") '手機號碼後六碼
Dim str_client_email : str_client_email  = Request("txt_email") '客戶email
Dim str_client_ip : str_client_ip = getIpAddress '客戶ip
Dim str_iscustom : str_iscustom = Request("iscustom") '是否購買
Dim str_client_pay_cash : str_client_pay_cash = Request("pay_cash") '價錢

Dim int_dsn : int_dsn = request("dsn") '付款相關table client_payment_detail.sn
Dim str_product_name : str_product_name = "VIPABC" '購買產品名稱
Dim int_valid_month  : int_valid_month = 1 '有效期限
Dim int_contract_sn : int_contract_sn = 0 '合約編號
Dim int_lead_sn 'lead_sn
Dim str_ostr '買的系統

'抓取欲購買產品名稱及期限
Dim product_sql : product_sql = ""

'抓取欲購買合約之相關資訊
Dim contract_sql : contract_sql = ""

Dim hitrust_record_all_sn : hitrust_record_all_sn = "" 'hitrust_record_all.sn

'塞資料的相關設定值
Dim var_arr '傳excuteSqlStatementRead的陣列
Dim arr_result '接回來的陣列值
Dim str_tablecol 'table欄位
Dim str_tablerow 'table列數

if ( int_product_sn = "") then
	int_product_sn  = request("productsn")
end if

'判斷參數是否存在 如果沒有 導回首頁
if ( int_product_sn = "" OR int_dsn = "" OR g_var_client_sn = "" ) then
	'如果session 掉了 導回首頁
	alertGo "网页逾时，请重新登入","http://"&CONST_VIPABC_WEBHOST_NAME&"/index.asp",CONST_ALERT_THEN_REDIRECT
end if

'----------------- 抓取欲購買產品名稱及期限----------------------- start -----------------
product_sql = " SELECT Cname, valid_month FROM dbo.cfg_product WHERE (sn =@product_sn) "
var_arr = Array(int_product_sn)
arr_result = excuteSqlStatementRead(product_sql,var_arr,CONST_VIPABC_RW_CONN)

if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
		str_product_name = arr_result(0,0)
		int_valid_month = arr_result(1,0)
	else
		str_product_name="VIPABC"
		int_valid_month=1
	end if
end if
set product_sql = nothing
set var_arr = nothing
set arr_result = nothing
'----------------抓取欲購買產品名稱及期限 ---------- End -------------------

'----------------- 抓取欲購買合約之相關資訊----------------------- start ------------------
contract_sql = " SELECT client_payment_record.contract_sn,client_temporal_contract.lead_sn "
contract_sql = contract_sql & " FROM client_temporal_contract  "
contract_sql = contract_sql & " INNER JOIN client_payment_record  "
contract_sql = contract_sql & " ON client_temporal_contract.sn = client_payment_record.contract_sn "
contract_sql = contract_sql & " INNER JOIN client_payment_detail ON client_payment_record.sn = client_payment_detail.payrecord_sn "
contract_sql = contract_sql & " WHERE (client_payment_detail.payment_mode in (23,34)) "
contract_sql = contract_sql & " AND (client_payment_detail.sn =@g_int_dsn )   "
contract_sql = contract_sql & " AND (client_temporal_contract.client_sn = @client_sn) "
var_arr = Array(int_dsn,g_var_client_sn)
arr_result = excuteSqlStatementRead(contract_sql,var_arr,CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
		int_contract_sn = arr_result(0,0)
		int_lead_sn = arr_result(1,0)
	else
		str_product_name="VIPABC"
		int_valid_month=1
	end if
end if

set contract_sql = nothing
set var_arr = nothing
set arr_result = nothing
'----------------- 抓取欲購買合約之相關資訊----------------------- End ------------------

'----------------- 更新client_basic居住地址----------------------- Begin ------------------
var_arr = Array(str_client_address,g_var_client_sn)
arr_result = excuteSqlStatementWrite(" UPDATE client_basic SET client_address = @client_address WHERE sn = @g_int_client_sn ", var_arr, CONST_VIPABC_RW_CONN)

if (arr_result >= 0) then 
	'Response.Write("更新影響筆數 " & invoice_effected_row & "<br>")
else
	Response.Write(getWord("CONTRACT_AGREE_2") & g_str_run_time_err_msg & "<br>")
end if

set var_arr = nothing
set arr_result = nothing
'----------------- 更新client_basic居住地址----------------------- end ------------------

'----------------- 更新lead_basic居住地址----------------------- Begin ------------------
var_arr = Array(str_client_address,int_lead_sn)
arr_result = excuteSqlStatementWrite(" UPDATE lead_basic SET client_address = @client_address WHERE sn = @g_int_lead_sn ", var_arr, CONST_VIPABC_RW_CONN)

if (arr_result >= 0) then 
	'Response.Write("更新影響筆數 " & invoice_effected_row & "<br>")
else
	Response.Write(getWord("CONTRACT_AGREE_2")& g_str_run_time_err_msg & "<br>")
end if
set var_arr = nothing
set arr_result = nothing
'----------------- 更新lead_basic居住地址----------------------- end ------------------

'----------------- 新增online_payment_data ----------------------- start ------------------
str_tablecol = "account_sn,cname,tel_code,tel,email,cip,cmoney,client_sn"
str_tablerow = "@account_sn,@cname,@tel_code,@tel,@email,@IP,@cmoney,@client_sn"
var_arr = Array(int_contract_sn,str_client_cname,str_client_tel_code,str_client_tel,str_client_email,getIpAddress,str_client_pay_cash,g_var_client_sn)

arr_result = excuteSqlStatementWrite(" INSERT INTO online_payment_data ("& str_tablecol &") VALUES ("& str_tablerow &") ", var_arr, CONST_VIPABC_RW_CONN)
if ( arr_result >= 0 ) then 
	'Response.Write("更新影響筆數 " & invoice_effected_row & "<br>")
else
	Response.Write(getWord("CONTRACT_AGREE_2")& g_str_run_time_err_msg & "<br>")
end if

set var_arr = nothing
set arr_result = nothing
'----------------- 新增online_payment_data ----------------------- end ------------------

'----------------- 新增hitrust_record_all Begin ----------------------- start ------------------
IF ( not isEmptyOrNull(int_product_sn) ) Then
	str_ostr = str_product_name
Else
	str_ostr = "VIPABC"
End IF

'if ( g_var_client_sn = "223543" ) then
    intContractSn = int_contract_sn
%>
<!--#include virtual="/program/member/choosePaymentAccount.asp" -->
<% 
'end if 

Dim int_comp_type : int_comp_type = CONST_BANK_PLATFORM_PARAMETER_CMBCHINA '1=網際威信 2=華南銀行 3=陸網銀聯 6.快钱

if ( true = bolBJChannel ) then
    int_comp_type = CONST_BANK_PLATFORM_PARAMETER_BJ_CMBCHINA
elseif ( true = bolSHChannel ) then
    int_comp_type = CONST_BANK_PLATFORM_PARAMETER_CMBCHINA
end if

dim strPlamName : strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_CMBCHINA_NAME

if ( true = bolBJChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_BJ_CMBCHINA_NAME
elseif ( true = bolSHChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_SH_CMBCHINA_NAME
end if

str_tablecol = "Account_sn,Storeid,Ordernumber,comp_type,Orderdesc,Amount"
str_tablecol = str_tablecol & ",Currency,Depositflag,Queryflag,payment_detail_sn" 'card_serve_bank
str_tablerow ="@Account_sn,@Storeid,@Ordernumber,@comp_type,@Orderdesc,@Amount"
str_tablerow = str_tablerow & ",@Currency,@Depositflag,@Queryflag,@payment_detail_sn" ',@card_serve_bank
var_arr = Array(int_contract_sn, str_MerId, int_contract_sn, int_comp_type, str_ostr, str_client_pay_cash, "RMB", "1", "0", int_dsn) ',"236"

arr_result = excuteSqlStatementScalar(" INSERT INTO hitrust_record_all(" & str_tablecol & ") VALUES (" & str_tablerow & ") SELECT SCOPE_IDENTITY()", var_arr,CONST_VIPABC_RW_CONN)
'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
	 	hitrust_record_all_sn = arr_result(0)
	else 
		'無資料
	end if 
else
	'錯誤訊息
end if
str_tride_order = Right("0000000" & hitrust_record_all_sn, 6) '(交易订单号)

set var_arr = nothing
set arr_result = nothing
'----------------- 新增hitrust_record_all Begin ----------------------- end ------------------
str_product_name = replace(str_product_name," ","")

'功能函数。获取10位的日期
Function getDateStr() 
    dim dateStr1,dateStr2,strTemp 
    dateStr1 = split(cstr(formatdatetime(date(),2)),"/") 

    for each StrTemp in dateStr1 
        if len(StrTemp)<2 then 
	        getDateStr=getDateStr & "0" & strTemp 
        else 
	        getDateStr=getDateStr & strTemp 
        end if 
    next 
End function 
'功能函数。获取14位的日期。结束

dim objLogOpt : objLogOpt = null
dim strFileName : strFileName = "" 'log檔名 
dim arrColumn : arrColumn = "" '資料欄位名稱
dim arrData : arrData = "" 'log資料
dim intCount : intCount = 0 '欄位筆數
dim strData : strData = "" 'log資料串
dim strUserIP : strUserIP = ""

Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")

strFileName = "log_SendCard_2"
strUserIP = getIpAddress()
arrColumn = array("HitrustRecordAllSn", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "PayType", "InputTime", "ProductName", "ClientName", "ClientEmail", "ClientIP")
arrData = array(hitrust_record_all_sn, int_dsn, g_var_client_sn, int_contract_sn, str_client_pay_cash, strPlamName, now(), str_product_name, str_client_cname, str_client_email, strUserIP)
for intCount = 0 to UBound(arrData)
    if ( intCount > 0 ) then
        strData = strData & ","
    end If
    strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
next

strData = strData & vbCrLF
Call objLogOpt.WriteLog(strFileName, strData)
Set objLogOpt = nothing

dim strEmailToName : strEmailToName = ""
dim strEmailToAddr : strEmailToAddr = ""
dim strEmailFormAddr : strEmailFormAddr = ""
dim strEmailbody : strEmailbody = ""
dim strEmailSubject : strEmailSubject = ""
dim arrEmailFromParas : arrEmailFromParas = null
dim arrEmailToParas : arrEmailToParas = null
dim arrSetParas : arrSetParas = null
dim arrOptParas : arrOptParas = null

strEmailbody = "客户 ( " & str_client_cname & " ):<br />"
strEmailbody = strEmailbody & "于 " & now() & " 使用 <strong>招商银行</strong> 刷卡页面进行刷卡！<br />"
strEmailbody = strEmailbody & "(产品: " & str_product_name & " 金额: " & str_client_pay_cash & ")"
strEmailSubject = "客户刷卡提示系统信"
strEmailToAddr = CONST_EMAIL_CREDIT_CARD_PAYMENT
strEmailFormAddr = CONST_MAIL_ADDRESS_SERVICES_VIPABC

arrEmailFromParas = Array(strEmailFormAddr, "VIPABC-刷卡系統", strEmailSubject, strEmailbody, "", "", "")
arrEmailToParas = Array(strEmailToAddr, "系統判定群組", "", "")
arrSetParas = Array(CONST_MAIL_SERVER_CLIENT, "utf-8", "", "", "")
arrOptParas = Array(true, false, false, "", "", "", "")

'g_int_err_code = sendEmail(arrEmailFromParas, arrEmailToParas, arrSetParas, arrOptParas, CONST_TUTORABC_WEBSITE)
'if (g_int_err_code = CONST_FUNC_EXE_SUCCESS) then
'else
'end if
%>
    <table border="0" cellspacing="5" cellpadding="5" width="100%">
        <tr>
            <td>
                <img id="img_ajax_loader" src="/lib/javascript/JQuery/images/ajax-loader.gif" border="0" />
                <a style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 15px; text-decoration: none;">刷卡处理中, 请勿关闭此页面…<br />(易造成刷卡失败)</a>
            </td>
        </tr>
    </table>
    <form name="kqPay" method="POST" action="<%=strBankPostUrl%>" />
    <input type="hidden" name="BranchID" value="<%=CONST_CREDIT_CARD_CMBCHINA_BANK_KEY%>">
    <input type="hidden" name="CoNo" value="<%=CONST_CREDIT_CARD_CMBCHINA_KEY%>">
    <input type="hidden" name="MerchantPara" value="<%=int_dsn%>">
    <input type="hidden" name="BillNo" value="<%=str_tride_order%>">
    <input type="hidden" name="Amount" value="<%=str_client_pay_cash%>">
    <input type="hidden" name="Date" value="<%=getDateStr()%>">
    <input type="hidden" name="MerchantUrl" value="http://<%=CONST_VIPABC_WEBHOST_NAME%>/program/member/payment_return_page_for_cmbchina.asp">
    </form>
    <script type="text/javascript" language="javascript">
        document.kqPay.submit();
    </script>
</body>
</html>
