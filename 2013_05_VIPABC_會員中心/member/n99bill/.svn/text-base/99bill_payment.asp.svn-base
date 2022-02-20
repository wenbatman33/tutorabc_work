<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>快钱付款网关</title>
</head>
<body>
<%
dim bolDebugMode : bolDebugMode = false '除錯模式
'dim strTransBeijingDomain : strTransBeijingDomain = CONST_TRANS_BEIJING_DOMAIN
dim strBankPostUrl : strBankPostUrl = ""

'if ( "1" = strTransBeijingDomain ) then
    'strBankPostUrl = "http://www.vipabc.com/program/member/n99bill/99bill_payment.asp"
'else
    strBankPostUrl = "https://www.99bill.com/gateway/recvMerchantInfoAction.htm"
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
if ( int_product_sn = "" or int_dsn = "" or g_var_client_sn="" ) then
	'如果session 掉了 導回首頁
	alertGo "网页逾时，请重新登入","http://"&CONST_VIPABC_WEBHOST_NAME&"/index.asp",CONST_ALERT_THEN_REDIRECT
end if

'----------------- 抓取欲購買產品名稱及期限----------------------- start -----------------
product_sql="SELECT Cname, valid_month FROM dbo.cfg_product WHERE (sn =@product_sn)"
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
contract_sql="SELECT client_payment_record.contract_sn,client_temporal_contract.lead_sn "
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

'if ( g_var_client_sn = "223543" ) then
	intContractSn = int_contract_sn
%>
<!--#include virtual="/program/member/choosePaymentAccount.asp" -->
<%
'end if

'----------------- 更新client_basic居住地址----------------------- Begin ------------------
Dim int_comp_type : int_comp_type = CONST_BANK_PLATFORM_PARAMETER_99BILL '1=網際威信 2=華南銀行 3=陸網銀聯 6.快钱 9.快錢北京

if ( true = bolBJChannel ) then
    int_comp_type = CONST_BANK_PLATFORM_PARAMETER_BJ_99BILL
elseif ( true = bolSHChannel ) then
    int_comp_type = CONST_BANK_PLATFORM_PARAMETER_99BILL
end if

dim strPlamName : strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_99BILL_NAME

if ( true = bolBJChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_BJ_99BILL_NAME
elseif ( true = bolSHChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_SH_99BILL_NAME
end if

var_arr = Array(str_client_address, g_var_client_sn)
arr_result = excuteSqlStatementWrite( " UPDATE client_basic SET client_address = @client_address WHERE sn = @g_int_client_sn ", var_arr, CONST_VIPABC_RW_CONN)

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
arr_result = excuteSqlStatementWrite(" Update lead_basic SET client_address = @client_address where sn = @g_int_lead_sn ", var_arr, CONST_VIPABC_RW_CONN)

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
if ( arr_result >= 0) then 
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

str_tablecol = "Account_sn,Storeid,Ordernumber,comp_type,Orderdesc,Amount"
str_tablecol = str_tablecol & ",Currency,Depositflag,Queryflag,payment_detail_sn" 'card_serve_bank
str_tablerow ="@Account_sn,@Storeid,@Ordernumber,@comp_type,@Orderdesc,@Amount"
str_tablerow = str_tablerow & ",@Currency,@Depositflag,@Queryflag,@payment_detail_sn" ',@card_serve_bank
var_arr = Array(int_contract_sn,str_MerId,int_contract_sn,int_comp_type,str_ostr,str_client_pay_cash,"RMB","1","0",int_dsn) ',"236"

arr_result = excuteSqlStatementScalar("INSERT INTO hitrust_record_all ("&str_tablecol&") VALUES ("&str_tablerow&") SELECT SCOPE_IDENTITY()", var_arr,CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
	 	hitrust_record_all_sn = arr_result(0)
	else 
		'無資料
	end if 
else
	'錯誤訊息
end if

str_tride_order = Right("0000000" & hitrust_record_all_sn,7) '(交易订单号)

set var_arr = nothing
set arr_result = nothing
'----------------- 新增hitrust_record_all Begin ----------------------- end ------------------
str_product_name = replace(str_product_name," ","")

'*
'* @Description: 快钱人民币支付网关接口范例
'* @Copyright (c) 上海快钱信息服务有限公司
'* @version 2.0
'*

'人民币网关账户号
''请登录快钱系统获取用户编号，用户编号后加01即为人民币网关账户号。
merchantAcctId = CONST_CREDIT_CARD_99BILL_ACCTID '測試
'merchantAcctId="1002021291101" '正式

'人民币网关密钥
''区分大小写.请与快钱联系索取
key = CONST_CREDIT_CARD_99BILL_KEY '測試
'key="J4CNSQHLM7JAYRRB" '正式

if ( g_var_client_sn = "223543" ) then
	if ( true = bolDebugMode ) then
		response.write "merchantAcctId:" & merchantAcctId & "<br />"
		response.write "key:" & key & "<br />"
		response.end
	end if
end if

'字符集.固定选择值。可为空。
''只能选择1、2、3.
''1代表UTF-8; 2代表GBK; 3代表gb2312
''默认值为1
inputCharset = "1"

'服务器接受支付结果的后台地址.与[pageUrl]不能同时为空。必须是绝对地址。
''快钱通过服务器连接的方式将交易结果发送到[bgUrl]对应的页面地址，在商户处理完成后输出的<result>如果为1，页面会转向到<redirecturl>对应的地址。
''如果快钱未接收到<redirecturl>对应的地址，快钱将把支付结果GEt到[pageUrl]对应的页面。
bgUrl = "http://" & CONST_VIPABC_WEBHOST_NAME & "/program/member/n99bill/payment_receive.asp"
	
'网关版本.固定值
''快钱会根据版本号来调用对应的接口处理程序。
''本代码版本号固定为v2.0
version = "v2.0"

'语言种类.固定选择值。
''只能选择1、2、3
''1代表中文；2代表英文
''默认值为1
language = "1"

'签名类型.固定值
''1代表MD5签名
''当前版本固定为1
signType = "1"
   
'支付人姓名
''可为中文或英文字符
payerName = str_client_cname

'支付人联系方式类型.固定选择值
''只能选择1
''1代表Email
payerContactType = "1"

'支付人联系方式
''只能选择Email或手机号
payerContact = str_client_email

'商户订单号
''由字母、数字、或[-][_]组成
orderId = str_tride_order

'订单金额
''以分为单位，必须是整型数字
''比方2，代表0.02元
orderAmount = (str_client_pay_cash*100)
	
'订单提交时间
''14位数字。年[4位]月[2位]日[2位]时[2位]分[2位]秒[2位]
''如；20080101010101
orderTime = getDateStr()

'商品名称
''可为中文或英文字符
productName = str_product_name

'商品数量
''可为空，非空时必须为数字
productNum = "1"

'商品代码
''可为字符或者数字
productId = int_product_sn

'商品描述
productDesc = str_product_name
	
'扩展字段1
''在支付结束后原样返回给商户
ext1 = ""

'扩展字段2
''在支付结束后原样返回给商户
ext2 = ""
	
'支付方式.固定选择值
''只能选择00、10、11、12、13、14
''00：组合支付（网关支付页面显示快钱支持的各种支付方式，推荐使用）10：银行卡支付（网关支付页面只显示银行卡支付）.11：电话银行支付（网关支付页面只显示电话支付）.12：快钱账户支付（网关支付页面只显示快钱账户支付）.13：线下支付（网关支付页面只显示线下支付方式）
payType = "00"

'同一订单禁止重复提交标志
''固定选择值： 1、0
''1代表同一订单号只允许提交1次；0表示同一订单号在没有支付成功的前提下可重复提交多次。默认为0建议实物购物车结算类商户采用0；虚拟产品类商户采用1
redoFlag = "0"

'快钱的合作伙伴的账户号
''如未和快钱签订代理合作协议，不需要填写本参数
pid = ""


'生成加密签名串
''请务必按照如下顺序和规则组成加密串！
	signMsgVal=appendParam(signMsgVal,"inputCharset",inputCharset)
	signMsgVal=appendParam(signMsgVal,"bgUrl",bgUrl)
	signMsgVal=appendParam(signMsgVal,"version",version)
	signMsgVal=appendParam(signMsgVal,"language",language)
	signMsgVal=appendParam(signMsgVal,"signType",signType)
	signMsgVal=appendParam(signMsgVal,"merchantAcctId",merchantAcctId)
	signMsgVal=appendParam(signMsgVal,"payerName",payerName)
	signMsgVal=appendParam(signMsgVal,"payerContactType",payerContactType)
	signMsgVal=appendParam(signMsgVal,"payerContact",payerContact)
	signMsgVal=appendParam(signMsgVal,"orderId",orderId)
	signMsgVal=appendParam(signMsgVal,"orderAmount",orderAmount)
	signMsgVal=appendParam(signMsgVal,"orderTime",orderTime)
	signMsgVal=appendParam(signMsgVal,"productName",productName)
	signMsgVal=appendParam(signMsgVal,"productNum",productNum)
	signMsgVal=appendParam(signMsgVal,"productId",productId)
	signMsgVal=appendParam(signMsgVal,"productDesc",productDesc)
	signMsgVal=appendParam(signMsgVal,"ext1",ext1)
	signMsgVal=appendParam(signMsgVal,"ext2",ext2)
	signMsgVal=appendParam(signMsgVal,"payType",payType)
	signMsgVal=appendParam(signMsgVal,"redoFlag",redoFlag)
	signMsgVal=appendParam(signMsgVal,"pid",pid)
	signMsgVal=appendParam(signMsgVal,"key",key)
signMsg= Ucase(md5(signMsgVal))

'功能函数。将变量值不为空的参数组成字符串
Function appendParam(returnStr,paramId,paramValue)
	If returnStr <> "" Then
		If paramValue <> "" then
			returnStr=returnStr&"&"&paramId&"="&paramValue
		End if
	Else 
		If paramValue <> "" then
			returnStr=paramId&"="&paramValue
		End if
	End if
		
	appendParam=ReturnStr
End Function
'功能函数。将变量值不为空的参数组成字符串。结束

'功能函数。获取14位的日期
Function getDateStr() 
    dim dateStr1,dateStr2,strTemp 
    dateStr1=split(cstr(formatdatetime(date(),2)),"/") 
    dateStr2=split(cstr(formatdatetime(time(),4) & ":" & second(now)),":") 

    for each StrTemp in dateStr1 
	    if len(StrTemp)<2 then 
	        getDateStr=getDateStr & "0" & strTemp 
	    else 
	        getDateStr=getDateStr & strTemp 
	    end if 
    next 

    for each StrTemp in dateStr2 
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
for intCount = 0 to UBound(arrColumn)
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

strEmailbody = "客户 " & str_client_cname & ":<br />"
strEmailbody = strEmailbody & "于 " & now() & " 使用 <strong>快钱支付</strong> 刷卡页面进行刷卡！<br />"
strEmailbody = strEmailbody & "(产品: " & str_product_name & ", 金额: " & str_client_pay_cash & ")"
strEmailSubject = "客户刷卡提示系统信"
strEmailToAddr = CONST_EMAIL_CREDIT_CARD_PAYMENT
strEmailFormAddr = CONST_MAIL_ADDRESS_SERVICES_VIPABC

arrEmailFromParas = Array(strEmailFormAddr, "VIPABC-刷卡系統", strEmailSubject, strEmailbody, "", "", "")
arrEmailToParas = Array(strEmailToAddr, "系統判定群組", "", "")
arrSetParas = Array(CONST_MAIL_SERVER_CLIENT, "utf-8", "", "", "")
arrOptParas = Array(true, false, false, "", "", "", "")

'g_int_err_code = sendEmail(arrEmailFromParas, arrEmailToParas, arrSetParas, arrOptParas, CONST_TUTORABC_WEBSITE)
'if ( g_int_err_code = CONST_FUNC_EXE_SUCCESS ) then
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
    <!--<form name="kqPay" method="post" action="http://sandbox.99bill.com/gateway/recvMerchantInfoAction.htm"/>-->
    <input type="hidden" name="inputCharset" value="<%=inputCharset %>">
    <input type="hidden" name="bgUrl" value="<%=bgUrl %>">
    <input type="hidden" name="version" value="<%=version %>">
    <input type="hidden" name="language" value="<%=language %>">
    <input type="hidden" name="signType" value="<%=signType %>">
    <input type="hidden" name="signMsg" value="<%=signMsg %>">
    <input type="hidden" name="merchantAcctId" value="<%=merchantAcctId %>">
    <input type="hidden" name="payerName" value="<%=payerName %>">
    <input type="hidden" name="payerContactType" value="<%=payerContactType %>">
    <input type="hidden" name="payerContact" value="<%=payerContact %>">
    <input type="hidden" name="orderId" value="<%=orderId %>">
    <input type="hidden" name="orderAmount" value="<%=orderAmount %>">
    <input type="hidden" name="orderTime" value="<%=orderTime %>">
    <input type="hidden" name="productName" value="<%=productName %>">
    <input type="hidden" name="productNum" value="<%=productNum %>">
    <input type="hidden" name="productId" value="<%=productId %>">
    <input type="hidden" name="productDesc" value="<%=productDesc %>">
    <input type="hidden" name="ext1" value="<%=ext1 %>">
    <input type="hidden" name="ext2" value="<%=ext2 %>">
    <input type="hidden" name="payType" value="<%=payType %>">
    <input type="hidden" name="redoFlag" value="<%=redoFlag %>">
    <input type="hidden" name="pid" value="<%=pid %>">
    </form>
    <script type="text/javascript" language="javascript">
        document.kqPay.submit();
    </script>
</body>
</html>
