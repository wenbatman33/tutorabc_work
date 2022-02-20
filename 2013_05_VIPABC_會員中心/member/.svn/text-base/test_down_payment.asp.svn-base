<!--#include virtual="/lib/include/global.inc"-->
<html>
<head>
</head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Language" content="utf-8">
<body>
<%
'g_var_client_sn
Dim str_client_cname  : str_client_cname = request("txt_cname")                     '客戶中文姓名
Dim int_product_sn : int_product_sn =  session("product_sn")  						'購買產品的序號(cfg_product.sn)

Dim str_client_address : str_client_address = Request("txt_client_address") 		'客戶住址
Dim str_client_tel_code : str_client_tel_code = Request("txt_tel_code") 			'手機號碼前四碼
Dim str_client_tel :str_client_tel  = Request("txt_tel")							'手機號碼後六碼
Dim str_client_email :str_client_email  = Request("txt_email")						'客戶email
Dim str_client_ip : str_client_ip = getIpAddress									'客戶ip
Dim str_iscustom : str_iscustom = Request("iscustom")								'是否購買
Dim str_client_pay_cash : str_client_pay_cash = Request("pay_cash") 				'價錢

Dim int_dsn : int_dsn = request("dsn")										        '付款相關table client_payment_detail.sn
Dim str_product_name : str_product_name = "VIPABC" 									'購買產品名稱
Dim int_valid_month  : int_valid_month = 1   										'有效期限
Dim int_contract_sn : int_contract_sn = 0											'合約編號
Dim int_lead_sn       'lead_sn
Dim int_comp_type : int_comp_type = 3		  			  '1=網際威信 2=華南銀行 3=陸網銀聯
Dim str_ostr          '買的系統
'----ryanlin 新增內卡外卡的判斷--------start------20100520---
'Dim str_MerId : str_MerId = "808080112599361" 					'test：MerId为ChinaPay统一分配给商户的商户号，15位长度，必填
'Dim str_MerId : str_MerId = "808080430002799" 					'外卡
'Dim str_MerId : str_MerId = "808080430002798" 					'內卡
Dim str_MerId : str_MerId = "808080430002798"
Dim str_sub_MerId : str_sub_MerId = "000002798"						'MerId为ChinaPay统一分配给商户的商户号，15位长度，必填
Dim int_credit_card_type : int_credit_card_type = CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY '預設為cinapay 銀聯
if (request("change_bank_for_chinapay") = CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY) then
	'內卡
	str_MerId = "808080430002798"
	str_sub_MerId = "000002798"	
	int_credit_card_type = CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY
elseif (request("change_bank_for_chinapay") = CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY) then
	'外卡
	str_MerId = "808080430002799"
	str_sub_MerId = "000002799"	
	int_credit_card_type = CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY
end if
'----ryanlin 新增內卡外卡的判斷--------end------20100520---
'response.write "<br>str_client_pay_cash....."&str_client_pay_cash&"<br>"
'Session("appcode")=str_ostr '????????????

'抓取欲購買產品名稱及期限
Dim product_sql : product_sql = ""

'抓取欲購買合約之相關資訊
Dim contract_sql : contract_sql = ""

Dim hitrust_record_all_sn : hitrust_record_all_sn = "" 'hitrust_record_all.sn


'----------------- 刷卡傳hidden值到下一頁所需要的變數 --------------------- start ------------------
Dim str_CuryId : str_CuryId = "156"  										'订单交易币种，3位长度，固定为人民币156，必填
Dim str_TransDate : str_TransDate = getFormatDateTime(date(),7) 			'订单交易日期，8位长度，必填
Dim str_TransType : str_TransType = "0001"            						'交易类型，4位长度，必填 0001為消費 0002為退費
Dim str_Version : str_Version = "20070129"            						'支付接入版本号
Dim str_GateId : str_GateId = ""
Dim str_tride_order : str_tride_order = "" 									'(交易订单号)
Dim str_amount : str_amount = ""											'订单交易金额，12位长度，左补0，必填,单位为分 'request("amount")
Dim str_ChkValue : str_ChkValue = ""
	
str_amount = Right("0000000000000000" & str_client_pay_cash * CONST_CHINA_TO_TAIWAN_EXCHANGE_RATE ,12)		
'订单交易金额，12位长度，左补0，必填,单位为分 'request("amount")     
'----------------- 刷卡傳hidden值到下一頁所需要的變數 ----------------------- end ----------------

'塞資料的相關設定值
Dim var_arr 		'傳excuteSqlStatementRead的陣列
Dim arr_result		'接回來的陣列值
Dim str_tablecol 	'table欄位
Dim str_tablerow 	'table列數


'暫時先寫死作測試用的 Begin
'if (g_var_client_sn = "") then g_var_client_sn= "149532" else  g_var_client_sn = g_var_client_sn end if    '102334
'暫時先寫死作測試用的 End
	
if ( int_product_sn = "") then
	int_product_sn  = request("productsn")
end if



'----------------- 抓取欲購買產品名稱及期限----------------------- start -----------------
product_sql="SELECT Cname, valid_month FROM dbo.cfg_product WHERE (sn =@product_sn)"
var_arr = Array(int_product_sn)
arr_result = excuteSqlStatementRead(product_sql,var_arr,CONST_VIPABC_RW_CONN)

'response.write g_str_sql_statement_for_debug
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



'response.write "欲購買產品名稱及期限..."&str_product_name&"............."&int_valid_month
'----------------- 抓取欲購買產品名稱及期限----------------------- end -----------------

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
		'Response.write("欄:" & Ubound(arr_result) & "<br>")
		'Response.write("列:" & Ubound(arr_result,2) & "<br>")	
		int_contract_sn = arr_result(0,0)
		int_lead_sn = arr_result(1,0)
	else
		str_product_name="VIPabc"
		int_valid_month=1
	end if
end if
set contract_sql = nothing
set var_arr = nothing
set arr_result = nothing

'response.write "購買合約之相關資訊..."&int_contract_sn&"............."&int_lead_sn
'----------------- 抓取欲購買合約之相關資訊----------------------- Begin ------------------



'TODO：發票資訊,暫時先不寫
''======================================
''更新發票資訊 Begin
'Dim invoice_effected_row
'g_var_arr = Array(int_contract_sn)
'g_invoice_effected_row = excuteSqlStatementWrite("Update client_payment_record SET invoice = @g_int_invoice,invoice_num=@g_int_inovice_num,invoice_header=@g_str_invoice_header where contract_sn = @g_int_contract_sn", var_arr, CONST_VIPABC_RW_CONN)
'
'if (effected_row >= 0) then 
'	'Response.Write("更新影響筆數 " & invoice_effected_row & "<br>")
'else
'	Response.Write("更新錯誤訊息" & g_str_run_time_err_msg & "<br>")
'end if
'set invoice_effected_row = nothing
'set var_arr = nothing
'set arr_result = nothing
''更新發票資訊 End
''=========================================


'----------------- 更新client_basic居住地址----------------------- Begin ------------------
var_arr = Array(str_client_address,g_var_client_sn)
arr_result = excuteSqlStatementWrite("Update client_basic SET client_address = @client_address where sn = @g_int_client_sn", var_arr, CONST_VIPABC_RW_CONN)

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
arr_result = excuteSqlStatementWrite("Update lead_basic SET client_address = @client_address where sn = @g_int_lead_sn", var_arr, CONST_VIPABC_RW_CONN)

if (arr_result >= 0) then 
	'Response.Write("更新影響筆數 " & invoice_effected_row & "<br>")
else
	Response.Write(getWord("CONTRACT_AGREE_2")& g_str_run_time_err_msg & "<br>")
end if
set var_arr = nothing
set arr_result = nothing
'----------------- 更新lead_basic居住地址----------------------- end ------------------

'----------------- 新增online_payment_data ----------------------- start ------------------
'if ( str_iscustom = "1" ) then 
'	str_client_pay_cash = str_client_pay_cash
'else 
'	str_client_pay_cash = NULL
'end if 

str_tablecol = "account_sn,cname,tel_code,tel,email,cip,cmoney,client_sn"
str_tablerow = "@account_sn,@cname,@tel_code,@tel,@email,@IP,@cmoney,@client_sn"
var_arr = Array(int_contract_sn,str_client_cname,str_client_tel_code,str_client_tel,str_client_email,getIpAddress,str_client_pay_cash,g_var_client_sn)

arr_result = excuteSqlStatementWrite("Insert into online_payment_data("&str_tablecol&")values("&str_tablerow&") ", var_arr, CONST_VIPABC_RW_CONN)
'response.write "sql="&g_str_sql_statement_for_debug
'response.end
'response.write "......................"& payment_data_effected_row &"......................"
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

arr_result = excuteSqlStatementScalar("Insert into hitrust_record_all("&str_tablecol&")values("&str_tablerow&") Select SCOPE_IDENTITY()", var_arr,CONST_VIPABC_RW_CONN)
'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
'response.end
'Response.Write("錯誤訊息sql....:" & g_str_run_time_err_msg & "<br>")
'Response.Write("錯誤訊息sql....:" & g_int_run_time_err_code & "<br>")
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
	 	hitrust_record_all_sn = arr_result(0)
	else 
		'無資料
	end if 
else
	'錯誤訊息
end if

set var_arr = nothing
set arr_result = nothing
'response.end 
'----------------- 新增hitrust_record_all Begin ----------------------- end ------------------


'----------------- 刷卡傳hidden值到下一頁所需要的變數 --------------------- start ------------------

'str_tride_order = "000099361"&Right("0000000" & hitrust_record_all_sn,7) 'test：(交易订单号)
str_tride_order = str_sub_MerId&Right("0000000" & hitrust_record_all_sn,7) '(交易订单号)
str_ChkValue = getNetPaySignCode(str_MerId,str_tride_order,str_amount,str_CuryId,str_TransDate,str_TransType,"VIPABC",int_credit_card_type)
'----------------- 刷卡傳hidden值到下一頁所需要的變數 ----------------------- end ------------------


'Response.Write  "<br>"
'response.write "hitrust_record_all_sn............" & hitrust_record_all_sn & "<br>"
'Response.Write("contract_sn....."&request("contract_sn")) & "<br>"
'Response.Write("str_MerId....."&str_MerId) & "<br>"
'Response.Write("str_tride_order....."&str_tride_order) & "<br>"
'Response.Write("str_amount....."&str_amount) & "<br>"
'Response.Write("str_CuryId....."&str_CuryId) & "<br>"
'Response.Write("str_TransDate....."&str_TransDate) & "<br>"
'Response.Write("str_TransType....."&str_TransType) & "<br>"
'Response.Write("str_Version....."&str_Version) & "<br>"
'Response.Write("str_ChkValue....."&str_ChkValue) & "<br>"


'response.end

'Response.end
'CONST_VIPABC_IMSHOST_NAME

	
%>

<form name="SendToMer"  action="http://payment.chinapay.com/pay/TransGet" METHOD="POST">
<!--<form name="SendToMer"  action="http://payment-test.chinapay.com/pay/TransGet" METHOD="POST">-->
<input type="hidden" name="MerId" value="<%=str_MerId%>"/><!-- MerId为ChinaPay统一分配给商户的商户号，15位长度，必填 -->
<input type="hidden" name="OrdId" value="<%=str_tride_order%>"/><!-- 商户提交给ChinaPay的交易订单号，16位长度，必填 -->
<input type="hidden" name="TransAmt" value="<%=str_amount%>"/>     <!-- 订单交易金额，12位长度，左补0，必填,单位为分 -->
<input type="hidden" name="CuryId" value="<%=str_CuryId%>"/>                <!-- 订单交易币种，3位长度，固定为人民币156，必填 -->
<input type="hidden" name="TransDate" value="<%=str_TransDate%>"/>        <!-- 订单交易日期，8位长度，必填 -->
<input type="hidden" name="TransType" value="<%=str_TransType%>"/>            <!-- 交易类型，4位长度，必填 -->
<input type="hidden" name="Version" value="<%=str_Version%>"/>          <!-- 支付接入版本号，必填 -->
<input type="hidden" name="BgRetUrl" value="http://<%=CONST_VIPABC_WEBHOST_NAME%>/program/member/payment_return.asp?score=BgRet"/>
																<!-- 后台交易接收URL，长度不要超过80个字节，必填 -->
<!--<input type="hidden" name="PageRetUrl" value="http://<%=CONST_VIPABC_WEBHOST_NAME %>/program/member/payment_return.asp"/>-->
<input type="hidden" name="PageRetUrl" value="http://<%=CONST_VIPABC_WEBHOST_NAME%>/program/member/payment_return.asp?score=pageRet"/>
																<!-- 页面交易接收URL，长度不要超过80个字节，必填 -->
<input type="hidden" name="GateId" value="<%=str_GateId%>">                <!-- 支付网关号，可选 -->                                                             
<input type="hidden" name="Priv1" value="VIPABC">               <!-- 商户私有域，长度不要超过60个字节 -->
<input type="hidden" name="ChkValue" value="<%=str_ChkValue%>">              
<!-- 256字节长的ASCII码,为此次交易提交关键数据的数字签名，必填 -->
</form>

<script Language="JavaScript">
    document.SendToMer.submit();
</script>
</body>
</html>
