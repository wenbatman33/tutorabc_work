<!--#include virtual="/lib/include/global.inc"-->
<%
dim bolDebugMode : bolDebugMode = false '除錯模式

Dim str_client_cname  : str_client_cname = unescape(request("txt_cname")) '客戶中文姓名
Dim int_product_sn : int_product_sn = session("product_sn") '購買產品的序號(cfg_product.sn)
Dim str_client_tel_code : str_client_tel_code = unescape(Request("txt_tel_code")) '手機號碼前四碼
Dim str_client_tel :str_client_tel  = unescape(Request("txt_tel")) '手機號碼後六碼
Dim str_client_email :str_client_email  = unescape(Request("txt_email")) '客戶email
Dim str_client_ip : str_client_ip = getIpAddress '客戶ip
Dim str_iscustom : str_iscustom = unescape(Request("iscustom")) '是否購買
Dim str_client_pay_cash : str_client_pay_cash = unescape(Request("pay_cash")) '價錢

Dim int_dsn : int_dsn = unescape(request("dsn")) '付款相關table client_payment_detail.sn
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
	
if ( int_product_sn = "" ) then
	int_product_sn  = unescape(request("productsn"))
end if

'檢查產品代號
if ( int_product_sn = "" ) then
    response.Write ""
    '如果沒有產品代號 不傳值
    response.end
end if
'----------------- 抓取欲購買產品名稱及期限----------------------- start -----------------
product_sql = " SELECT Cname, valid_month FROM dbo.cfg_product WHERE (sn =@product_sn) "
var_arr = Array(int_product_sn)
arr_result = excuteSqlStatementRead(product_sql,var_arr,CONST_VIPABC_RW_CONN)

'response.Write g_str_sql_statement_for_debug
if ( isSelectQuerySuccess(arr_result) ) then
	if ( Ubound(arr_result) >= 0 ) then
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
'檢查payment_detail_sn 與 g_var_client_sn是否存在
if ( int_dsn = "" or g_var_client_sn = "" ) then
    response.Write ""
    '如果沒有產品代號 不傳值
    response.end
end if

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

if ( isSelectQuerySuccess(arr_result) ) then
	if ( Ubound(arr_result) >= 0 ) then
		'response.Write("欄:" & Ubound(arr_result) & "<br>")
		'response.Write("列:" & Ubound(arr_result,2) & "<br>")	
		int_contract_sn = arr_result(0,0)
		int_lead_sn = arr_result(1,0)
	else
		str_product_name = "VIPABC"
		int_valid_month = 1
	end if
end if
set contract_sql = nothing
set var_arr = nothing
set arr_result = nothing
'----------------- 新增online_payment_data ----------------------- start ------------------

'if ( g_var_client_sn = "223543" ) then
	intContractSn = int_contract_sn
%>
<!--#include virtual="/program/member/choosePaymentAccount.asp" -->
<% 
'end if 

Dim int_comp_type : int_comp_type = CONST_BANK_PLATFORM_PARAMETER_ALIPAY '1=網際威信 2=華南銀行 3=陸網銀聯 4=支付寶

if ( true = bolBJChannel ) then
    int_comp_type = CONST_BANK_PLATFORM_PARAMETER_BJ_ALIPAY
elseif ( true = bolSHChannel ) then
    int_comp_type = CONST_BANK_PLATFORM_PARAMETER_ALIPAY
end if

partner = CONST_CREDIT_CARD_ALIPAY_CODE '支付宝的账户的合作者身份ID
key = CONST_CREDIT_CARD_ALIPAY_KEY '支付宝的安全校验码
seller_email = CONST_CREDIT_CARD_ALIPAY_SELLER_EMAIL '请设置成您自己的支付宝帐户 PS.還沒拿到
show_url = CONST_CREDIT_CARD_ALIPAY_SHOW_URL '网站的网址
%>
<!--#include virtual="/program/member/alipayto/Alipay_Payto.asp"-->
<%

str_tablecol = "account_sn,cname,tel_code,tel,email,cip,cmoney,client_sn"
str_tablerow = "@account_sn,@cname,@tel_code,@tel,@email,@IP,@cmoney,@client_sn"
var_arr = Array(int_contract_sn, str_client_cname, str_client_tel_code, str_client_tel, str_client_email ,getIpAddress, str_client_pay_cash, g_var_client_sn)

arr_result = excuteSqlStatementWrite(" INSERT INTO online_payment_data (" & str_tablecol & ") VALUES (" & str_tablerow & ") ", var_arr, CONST_VIPABC_RW_CONN)
'response.Write "sql="&g_str_sql_statement_for_debug
'response.Write "......................"& payment_data_effected_row &"......................"
if ( arr_result >= 0) then 
	'response.Write("更新影響筆數 " & invoice_effected_row & "<br>")
else
	response.Write(getWord("CONTRACT_AGREE_2")& g_str_run_time_err_msg & "<br>")
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

str_tablecol = "Account_sn,Ordernumber,comp_type,Orderdesc,Amount"
str_tablecol = str_tablecol & ",Currency,Depositflag,Queryflag,payment_detail_sn" 'card_serve_bank
str_tablerow ="@Account_sn,@Ordernumber,@comp_type,@Orderdesc,@Amount"
str_tablerow = str_tablerow & ",@Currency,@Depositflag,@Queryflag,@payment_detail_sn" ',@card_serve_bank
var_arr = Array(int_contract_sn,int_contract_sn,int_comp_type,str_ostr,str_client_pay_cash,"RMB","1","0",int_dsn) ',"236"
arr_result = excuteSqlStatementScalar(" INSERT INTO hitrust_record_all (" & str_tablecol & ") VALUES (" & str_tablerow & ") SELECT SCOPE_IDENTITY()", var_arr, CONST_VIPABC_RW_CONN)
'response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
'response.Write("錯誤訊息sql....:" & g_str_run_time_err_msg & "<br>")
'response.Write("錯誤訊息sql....:" & g_int_run_time_err_code & "<br>")
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
    
if ( hitrust_record_all_sn <> "" ) then         
    '如果有成功寫入傳回網址導入支付寶
    Dim date_now_pay_time '刷卡的時間
    Dim str_pay_time_formate '傳送給支付寶的時間格式
    Dim subject, body, out_trade_no, price, quantity, discount
    Dim AlipayObj
    'date_now_pay_time = now()
    'str_pay_time_formate = year(date_now_pay_time) & month(date_now_pay_time) & day(date_now_pay_time) & hour(date_now_pay_time) & minute(date_now_pay_time) & second(date_now_pay_time)
     sTime = now()
     str_pay_time_formate = year(sTime) & right("0" & month(sTime),2) & right("0" & day(sTime),2) & right("0" & hour(sTime),2) & right("0" & minute(sTime),2) & right("0" & second(sTime),2)
    '客户网站订单号，（现取系统时间，可改成网站自己的变量）
    'table名稱 hitrust_record_all
    subject = hitrust_record_all_sn '商品名称 orderdesc 
    body = int_product_sn 'body商品描述
    out_trade_no = str_pay_time_formate       
    price = str_client_pay_cash 'price商品单价 0.01～50000.00
    quantity = "1" '商品数量,如果走购物车默认为1
    discount = "0" '商品折扣
    seller_email = seller_email '卖家的支付宝帐号
    Set AlipayObj = New creatAlipayItemURL
    Dim itemURL
    itemUrl = AlipayObj.creatAlipayItemURL(subject, body, out_trade_no, price, quantity, seller_email)	
    '要導向的網址
    response.Write itemUrl
else
    '錯誤回傳空
    response.Write ""
end if
%>
