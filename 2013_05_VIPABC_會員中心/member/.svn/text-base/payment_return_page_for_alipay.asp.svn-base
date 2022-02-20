<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<html>
<head>
</head>
<meta http-equiv="Content-Language" content="utf-8">
<body>
<%
Dim str_status : str_status = ""
Dim str_hitrust_sn : str_hitrust_sn = "" 'hitrust_record_all.sn
Dim str_dsn_sn : str_dsn_sn = "" 'dsn
Dim str_contract_sn : str_contract_sn = "" 		'合約編號
Dim str_return_value : str_return_value = ""	'寫入hitrust_record_all中
Dim str_client_sn :  str_client_sn = "" 'client_sn

'Dim str_message : str_message = Request("Message") '應答碼的中文含義
'塞資料的相關設定值
Dim str_sql : str_sql = ""
Dim var_arr 		'傳excuteSqlStatementRead的陣列
Dim arr_result		'接回來的陣列值
Dim str_tablecol 	'table欄位
Dim str_tablevalue 	'table值

'取出hitrust_sn
if (Request("subject") <> "") then
    str_hitrust_sn = trim(Request("subject")) 'hitrust_record_all.sn
    str_status = CONST_CREDIT_CARD_ALIPAY_PAY_SUCCEED_CODE
else
    str_status = ""
end if

'----------------------- 捉回合約相關資訊 ---------------------- start -----------------
str_sql = "select top 1 payment_detail_sn,Account_sn from hitrust_record_all where sn = @sn"
var_arr = Array(str_hitrust_sn)
arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then
	'如果有查尋hitrust_record_all資料
	if (Ubound(arr_result) >= 0) then 
	    '回傳detail_sn
		str_dsn_sn = arr_result(0,0)
		str_contract_sn = arr_result(1,0)
	else
		'錯誤訊息
		'response.write "ccc........." &  arr_result(0,0)& "<br>"
	end if 
end if 
'----------------------- 捉回合約相關資訊 ---------------------- end -----------------

'--------------- 付款成功 -------- start ----------
if ( str_status <> "") then
%>
    <form name="frm_payment_return" id="frm_payment_return" action="buy_access.asp" method="post">
        <input name="hdn_contract_sn" type="hidden" value="<%=str_contract_sn%>" />
        <input name="hdn_dsn_sn" type="hidden" value="<%=str_dsn_sn%>" />
    <script type="text/javascript" language="javascript">
        document.frm_payment_return.submit();
    </script>
    </form>    
<% else	%>
    <form name="frm_payment_return" id="Form1" action="buy_fail.asp" method="post">
        <input name="hdn_contract_sn" type="hidden" value="<%=str_contract_sn%>" />
        <input name="hdn_dsn_sn" type="hidden" value="<%=str_dsn_sn%>" />
    <script type="text/javascript" language="javascript">
        document.frm_payment_return.submit();
    </script>
    </form>       
<% end if %>
</body>
</html>
