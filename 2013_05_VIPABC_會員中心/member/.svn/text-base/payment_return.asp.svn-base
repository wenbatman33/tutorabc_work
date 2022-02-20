<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<html>
<head>
</head>
<meta http-equiv="Content-Language" content="utf-8">
<body>
<%

'str_hitrust_sn
'Dim str_Responesecode : str_Responesecode = Request("Responesecode") '應答碼，成功時為0
'if ( str_Responesecode = "0") then 
'end if 

Dim str_merid : str_merid = Request("merid") '商店代碼
Dim str_orderno : str_orderno = Request("orderno")	'訂單編號
Dim str_transdate : str_transdate = Request("transdate")
Dim str_amount : str_amount = Request("amount") 
Dim str_currencycode : str_currencycode = Request("currencycode")
Dim str_transtype : str_transtype = Request("transtype")
Dim str_status : str_status = Request("status")
Dim str_checkvalue : str_checkvalue = Request("checkvalue")
Dim str_GateId : str_GateId = Request("GateId")
Dim str_Priv1 : str_Priv1= Request("Priv1")
Dim str_hitrust_sn : str_hitrust_sn = "" 'hitrust_record_all.sn
Dim str_dsn_sn : str_dsn_sn = "" 'dsn
Dim str_score : str_score = Request("score")	'資料來源
Dim str_contract_sn : str_contract_sn = "" 		'合約編號
Dim str_return_value : str_return_value = ""	'寫入hitrust_record_all中
Dim str_org_value : str_org_value = "" '儲存原始回傳刷卡資料
Dim str_client_sn :  str_client_sn = "" 'client_sn

'Dim str_message : str_message = Request("Message") '應答碼的中文含義
'塞資料的相關設定值
Dim str_sql : str_sql = ""
Dim str_addnew_sql : str_addnew_sql  = ""
Dim var_arr 		'傳excuteSqlStatementRead的陣列
Dim arr_result		'接回來的陣列值
Dim str_tablecol 	'table欄位
Dim str_tablevalue 	'table值

'mail的部份
Dim str_email_flag
Dim str_change_flag	'是否有開通成功

'str_hitrust_sn = CLng( Mid(str_orderno,10, Len(str_orderno)-9) ) 'hitrust_record_all.sn
'str_hitrust_sn = CInt( Mid(str_orderno,10, Len(str_orderno)-9) ) 'hitrust_record_all.sn
str_hitrust_sn = CLng( Right(str_orderno,5)) 'hitrust_record_all.sn
if ( IsNumeric(str_amount) ) then
	str_amount  = str_amount / CONST_CHINA_TO_TAIWAN_EXCHANGE_RATE
end if 

str_return_value = str_status & "@_" & str_merid & "@_" & str_orderno & "@_" &str_transdate & "@_" & str_amount & "@_" & str_currencycode & "@_" & str_transtype & "@_" & str_GateId & "@_" & str_Priv1 

'response.end 
'str_return_value = "0@_00@"&str_return_value 
'0 交易狀態碼
'@_
'00 交易錯誤碼
'@_
'029242 特店代碼
'@_
'3 交易金額
'@_
'822913 訂單編號
'@_
'F267204B00000162660_82291311412 交易序號
'@_
'交易授權碼
'@_
'卡號未四碼
'@_
'331 授權失敗原因
'@_
'null 交易金額代號
'@_
'幣值指數

'暫時先寫死作測試用的 Begin
'if (g_var_client_sn = "") then g_var_client_sn= "149532" else  g_var_client_sn = g_var_client_sn end if    '102334
'暫時先寫死作測試用的 End

'response.write "merid........." & str_merid & "<br>"
'response.write "orderno........." & str_orderno & "<br>"
'response.write "transdate........." & str_transdate & "<br>"
'response.write "amount........." & str_amount & "<br>"
'response.write "currencycode........." & str_currencycode & "<br>"
'response.write "transtype........." & str_transtype & "<br>"
'response.write "status........." & str_status & "<br>"
'response.write "checkvalue........." & str_checkvalue & "<br>"
'response.write "GateId........." & str_GateId & "<br>"
'response.write "Priv1........." & str_Priv1 & "<br>"
'response.write "sn........." & str_hitrust_sn & "<br>"
'response.write "dsn........." & int_dsn & "<br>"

'----------------------- 捉回client_sn ---------------------- start -----------------
'如果還有session就抓session值
str_client_sn = g_var_client_sn
'如果沒有就去資料庫撈資料
if( isEmptyOrNull(str_client_sn) ) then	
    str_sql = " SELECT TOP 1 payment_detail_sn FROM hitrust_record_all WHERE sn = @sn "
    var_arr = Array(str_hitrust_sn)
    arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
    if (isSelectQuerySuccess(arr_result)) then
    	'如果有查尋hitrust_record_all資料
    	if (Ubound(arr_result) >= 0) then 
    	    '回傳detail_sn
			str_dsn_sn = arr_result(0,0)
			'根據detail_sn查出client_sn
			str_sql = " SELECT TOP 1 client_sn FROM contract_pay WHERE dsn = @dsn "	
            var_arr = Array(str_dsn_sn)
            arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
            'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
            if (isSelectQuerySuccess(arr_result)) then
	            '有資料
	            if (Ubound(arr_result) >= 0) then 
		            str_client_sn = arr_result(0,0)			
	            else
		            '錯誤訊息
		            'response.write "ccc........." &  arr_result(0,0)& "<br>"
	            end if 
            end if 			
		else
			'錯誤訊息
			'response.write "ccc........." &  arr_result(0,0)& "<br>"
		end if 
	end if 
end if
'----------------------- 捉回client_sn ---------------------- end -----------------

'----------------------- 寫刷卡記錄到 hitrust_record_all ---------------------- start -----------------
'以華南刷卡記錄為主
if ( str_status = "1001" ) then
	'刷卡日報表是找第一個為0的資料
	str_return_value = "0@" & str_return_value 
else

end if 

'str_score：是誰Call這一頁：pageRet：成功 / 另一個：失敗 (需另開一個欄位來存放)  str_org_value
arr_result = excuteSqlWriteHitrustRecord(str_hitrust_sn,str_score,str_return_value)

if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then  	'有資料
		str_contract_sn = arr_result(0) '合約編號
		str_dsn_sn = arr_result(1)		'dsn
	end if 
else
	'response.write "Error............."& str_dsn_sn
end if 
'response.write str_contract_sn &"............."& str_dsn_sn
'----------------------- 寫刷卡記錄到 hitrust_record_all ---------------------- end -----------------

'--------------- 付款成功 -------- start ----------
if ( str_status = "1001" ) then

	'---------- 似乎沒用 ------ start --------
	Dim str_user_name :  str_user_name = ""
	Dim str_user_email : str_user_email = "" 
 
	str_sql = " SELECT online_payment_data.cname, online_payment_data.email "
	str_sql = str_sql & " FROM client_purchase "
	str_sql = str_sql & " INNER JOIN online_payment_data ON client_purchase.account_sn = online_payment_data.account_sn "
	str_sql = str_sql & " WHERE (client_purchase.contract_sn = @contract_sn) "
	var_arr = Array(str_contract_sn)
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
	'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
	if (isSelectQuerySuccess(arr_result)) then
	
		if (Ubound(arr_result) > 0) then  	'有資料
			str_user_name = arr_result(0,0)
			str_user_email = arr_result(1,0)
		else
			'無資料
		end if 
	end if 

'if ( g_var_client_sn = "223543" ) then
    intContractSn = str_contract_sn
%>
<!--#include virtual="/program/member/choosePaymentAccount.asp" -->
<%
'end if

dim strPlamName : strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_CHINAPAY_NAME

if ( true = bolBJChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_BJ_CHINAPAY_NAME
elseif ( true = bolSHChannel ) then
    strPlamName = CONST_PAYMENT_PLATFORM_VIPABC_SH_CHINAPAY_NAME
end if

	'---------- 似乎沒用 ------ start --------
	str_change_flag = changeCardPaymentDate(str_contract_sn,str_dsn_sn)
	'response.write str_change_flag
	
	'--------------- 刷卡成功寄信 --------------- start ------------
	'----------------寄給業務  &  主管 & 財務部
	'str_email_flag = SendCardPaymentMail(str_client_sn,str_dsn_sn,str_status,strPlamName)
    str_email_flag = SendCardPaymentMail_test(str_client_sn, str_dsn_sn, str_hitrust_sn, str_status, strPlamName)
	'response.write str_email_flag
	'--------------- 刷卡成功寄信 --------------- end ------------
	
else	
	'--------------- 刷卡失敗寄信 --------------- start ------------
	'Call alertGo("", "/program/member/member.asp?status='"&str_status&"'", CONST_NOT_ALERT_THEN_REDIRECT )
	'Response.end
	'str_email_flag = SendCardPaymentMail(str_client_sn,str_dsn_sn,str_status,strPlamName)
    str_email_flag = SendCardPaymentMail_test(str_client_sn, str_dsn_sn, str_hitrust_sn, str_status, strPlamName)
	'response.write str_email_flag
	'--------------- 刷卡失敗寄信 --------------- end ------------	 
end if 

'response.write "str_contract_sn........." & str_contract_sn & "<br>"
'response.end
%>
    <form name="frm_payment_return" id="frm_payment_return" action="buy_access.asp" method="post">
    <input name="hdn_contract_sn" type="hidden" value="<%=str_contract_sn%>" />
    <input name="hdn_dsn_sn" type="hidden" value="<%=str_dsn_sn%>" />
    <script type="text/javascript" language="javascript">
        document.frm_payment_return.submit();
    </script>
    </form>
</body>
</html>
