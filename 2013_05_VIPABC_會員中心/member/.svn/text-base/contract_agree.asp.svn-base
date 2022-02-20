<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--include virtual="/lib/include/html/template/include/conn_new_not_login.inc" -->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->
<!--#include virtual="/lib/functions/functions_contract.asp" -->
<!--#include virtual="/lib/functions/functions_tool.asp"-->
<!--#include virtual="/lib/functions/function_dialer.asp"-->
<!--#include virtual="/lib/functions/FunctionsLeadsAssignmentRule.asp"--> <!--寫關懷的function-->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"--><!--TREE-->
<%
'若已點過電子合約則導回首頁
'Dim str_BrushValue : str_BrushValue = getClientBrushCardSn(Session("contract_sn"),"contract")

dim bolDebugMode : bolDebugMode = false '除錯模式, 正式上線改為false

if ( isEmptyOrNull(Session("contract_sn")) ) then
	 'Call alertGo(getWord("contract_agree_3"), "/index.asp", CONST_ALERT_THEN_REDIRECT)
     'Response.End
end if

Function BytesToBstr(body,Cset)
    dim objstream
    set objstream = Server.CreateObject("adodb.stream")
    objstream.Type = 1
    objstream.Mode = 3
    objstream.Open
	objstream.Write body
    objstream.Position = 0
    objstream.Type = 2
    objstream.Charset = Cset
    BytesToBstr = objstream.ReadText
    objstream.Close
    set objstream = nothing
End Function

Dim int_client_sn : int_client_sn = ""								'客戶編號
Dim int_contract_sn : int_contract_sn = request("contract_sn")		'合約編號
if ( isEmptyOrNull(int_client_sn) ) then
   int_client_sn = g_var_client_sn
end if
if ( isEmptyOrNull(int_contract_sn) ) then
   int_contract_sn = session("contract_sn")
end if

'判斷是否參數有值，沒有導回首頁
Dim str_tmp_go_page : str_tmp_go_page = ""

if ( isEmptyOrNull(int_client_sn) OR isEmptyOrNull(int_contract_sn) ) then
	str_tmp_go_page =  "member_login.asp"
	'call sub alertGo
	Call alertGo( getWord("contract_agree_1") , str_tmp_go_page, CONST_NOT_ALERT_THEN_REDIRECT)
	Response.End
end if

Dim str_url : str_url = ""
Dim str_mail_body : str_mail_body = ""
str_url = "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/member/ajax_contract.asp?contract_sn=" & request("contract_sn")&"&checkbox1="& request("checkbox1")&"&checkbox2="& request("checkbox2")&"&checkbox3="& request("checkbox3")&"&checkbox4="& request("checkbox4")&"&checkbox5="& request("checkbox5")&"&checkbox6="& request("checkbox6")&"&checkbox7="& request("checkbox7")&"&checkbox8="& request("checkbox8")&"&contract_sn_agree=y&client_sn="&request("client_sn")

dim xmlhttp : set xmlhttp = server.createobject("winhttp.winhttprequest.5.1")
xmlhttp.open "GET", str_url, false
xmlhttp.send()

str_mail_body = BytesToBstr(xmlhttp.responseBody,"utf-8")

'======================================
'回寫client_temporal_contract Begin
'response.write "str_url...."&str_url&"<br>"
'response.write str_mail_body
'response.end
Dim client_temporal_contract_effected_row
	var_arr = Array("1",now,str_mail_body,int_contract_sn)
	client_temporal_contract_effected_row = excuteSqlStatementWrite("Update client_temporal_contract SET cagree = @g_int_cagree,cadate=@g_date_cadate,cdata=@g_str_cdata where sn = @g_int_contract_sn", var_arr, CONST_VIPABC_RW_CONN)
'Response.Write("err1=" & g_str_sql_statement_for_debug & "<br>")
if (client_temporal_contract_effected_row >= 0) then
	'Response.Write("更新影響筆數 " & client_temporal_contract_effected_row & "<br>")
else
	Response.Write(getWord("contract_agree_2") & g_str_run_time_err_msg & "<br>")
end if
set client_temporal_contract_effected_row = nothing
set var_arr = nothing
'回寫client_temporal_contract End
'=========================================
'找出合約類型 0:一般 1:短期 2:學習儲值卡
Dim strSql : strSql = ""
Dim arrParameter
Dim intContractType : intContractType = 0
'提供ACH使用的參數
dim contract_pmethod : contract_pmethod = ""

'20121214 阿捨新增 找尋產品編號
'response.write "int_contract_sn = " & int_contract_sn & "<br />"
dim intProductSn : intProductSn = ""
dim intPaytype : intPaytype = 1
Set objContractAgree = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
if(not isEmptyOrNull(int_contract_sn)) then
    arrParameter = Array(int_contract_sn)
    strSql = " SELECT isnull(contract_type,0) AS contract_type, client_temporal_contract.prodbuy, isnull(client_temporal_contract.ptype,1) AS ptype, client_temporal_contract.pmethod "
    strSql = strSql & " FROM client_temporal_contract WITH (NOLOCK) "
    strSql = strSql & " INNER JOIN cfg_product WITH (NOLOCK) ON client_temporal_contract.prodbuy = cfg_product.sn "
    strSql = strSql & " WHERE client_temporal_contract.sn = @ContractSn "
    intResult = objContractAgree.excuteSqlStatementEach(strSql , arrParameter , CONST_VIPABC_RW_CONN)
    if ( not objContractAgree.eof ) then
        intContractType = objContractAgree("contract_type")
        intProductSn = objContractAgree("prodbuy")
		intPaytype = objContractAgree("ptype")
		contract_pmethod = objContractAgree("pmethod")
    end if
    objContractAgree.close

    'Penny Learning Card 自動開通
    if ( 2 = sCInt(intContractType) ) then
        Call getVIPABCPurchaseContractNew(int_contract_sn,"275","1", 0, 0)
    end if
end if
Set objContractAgree = Nothing

'=====================================================
'判斷自動開通條件 Justin.Lin
dim arrParam : arrParam = null
dim arrSubParam : arrSubParam = null
dim strSqlAuto : strSqlAuto = "" 'sql
dim strSubSql : strSubSql = "" 'sql
dim objAutoOpenData : objAutoOpenData = null
dim arrSqlResult : arrSqlResult = null
dim intStatus : intStatus = 9 '預設9為須自動開通但合約未確認,1為合約已開通，自動撥入堂數
dim intDetailSn : intDetailSn = 0
arrParam = Array(int_contract_sn)
strSqlAuto = " SELECT contract_sn, status, ClientPaymentDetailSn FROM ContractStatus WHERE contract_sn = @intContractSn AND auto_open = 0 "
Set objAutoOpenData = excuteSqlStatementReadEach(strSqlAuto, arrParam, CONST_VIPABC_RW_CONN)
if ( true = bolDebugMode ) then
    Response.Write("錯誤訊息Sql for 讀取需自動開通的資料表:" & g_str_sql_statement_for_debug & "<br>")
end if

while ( Not objAutoOpenData.eof ) '(刷卡半自動開通)
    intStatus = objAutoOpenData("status")
    intDetailSn = objAutoOpenData("ClientPaymentDetailSn")
    arrSubParam= Array(int_contract_sn)
    strSubSql = " UPDATE ContractStatus SET auto_open = 1, adate = GETDATE() WHERE contract_sn = @intContractSn "
    arrSqlResult = excuteSqlStatementWrite(strSubSql, arrSubParam, CONST_VIPABC_RW_CONN)
    if ( true = bolDebugMode ) then
        Response.Write("錯誤訊息Sql for 關閉自動開通flag:" & g_str_sql_statement_for_debug & "<br>")
    end if
    if ( arrSqlResult >= 0 ) then
    else
    end if
    if("1" = scstr(intStatus)) then
        'ACH: 合約已開通，撥入堂數
        Call AddSession( intDetailSn, int_client_sn)
    else
        Call getVIPABCPurchaseContractNew(int_contract_sn,"275",intPaytype, 0, 0)
        '20121214 阿捨新增 若為新產品 開通後 需要設定新錄影檔扣點規則
        if ( true = isNewProduct(intProductSn, 1, 0, CONST_VIPABC_RW_CONN) ) then
            Call setNewProductRecordRule(int_contract_sn, int_client_sn, 1, "275", 0, CONST_VIPABC_RW_CONN)
        end if
    end if
    objAutoOpenData.movenext
wend

'20121130 阿捨新增 收款人及日期均存在 且client_purchase尚未建立 (全自動開通)
dim bolAuditContract : bolAuditContract = isAuditContract(int_contract_sn, 0, CONST_VIPABC_RW_CONN)
if ( true = bolAuditContract ) then
    'todo 上線前 TestCase要變0
    Call getVIPABCPurchaseContractNew(int_contract_sn, "275", intPaytype, 0, 0)

    '20121214 阿捨新增 若為新產品 開通後 需要設定新錄影檔扣點規則
    if ( true = isNewProduct(intProductSn, 1, 0, CONST_VIPABC_RW_CONN) ) then
        Call setNewProductRecordRule(int_contract_sn, int_client_sn, 1, "275", 0, CONST_VIPABC_RW_CONN)
    end if

    '20130219 阿捨新增 若為套餐產品 開通後 需判斷是否有設定錄影檔上限
    if ( true = isComboProduct(intProductSn, 0, CONST_VIPABC_RW_CONN) ) then
        Call setNewProductRecordRule(int_contract_sn, int_client_sn, 2, "275", 0, CONST_VIPABC_RW_CONN)
    end if
	
	'20130326 阿捨新增手動開通轉移VIPABCJR客戶
	if ( true = isVIPABCJRClient(int_client_sn, CONST_VIPABC_RW_CONN) ) then
		'todo 上正式機 要改回0
		Call setVIPABCJRProfile(int_client_sn, 0, CONST_VIPABC_RW_CONN)
		'Call setVIPABCJRProfile(int_client_sn, 1, CONST_VIPABC_RW_CONN)
		Call setVIPABCJRContract(int_client_sn, 0, CONST_VIPABC_RW_CONN)
		'Call setVIPABCJRContract(int_client_sn, 1, CONST_VIPABC_RW_CONN)
	end if
end if
%>
<%
'20130412 Fox ACH導頁至刷入卡號，client_temporal_contract欄位ptype為3則是ACH，要合約已確認才會成立***********************************
session("ach_contract_sn") = ""
if (intPaytype = "3" and (contract_pmethod = "48" or contract_pmethod = "50")) then
    session("ach_contract_sn") = int_contract_sn
	'判斷是否填寫過卡號資訊
    response.redirect ("https://" & CONST_VIPABC_WEBHOST_NAME & "/program/member/CardExpireDateCheck.asp?ContractSn=" & int_contract_sn )    '要改成https		
end if
'**************************************************************************************************************
%>
<script language="javascript">
alert("合同确认完成");
window.location.href = "http://www.vipabc.com";
</script>