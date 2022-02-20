<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
Dim intClientSn : intClientSn = request("int_client_sn")		'客戶編號
Dim intContractSn : intContractSn = request("int_contract_sn")	'合約編號
Dim exmonth : exmonth = request("exmonth")	'有效期年
Dim exyear : exyear = request("exyear")	'有效期月
Dim cvs : cvs = request("cvs")	'CVS2
if ( isEmptyOrNull(intClientSn) ) then
   intClientSn = g_var_client_sn
end if

'response.write "intClientSn = " & intClientSn & "<br />"
'response.write "intContractSn = " & intContractSn & "<br />"
'response.write "exmonth = " & exmonth & "<br />"
'response.write "exyear = " & exyear & "<br />"
'response.write "cvs = " & cvs & "<br />"
'Response.End
'判斷是否參數有值，沒有導回首頁
if ( isEmptyOrNull(intClientSn) OR isEmptyOrNull(intContractSn) ) then
	str_tmp_go_page =  "member_login.asp"
	'call sub alertGo
	Call alertGo( getWord("contract_agree_1") , str_tmp_go_page, CONST_NOT_ALERT_THEN_REDIRECT)
	Response.End
end if


'找是否有新增過
Set objVIPABC_ACH_CARD_INFO_FIRST_PART = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
strSql = "SELECT top 1 CardInfoFirstPartId FROM VIPABC_ACH_CARD_INFO_FIRST_PART WHERE ContractSn = @ContractSn AND ClientSn = @ClientSn AND Valid = 1 order by CardInfoFirstPartId desc"
intSqlWriteResult = objVIPABC_ACH_CARD_INFO_FIRST_PART.excuteSqlStatementEach(strSql,  Array(intContractSn, intClientSn) , CONST_TUTORABC_RW_CONN)
if(not objVIPABC_ACH_CARD_INFO_FIRST_PART.eof) then
    intCardInfoFirstPartId = objVIPABC_ACH_CARD_INFO_FIRST_PART("CardInfoFirstPartId")
end if
objVIPABC_ACH_CARD_INFO_FIRST_PART.close
Set objVIPABC_ACH_CARD_INFO_FIRST_PART = nothing

if(isEmptyOrNull(intCardInfoFirstPartId) or 0 = intCardInfoFirstPartId) then  '新增
    arrParameters = Array(intClientSn, intContractSn, "", "", cvs, "0")
    strSql = " INSERT INTO [VIPABC_ACH_CARD_INFO_FIRST_PART]([ClientSn],[ContractSn],[CardBankAcctName],[CardBankId],[Cwthorizationcode],[CreateDete],[Creator]) "
    strSql = strSql & " VALUES(@ClientSn,@ContractSn,@PayName,@BankId,@CVVCode,getdate(),@StaffSn) "
    intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)

    Set objVIPABC_ACH_CARD_INFO_FIRST_PART = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
    strSql = "SELECT top 1 CardInfoFirstPartId FROM VIPABC_ACH_CARD_INFO_FIRST_PART WHERE ContractSn = @ContractSn AND ClientSn = @ClientSn order by CardInfoFirstPartId desc"
    intSqlWriteResult = objVIPABC_ACH_CARD_INFO_FIRST_PART.excuteSqlStatementEach(strSql,  Array(intContractSn, intClientSn) , CONST_TUTORABC_RW_CONN)
    if(not objVIPABC_ACH_CARD_INFO_FIRST_PART.eof) then
        intCardInfoFirstPartId = objVIPABC_ACH_CARD_INFO_FIRST_PART("CardInfoFirstPartId")
    end if
    objVIPABC_ACH_CARD_INFO_FIRST_PART.close
    Set objVIPABC_ACH_CARD_INFO_FIRST_PART = nothing
    
    arrParameters = Array(intCardInfoFirstPartId, "", exmonth & exyear, "0")
    strSql = " INSERT INTO [CashFlow].[dbo].[VIPABC_ACH_CARD_INFO_SECOND_PART]([CardInfoFirstPartId],[CardNo],[CardExpireDate],[Valid],[CreateDate],[Creator]) "
    strSql = strSql & " VALUES(@CardInfoFirstPartId, @CardNo, @CardExpireDate, 1, getdate(), @StaffSn) "
    intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
	
	'response.write "新增intSqlWriteResult = " & intSqlWriteResult & "<br />"
else    'update
    arrParameters = Array("", "", cvs, "0", intClientSn, intContractSn)
    strSql = " UPDATE VIPABC_ACH_CARD_INFO_FIRST_PART "
    strSql = strSql & " SET CardBankAcctName = @PayName, CardBankId = @BankId, Cwthorizationcode = @CVVCode, Creator = @StaffSn, CreateDete = getdate() "
    strSql = strSql & " WHERE ClientSn = @ClientSn AND ContractSn = @ContractSn "
    intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)

    arrParameters = Array("", exmonth & exyear, "0", intCardInfoFirstPartId)
    strSql = " UPDATE [CashFlow].[dbo].[VIPABC_ACH_CARD_INFO_SECOND_PART] "
    strSql = strSql & " SET   CardNo = @CardNo, CardExpireDate = @CardExpireDate, Creator = @StaffSn, CreateDate=GETDATE() "
    strSql = strSql & " WHERE CardInfoFirstPartId = @CardInfoFirstPartId "
    intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
	
	'response.write "更新intSqlWriteResult = " & intSqlWriteResult & "<br />"
end if

response.write ("<script>alert('您的資訊已記錄，感謝您');window.location.href = 'http://www.vipabc.com/';</script>")
%>


