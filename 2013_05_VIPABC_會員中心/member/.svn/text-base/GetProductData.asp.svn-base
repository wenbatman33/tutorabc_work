<!--#include virtual="/lib/include/global.inc" -->
<%
    Dim intType : intType = getRequest("Type",CONST_DECODE_NO)  '9:取得VIPABC ACH銀行的下拉選單
    Dim strProductSn : strProductSn = ""                        'cfg_product.sn
    Dim intResult : intResult = 0
    Dim strResult : strResult = ""
    Dim arrParam
    Dim strSql : strSql = ""
    Dim strPayMode : strPayMode = ""
    Dim strSelectName : strSelectName = ""

    Set objProduct = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
    if("9" = intType) then  '取得VIPABC ACH銀行的下拉選單
        strPayMode = getRequest("PayMode",CONST_DECODE_NO)
        strSelectName = getRequest("SelectID",CONST_DECODE_NO)
        if(not isEmptyOrNull(strPayMode)) then
            strSql = getSqlNotes() & " SELECT * FROM cfg_bank_info WITH (NOLOCK) WHERE Valid = 1 "
            if("50" = strPayMode) then
                strSql = strSql & " AND BankCardType = 2 "
            else
                strSql = strSql & " AND BankCardType = 1 "
            end if
			intSqlWriteResult = objProduct.excuteSqlStatementEach(strSql, "", CONST_TUTORABC_RW_CONN)
			while(not objProduct.eof)
                strResult = strResult & "<option value='" & objProduct("id") & "'>" & objProduct("BankCName") & "</option>"
                objProduct.movenext
            wend
        end if
        response.write "<select name=" & strSelectName & " id=" & strSelectName & ">"
        response.write strResult
        response.write "</select>"
    end if

    Set objProduct = nothing
%>