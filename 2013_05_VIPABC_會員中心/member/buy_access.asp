<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/lib/functions/functions_card_payment.asp"-->
<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->
<!--#include virtual="/lib/functions/functions_contract.asp" -->
<!--#include virtual="/lib/functions/functions_tool.asp"-->
<!--#include virtual="/lib/functions/function_dialer.asp"-->
<!--#include virtual="/lib/functions/FunctionsLeadsAssignmentRule.asp"--> <!--寫關懷的function-->
<script type="text/javascript">
    //當頁面出現時重整母視窗
    $(document).ready(function () {
        opener.location.reload();
    })
</script>
<%

dim bolDebugMode : bolDebugMode = false '除錯模式

'cookie補強session機制
if ( isEmptyOrNull(g_var_client_sn) ) then
    session("client_sn") = Request.Cookies("client_sn")
    g_var_client_sn = session("client_sn")
end if

'塞資料的相關設定值
dim str_sql : str_sql = ""
dim var_arr '傳excuteSqlStatementRead的陣列
dim arr_result '接回來的陣列值
dim str_tablecol 'table欄位
dim str_tablerow 'table列數

dim intContractSn : intContractSn = trim(request("hdn_contract_sn")) '合約編號
if ( true = bolDebugMode ) then
	response.write "intContractSn:" & intContractSn & "<br/>"
	'response.end
end if
dim intDetailSn : intDetailSn = trim( request("hdn_dsn_sn") ) 'dsn
dim strSendMail : strSendMail = trim( request("sendmail") )
dim intBolLog : intBolLog = trim( request("log") )
dim strSuccCode : strSuccCode = trim( request("succ_code") )
dim strStoreCode : strStoreCode = unescape( trim( request("store_code") ) )

dim strClientCName : strClientCName = "" '客戶名字
dim strProductName : strProductName = "" '產品名稱
dim intPayMoney : intPayMoney = "" '購買金額
dim intPorderNo : intPorderNo = "" '訂單編號
dim bol_is_payment_completed : bol_is_payment_completed = false '是否全部刷卡完款成功 預設還沒

'log紀錄
dim objLogOpt : objLogOpt = null
dim strFileName : strFileName = "" 'log檔名 
dim arrColumn : arrColumn = "" '資料欄位名稱
dim arrData : arrData = "" 'log資料
dim intCount : intCount = 0 '欄位筆數
dim strData : strData = "" 'log資料串
dim strUserIP : strUserIP = ""

dim strInputTime : strInputTime = ""
dim intClientSn : intClientSn = ""
dim intProductSn : intProductSn = ""

'==================== 客戶刷卡資料 Start ====================
dim arrSqlResult : arrSqlResult = "" '客戶刷卡資料
if ( Not isEmptyOrNull(intContractSn) ) then
    arrSqlResult = getClientCardPaymentDataNew(intContractSn, intDetailSn, CONST_VIPABC_RW_CONN)
	if ( isSelectQuerySuccess(arrSqlResult) ) then
		if ( Ubound(arrSqlResult) >= 0 ) then '有資料
			intClientSn = arrSqlResult(0,0)
            strProductName = arrSqlResult(1,0)
            intPayMoney = arrSqlResult(2,0)
            strClientCName = arrSqlResult(3,0)
		    intPorderNo = arrSqlResult(4,0)
            strInputTime = arrSqlResult(5,0)
            intProductSn = arrSqlResult(6,0)
            'intPorderNo = right("0000000" & p_intHitrustSn, 6)
		else
		end if
	else
	end if
elseif ( Not isEmptyOrNull(g_var_client_sn) ) then
    arrSqlResult = getClientCardPaymentData(g_var_client_sn, intDetailSn)
    if (isSelectQuerySuccess(arrSqlResult)) then
        if (Ubound(arrSqlResult) >= 0) then '有資料
	        strClientCName = arrSqlResult(3,0)
	        strProductName = arrSqlResult(1,0)
	        intPayMoney = arrSqlResult(2,0)
	        intPorderNo = arrSqlResult(4,0)
            'intPorderNo = right("0000000" & p_intHitrustSn, 6)
	        intContractSn = arrSqlResult(5,0)
	        intDetailSn = arrSqlResult(6,0)
            strInputTime = arrSqlResult(7,0)
            intProductSn = arrSqlResult(8,0)
        else
        end if
    else
    end if
end if
'==================== 客戶刷卡資料 End ====================

'==================== 檢查是否仍有款項還沒付款 Start ====================
str_sql = " SELECT sn FROM contract_pay "
str_sql = str_sql & " WHERE sn = @sn AND paymode = @paymode AND (isbuy IS NULL OR isbuy = @isbuy) "
str_sql = str_sql & " ORDER BY dsn "
'找客戶為線上刷卡，且還沒付款的刷卡款項
var_arr = Array(intContractSn, "線上刷卡", "0")	
arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then  	'有資料
		'如果還有沒付款的資料
		bol_is_payment_completed = false	
	else		
		'如果都付完款了
		bol_is_payment_completed = true
	end if 
else
		
end if 	
'==================== 檢查是否仍有款項還沒付款 End ====================

Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")

strFileName = "log_SendCard_7-start"
strUserIP = getIpAddress()
strData = ""
arrColumn = array("SuccCode", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "ProductName", "ProductNo", "InputTime", "ScoreSn", "ClientName", "ClientIP")
arrData = array(strSuccCode, intDetailSn, g_var_client_sn, intContractSn, intPayMoney, strProductName, intPorderNo, now(), strStoreCode, strClientCName, strUserIP)
for intCount = 0 to UBound(arrColumn)
    if ( intCount > 0 ) then
        strData = strData & ","
    end If
    strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
next

strData = strData & vbCrLF
Call objLogOpt.WriteLog(strFileName, strData)
Set objLogOpt = nothing

if ( Not isEmptyOrNull(strSendMail) ) then
    if ( "true" = strSendMail AND intBolLog = 0 ) then
	    str_status = strSuccCode
        if ( Not isEmptyOrNull(intDetailSn) AND Not isEmptyOrNull(g_var_client_sn) ) then
	        str_email_flag = SendCardPaymentMail(g_var_client_sn, intDetailSn, str_status, strStoreCode)
        else
            Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")

            strFileName = "log_SendCard_7-fail"
            strUserIP = getIpAddress()
            strData = ""
            arrColumn = array("SuccCode", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "ProductName", "ProductNo", "InputTime", "ScoreSn", "ClientName", "ClientIP")
            arrData = array(strSuccCode, intDetailSn, g_var_client_sn, intContractSn, intPayMoney, strProductName, intPorderNo, now(), strStoreCode, strClientCName, strUserIP)
            for intCount = 0 to UBound(arrColumn)
                if ( intCount > 0 ) then
                    strData = strData & ","
                end If
                strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
            next

            strData = strData & vbCrLF
            Call objLogOpt.WriteLog(strFileName, strData)
            Set objLogOpt = nothing
        end if
    end if	
end if

Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")

strFileName = "log_SendCard_7-end"
strUserIP = getIpAddress()
strData = ""
arrColumn = array("SuccCode", "PaymentDetailSn", "ClientSn", "ContractSn", "PaymentAmount", "ProductName", "ProductNo", "InputTime", "ScoreSn", "ClientName", "ClientIP")
arrData = array(strSuccCode, intDetailSn, g_var_client_sn, intContractSn, intPayMoney, strProductName, intPorderNo, now(), strStoreCode, strClientCName, strUserIP)
for intCount = 0 to UBound(arrColumn)
    if ( intCount > 0 ) then
        strData = strData & ","
    end If
    strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
next

strData = strData & vbCrLF
Call objLogOpt.WriteLog(strFileName, strData)
Set objLogOpt = nothing

'判斷自動開通條件 Justin.Lin
Dim strSqlAuto : strSqlAuto = ""
Dim arrParamAuto : arrParamAuto = null
Dim bolAutoOpenContract : bolAutoOpenContract = "" '是否自動開通
dim bolContractSnExist : bolContractSnExist = false '是否已經存在於需自動開通資料表
Dim intShouldPay : intShouldPay = 0
Set objContractAgreeAuto = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
if ( not isEmptyOrNull(intContractSn) ) then
	arrParamAuto = Array(intContractSn,intContractSn)

    strSql = "  SELECT ISNULL(ptype, 1) AS ptype , ISNULL(auditor, '') AS auditor , "
    strSql = strSql & "    ( ISNULL(dbo.ProductInstallment.MonthPay, 0) - ISNULL(Discount.DiscountMoney, 0) ) AS ShouldPay "
    strSql = strSql & " FROM   dbo.client_temporal_contract WITH ( NOLOCK ) "
    strSql = strSql & "     INNER JOIN dbo.ProductInstallment WITH ( NOLOCK ) ON dbo.client_temporal_contract.prodbuy = dbo.ProductInstallment.CfgProductSn "
    strSql = strSql & "                                                         AND dbo.client_temporal_contract.ptimes = dbo.ProductInstallment.PayPeriod "
    strSql = strSql & "                                                         AND dbo.ProductInstallment.Valid = 1 "
    strSql = strSql & "    LEFT JOIN ( SELECT  SUM(dbo.client_payment_detail.amount) AS DiscountMoney , contract_sn "
    strSql = strSql & "                FROM    client_payment_record WITH ( NOLOCK ) "
    strSql = strSql & "                        INNER JOIN client_payment_detail ON dbo.client_payment_record.sn = dbo.client_payment_detail.payrecord_sn "
    strSql = strSql & "                WHERE   client_payment_record.contract_sn = @ContractSn1 AND dbo.client_payment_detail.payment_mode IN ( 14 ) "
    strSql = strSql & "                GROUP BY contract_sn "
    strSql = strSql & "              ) AS Discount ON client_temporal_contract.sn = Discount.contract_sn "
    strSql = strSql & " WHERE  client_temporal_contract.sn = @ContractSn2 "

    intResult = objContractAgreeAuto.excuteSqlStatementEach(strSql, arrParamAuto, CONST_VIPABC_RW_CONN)
    if ( Not objContractAgreeAuto.eof ) then
        intPtype = objContractAgreeAuto("ptype")
        strAuditor = objContractAgreeAuto("auditor")
        intShouldPay = objContractAgreeAuto("ShouldPay")  'ACH一次需付多少錢
    end if
    objContractAgreeAuto.close

    strSqlAuto = "     SELECT  cagree , "
    strSqlAuto = strSqlAuto & "         contract_sn AS contract_sn , "
    strSqlAuto = strSqlAuto & "         amount_total AS amount_total , "
    strSqlAuto = strSqlAuto & "         SUM(amount_detail) AS amount_detail "
    strSqlAuto = strSqlAuto & " FROM    ( SELECT DISTINCT "
    strSqlAuto = strSqlAuto & "                     client_payment_detail.sn AS detail_sn , "
    strSqlAuto = strSqlAuto & "                     client_temporal_contract.cagree , "
    strSqlAuto = strSqlAuto & "                     client_payment_record.contract_sn , "
    strSqlAuto = strSqlAuto & "                     client_payment_record.amount AS amount_total , "
    strSqlAuto = strSqlAuto & "                     client_payment_detail.adate , "
    strSqlAuto = strSqlAuto & "                     client_payment_detail.auditor , "
    strSqlAuto = strSqlAuto & "                     client_payment_detail.paperno , "
    strSqlAuto = strSqlAuto & "                     hitrust_record_all.Retcode , "
    strSqlAuto = strSqlAuto & "                     client_payment_detail.amount AS amount_detail "
    strSqlAuto = strSqlAuto & "           FROM      client_payment_record "
    if(intPtype = "3") then
        strSqlAuto = strSqlAuto & "                 INNER JOIN ( SELECT MIN(sn) AS sn "
        strSqlAuto = strSqlAuto & "                              FROM   client_payment_record "
        strSqlAuto = strSqlAuto & "                              GROUP BY contract_sn "
        strSqlAuto = strSqlAuto & "                             ) AS client_payment_record_max ON client_payment_record.sn = client_payment_record_max.sn "
    else
        strSqlAuto = strSqlAuto & "                 INNER JOIN ( SELECT MAX(sn) AS sn "
        strSqlAuto = strSqlAuto & "                              FROM   client_payment_record "
        strSqlAuto = strSqlAuto & "                              GROUP BY contract_sn "
        strSqlAuto = strSqlAuto & "                            ) AS client_payment_record_max ON client_payment_record.sn = client_payment_record_max.sn "
    end if
    strSqlAuto = strSqlAuto & "                     INNER JOIN client_payment_detail ON client_payment_record.sn = client_payment_detail.payrecord_sn "
    strSqlAuto = strSqlAuto & "                     INNER JOIN client_temporal_contract ON client_temporal_contract.sn = client_payment_record.contract_sn "
    strSqlAuto = strSqlAuto & "                     LEFT JOIN hitrust_record_all ON client_payment_detail.sn = hitrust_record_all.payment_detail_sn "
    strSqlAuto = strSqlAuto & "           WHERE     client_payment_detail.payment_mode NOT IN ( 14 ) "
    strSqlAuto = strSqlAuto & "                     AND ( ( hitrust_record_all.Retcode LIKE '0%' ) "
    strSqlAuto = strSqlAuto & "                           OR ( client_payment_detail.auditor > '' "
    strSqlAuto = strSqlAuto & "                              ) "
    strSqlAuto = strSqlAuto & "                         ) "
    strSqlAuto = strSqlAuto & "                     AND client_payment_record.contract_sn NOT IN ( "
    strSqlAuto = strSqlAuto & "                     SELECT  tcontract_sn "
    strSqlAuto = strSqlAuto & "                     FROM    client_money_back_record "
    strSqlAuto = strSqlAuto & "                     WHERE   mb_type NOT IN ( '多付退回', '不退費', '換付款方式' ) "
    strSqlAuto = strSqlAuto & "                             AND tcontract_sn > '' ) "
    strSqlAuto = strSqlAuto & "                     AND client_temporal_contract.sn IN ( "
    strSqlAuto = strSqlAuto & "                     SELECT  sn "
    strSqlAuto = strSqlAuto & "                     FROM    dbo.client_temporal_contract "
    strSqlAuto = strSqlAuto & "                     WHERE   client_sn IN ( SELECT   sn "
    strSqlAuto = strSqlAuto & "                                            FROM     client_basic ) "
    strSqlAuto = strSqlAuto & "                             AND web_site = 'C' ) "
    strSqlAuto = strSqlAuto & "         ) AS data "
    strSqlAuto = strSqlAuto & " GROUP BY data.contract_sn , "
    strSqlAuto = strSqlAuto & "         data.amount_total , "
    strSqlAuto = strSqlAuto & "         data.cagree "
    strSqlAuto = strSqlAuto & " HAVING  data.amount_total <= SUM(data.amount_detail) "
    strSqlAuto = strSqlAuto & "         AND data.contract_sn = @contract_sn "

    intResult = objContractAgreeAuto.excuteSqlStatementEach(strSqlAuto, arrParamAuto, CONST_VIPABC_RW_CONN)
	if ( true = bolDebugMode ) then
        response.Write("錯誤訊息Sql for 讀取付款資訊:" & g_str_sql_statement_for_debug & "<br>")
    end if	
    if ( Not objContractAgreeAuto.eof ) then
	    intAmountTotal = objContractAgreeAuto("amount_total") '應付總金額
	    intAmountDetail = objContractAgreeAuto("amount_detail") '付款金額明細
	    intCagree = objContractAgreeAuto("cagree") '合約是否同意 0:不同意 1:同意		
        bolAutoOpenContract = false
    end if		
    objContractAgreeAuto.close
end if

if ( true = bolDebugMode ) then	
    response.write "strSqlAuto: " & strSqlAuto & "<br />"
    response.write "intContractSn: " & intContractSn & "<br />"
    response.write "intAmountTotal: " & intAmountTotal & "<br />"
    response.write "intAmountDetail: " & intAmountDetail & "<br />"
    response.write "intShouldPay: " & intShouldPay & "<br />"
    response.write "intCagree: " & intCagree & "<br />"
    response.write "intPtype: " & intPtype & "<br />"
    response.write "strAuditor: " & strAuditor & "<br />"
end if

if ( 1 = sCInt(intCagree) ) then '檢查合約已同意
    bolAutoOpenContract = true 'bolAutoOpenContract設為true
end if
if ( true = bolAutoOpenContract ) then 'bolAutoOpenContract=true 合約自動開通
    if(isEmptyOrNull(strAuditor) or "0" = strAuditor) then
        if("3" = intPtype) then 'ACH 自動開通
            if(scint(intAmountDetail) >= scint(intShouldPay)) then   '付款金額要大於等於頭期
                Call getVIPABCPurchaseContractNew(intContractSn,"275",intPtype, 0, 0)
            end if
        else
            '刷卡半自動開通
	        Call getVIPABCPurchaseContractNew(intContractSn,"275",intPtype, 0, 0)	'呼叫合約自動開通function (275:系統自動開通 1:一次付清)
        end if
    end if
    '20121214 阿捨新增 若為新產品 開通後 需要設定新錄影檔扣點規則
    if ( true = isNewProduct(intProductSn, 1, 0, CONST_VIPABC_RW_CONN) ) then
        Call setNewProductRecordRule(intContractSn, g_var_client_sn, 1, "275", 0, CONST_VIPABC_RW_CONN)
    end if
	
	'20130326 阿捨新增手動開通轉移VIPABCJR客戶
	if ( true = isVIPABCJRClient(g_var_client_sn, CONST_VIPABC_RW_CONN) ) then
		'todo 上正式機 要改回0
		Call setVIPABCJRProfile(g_var_client_sn, 0, CONST_VIPABC_RW_CONN)
		'Call setVIPABCJRProfile(g_var_client_sn, 1, CONST_VIPABC_RW_CONN)
		Call setVIPABCJRContract(g_var_client_sn, 0, CONST_VIPABC_RW_CONN)
		'Call setVIPABCJRContract(g_var_client_sn, 1, CONST_VIPABC_RW_CONN)
	end if
elseif ( false = bolAutoOpenContract ) then '合約不自動開通並把合約未同意的資料紀錄到新開立的table credit_contract_agree (在客戶點回合約的頁面 contract_agree.asp 撈取出來做自動開通)
    if(isEmptyOrNull(strAuditor) or "0" = strAuditor) then
	    'SELECT credit_contract_agree 查詢 credit_contract_agree 是否已有此筆合約紀錄
	    dim arrParamExist : arrParamExist = null
        dim intStatus : intStatus = 9 '預設9為須自動開通但合約未確認
    		
	    Set objContractExist = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	    arrParamExist = Array(intContractSn, intStatus)
	    strSqlExist = " SELECT contract_sn FROM ContractStatus WHERE contract_sn = @intContractSn AND status = @intStatus "
	    intResult = objContractExist.excuteSqlStatementEach(strSqlExist, arrParamExist, CONST_VIPABC_RW_CONN)
	    if ( Not objContractExist.eof ) then
		    bolContractSnExist = true
		    if ( intResult <> CONST_FUNC_EXE_SUCCESS ) then
			    'Response.write "SQL for Debug:" & objContractExist.FullSqlStatement &"<br>"
			    'Response.write "RESULT:" & objContractExist.ExcuteResult
		    end if
	    end if
	    objContractExist.close
	    Set objContractExist = nothing
	
	    'INSERT 刷卡時合約未同意但已完款 把合約資料紀錄到 credit_contract_agree
	    Dim intAutoOpen : intAutoOpen = 0 '預設0為未自動開通 (intAutoOpen:自動開通開關)
	    Dim arrParamCagreeExist : arrParamCagreeExist = null
	    dim strSqlCagreeExist : strSqlCagreeExist = ""

	    Set objCagreeExist = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	    if ( bolContractSnExist ) then 
        else '合約未同意且合約未存在於 ContractStatus
            if("3" = intPtype) then 'ACH 自動開通
                if(scint(intAmountDetail) >= scint(intShouldPay)) then   '付款金額要大於等於頭期
                    Call AddContractStatus(intContractSn, intStatus, intDetailSn)
                end if
            else
                Call AddContractStatus(intContractSn, intStatus, intDetailSn)
            end if
        end if
    end if
end if
Set objContractAgreeAuto = Nothing
%>
<div id="temp_contant">
    <!--內容start-->
    <div class="main_membox">
        <div class="buy">
            <div class="w">
                <%=getWord("BUY_access_1")%>：<span class="bold color_1"><%=getWord("BUY_access_2")%></span></div>
            <!--明細start-->
            <div>
                <div class="box">
                    <!--订单编号 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_1_7")%></div>
                        <div class="show">
                            <%=intPorderNo%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--订单编号 end-->
                    <!--产品名称 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_1_17")%></div>
                        <div class="show">
                            <%=strProductName%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--产品名称 end-->
                    <!--客户姓名 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_access_3")%></div>
                        <div class="show">
                            <%=strClientCName%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--客户姓名 end-->
                    <!--信用卡号码 start-->
                    <!--<div class="con">
                <div class="name"><%=getWord("BUY_access_4")%></div>
                <div class="show">xxxxx-xxxxx-78522 </div>
                <div class="clear"></div>
            </div>-->
                    <!--信用卡号码 end-->
                    <!--刷卡金额 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_access_5")%></div>
                        <div class="show">
                            <%=intPayMoney%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--刷卡金额 end-->
                    <!--返回结果 start-->
                    <div class="con">
                        <div class="name">
                            <%=getWord("BUY_access_6")%></div>
                        <div class="show color_1 bold">
                            <%=getWord("BUY_access_7")%></div>
                        <div class="clear">
                        </div>
                    </div>
                    <!--返回结果 end-->
                </div>
                <div class="w2">		
                    <input type="submit" name="contact_button" id="contact_button" value="+ <%=getWord("BUY_access_8")%>" class="btn_1" onclick="location.href='../contact/contact_us.asp'" />
                    <%  '判斷是否還有款項要繼續刷卡
						if (bol_is_payment_completed) then
                    %>
						<input type="submit" name="next_step_button" id="next_step_button" value="+ <%=getWord("BUY_2_27")%>" class="btn_1 m-left10" onclick="location.href='../member/contract.asp'" />
				</div>
					<% else %>
						<input type="submit" name="next_step_button" id="Submit1" value="+ <%="下一步"%>" class="btn_1 m-left10" onclick="location.href='../member/buy_1.asp'" /></div>
					<% end if %>					
		</div>
        <!--明細end-->
    </div>
</div>
<!--內容end-->
<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font></div>