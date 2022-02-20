<% 
'判斷金流渠道帳號密碼
'include檔盡量不要宣告(除非保證獨立變數)跟include其他檔案 
'dim intSalesCompany : intSalesCompany = ""
'dim intSalesDept : intSalesDept = ""
dim bolBJChannel : bolBJChannel = false
dim bolSHChannel : bolSHChannel = false
objStaffData = null
objContractData = null
arrParam = null
strSql = ""

'先不分渠道
'if ( isEmptyOrNull(intContractSn)  AND Not isEmptyOrNull(intHitrustSn) ) then
'    arrParam = Array(intHitrustSn)
'    strSql =  " SELECT ordernumber "
'    strSql =  strSql & " FROM hitrust_record_all WITH (NOLOCK) "
'    strSql =  strSql & " WHERE sn = @sn "
'    Set objContractData = excuteSqlStatementReadEach(strSql, arrParam, CONST_VIPABC_RW_CONN)
'    if ( true = bolDebugMode ) then
'        Response.Write("錯誤訊息Sql for 讀取刷卡筆數的合約編號:" & g_str_sql_statement_for_debug & "<br>")
'    end if
'    if ( Not objContractData.eof ) then
'        intContractSn = objContractData("ordernumber")
'    end if
'end if
'
'arrParam = Array(intContractSn)
'strSql =  " SELECT staff_basic.dept, cfg_dept.company "
'strSql =  strSql & " FROM client_temporal_contract WITH (NOLOCK) "
'strSql =  strSql & " INNER JOIN quick_client_record_last WITH (NOLOCK) ON client_temporal_contract.lead_sn = quick_client_record_last.client_sn "
'strSql =  strSql & " INNER JOIN staff_basic WITH (NOLOCK) ON quick_client_record_last.in_charge = staff_basic.sn "
'strSql =  strSql & " INNER JOIN cfg_dept WITH (NOLOCK) ON staff_basic.dept = cfg_dept.sn "
'strSql =  strSql & " WHERE client_temporal_contract.sn = @sn "
'Set objStaffData = excuteSqlStatementReadEach(strSql, arrParam, CONST_VIPABC_RW_CONN)
'if ( true = bolDebugMode ) then
'    Response.Write("錯誤訊息Sql for 讀取合約建單者上班公司別:" & g_str_sql_statement_for_debug & "<br>")
'end if
'if ( Not objStaffData.eof ) then
'    intSalesDept = objStaffData("dept")
'    intSalesCompany = objStaffData("company")
'end if
'
''11為VIPABC 北京與上海部門沒有分開 寫死判斷
'if ( 11 = sCInt(intSalesCompany) ) then
'    if ( Instr(CONST_DEPT_GROUP_NUMBER_VIPABC_BJ_BD, "," & sCInt(intSalesDept) & ",") > 0 ) then
'        bolBJChannel = true
'    else
'        bolSHChannel = true
'    end if
'end if

'寫死跑北京 for 測試帳號
'20121105 阿捨新增 寫死全面轉北京
bolBJChannel = true
'bolSHChannel = true

if ( true = bolBJChannel ) then
    '快錢(99BILL)帳戶判斷
    CONST_PAYMENT_PLATFORM_VIPABC_99BILL_NAME = CONST_PAYMENT_PLATFORM_VIPABC_BJ_99BILL_NAME
    CONST_CREDIT_CARD_99BILL_KEY = CONST_CREDIT_CARD_99BILL_BJ_KEY
    CONST_CREDIT_CARD_99BILL_ACCTID = CONST_CREDIT_CARD_99BILL_BJ_ACCTID
    '支付寶(ALIPAY)
    CONST_CREDIT_CARD_ALIPAY_CODE = CONST_CREDIT_CARD_ALIPAY_BJ_CODE    
    CONST_CREDIT_CARD_ALIPAY_KEY = CONST_CREDIT_CARD_ALIPAY_BJ_KEY
	CONST_CREDIT_CARD_ALIPAY_SELLER_EMAIL = CONST_CREDIT_CARD_ALIPAY_BJ_SELLER_EMAIL
	CONST_CREDIT_CARD_ALIPAY_SHOW_URL = CONST_CREDIT_CARD_ALIPAY_BJ_SHOW_URL
    '招商(CMBCHINA)
    CONST_CREDIT_CARD_CMBCHINA_KEY = CONST_CREDIT_CARD_CMBCHINA_BJ_KEY
    CONST_CREDIT_CARD_CMBCHINA_BANK_KEY = CONST_CREDIT_CARD_CMBCHINA_BJ_BANK_KEY
    '銀聯(CHINAPAY)
    CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_BJ_CODE
    CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_BJ_CODE
    CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_SUB_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_BJ_SUB_CODE
    CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_SUB_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_BJ_SUB_CODE
elseif ( true = bolSHChannel ) then
    '快錢(99BILL)帳戶判斷
    CONST_PAYMENT_PLATFORM_VIPABC_99BILL_NAME = CONST_PAYMENT_PLATFORM_VIPABC_SH_99BILL_NAME
    CONST_CREDIT_CARD_99BILL_KEY = CONST_CREDIT_CARD_99BILL_SH_KEY
    CONST_CREDIT_CARD_99BILL_ACCTID = CONST_CREDIT_CARD_99BILL_SH_ACCTID
    '支付寶(ALIPAY)
    CONST_CREDIT_CARD_ALIPAY_CODE = CONST_CREDIT_CARD_ALIPAY_SH_CODE    
    CONST_CREDIT_CARD_ALIPAY_KEY = CONST_CREDIT_CARD_ALIPAY_SH_KEY
	CONST_CREDIT_CARD_ALIPAY_SELLER_EMAIL = CONST_CREDIT_CARD_ALIPAY_SH_SELLER_EMAIL
	CONST_CREDIT_CARD_ALIPAY_SHOW_URL = CONST_CREDIT_CARD_ALIPAY_SH_SHOW_URL
    '招商(CMBCHINA)
    CONST_CREDIT_CARD_CMBCHINA_KEY = CONST_CREDIT_CARD_CMBCHINA_SH_KEY
    CONST_CREDIT_CARD_CMBCHINA_BANK_KEY = CONST_CREDIT_CARD_CMBCHINA_SH_BANK_KEY
    '銀聯(CHINAPAY)
    CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_SH_CODE
    CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_SH_CODE
    CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_SUB_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_CHINAPAY_SH_SUB_CODE
    CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_SUB_CODE = CONST_CREDIT_CARD_TYPE_VIPABC_UNIONPAY_SH_SUB_CODE
end if

if ( true = bolDebugMode ) then
	response.write "intSalesDept:" & intSalesDept & "<br />"
	response.write "intSalesCompany:" & intSalesCompany & "<br />"
	response.write "CONST_CREDIT_CARD_99BILL_KEY:" & CONST_CREDIT_CARD_99BILL_KEY & "<br />"
	response.write "CONST_CREDIT_CARD_99BILL_ACCTID:" & CONST_CREDIT_CARD_99BILL_ACCTID & "<br />"
    response.write "CONST_CREDIT_CARD_ALIPAY_CODE:" & CONST_CREDIT_CARD_ALIPAY_CODE & "<br />"
	response.write "CONST_CREDIT_CARD_ALIPAY_KEY:" & CONST_CREDIT_CARD_ALIPAY_KEY & "<br />"
	response.write "CONST_CREDIT_CARD_ALIPAY_SELLER_EMAIL:" & CONST_CREDIT_CARD_ALIPAY_SELLER_EMAIL & "<br />"
	response.end
end if
%>