<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->
<script type="text/javascript">
<!--
//列印的function
 function printElem(options)
 {
     $('#main_content').printElement(options);
 }

//如果按下不同意，把頁面跳回上緣
$(document).ready(function(){
    //如果選了列印，列印區塊檔案
    $("#print_contract").click(function() {
             printElem({overrideElementCSS: ['/css/css_cn_test.css']});
         });
    //如果選了不同意導到最上方
	$("#disagree").click(function(){
		$("#checkbox1").focus();
	});	
});

//======check 頁面上的checkbox是否全部勾選 ======
function checkdata(contract){
	//return false;
	for (var i=0; i<document.forms[0].elements.length; i++) {
		var e=document.forms[0].elements[i];			
		if (e.type=="checkbox"){
			if(e.checked!=1) {
				alert('<%=getWord("contract_1")%>');//请务必勾选所有条款
				document.getElementById(e.id).focus();
				return false;
			}	  
		}
	}
	alert("<%=getWord("contract_2")%>");//提醒您，堂数及帐号将于款项缴清后一并开通，若您的款项已缴清请忽略此讯息。
	document.f1.submit();
}
//-->
</script>
<%

dim bolTestCase : bolTestCase = false '測試狀態
dim bolDeBugMode : bolDeBugMode = false '除錯模式

'Dim int_contract_sn : int_contract_sn = Session("contract_sn")             '接Session-->合約編號
Dim int_contract_sn : int_contract_sn = ""
if ( isEmptyOrNull(int_contract_sn) ) then 
    int_contract_sn = getRequest("contract_sn", CONST_DECODE_NO) '無奈啊,我要嘆三聲了!!
end if

Dim str_tmp_go_page
'塞資料的相關設定值
Dim var_arr 																	'傳excuteSqlStatementRead的陣列
Dim arr_result																'接回來的陣列值
Dim str_tablecol 																'table欄位
Dim str_tablerow
Dim strCheckedBox : strCheckedBox = "" '是否已經點過合約 checkbox 	
Dim strDisabled : strDisabled = "" '是否已經點過合約 checkbox Disabled	
'---------- 重新找電子合約 ------- start -----
Dim str_contract_sql : str_contract_sql = ""
if ( isEmptyOrNull(int_contract_sn) ) then 

	str_contract_sql = " SELECT sn, cagree FROM client_temporal_contract WHERE (client_sn = @client_sn) AND (valid=1) AND (ctype = 1) ORDER BY sn DESC "
	var_arr = Array(g_var_client_sn)
	arr_result = excuteSqlStatementRead(str_contract_sql,var_arr,CONST_VIPABC_RW_CONN)
	'response.write g_str_sql_statement_for_debug&"<br />"
    if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then
			int_contract_sn = arr_result(0, 0)
            '如果已經點過合約
            if (arr_result(1, 0) = 1) then 
                strCheckedBox = "checked"
                strDisabled = "disabled"
            end if 
		else
			'Response.write "no data!!"
		end if
	end if
	set str_contract_sql = nothing
	set var_arr = nothing
	set arr_result = nothing
	'response.write g_str_sql_statement_for_debug&"<br />"
end if 
'判斷客戶的繳款狀態，繳清才能夠同意合約-----start----------
Dim bol_is_all_paid : bol_is_all_paid = false '是否繳清款項
'繳清 ture , 還沒 false
bol_is_all_paid = checkIsAllShouldPayClear(int_contract_sn,CONST_VIPABC_RW_CONN)
'判斷客戶的繳款狀態，繳清才能夠同意合約-----end----------
'---------- 重新找電子合約 ------- end -----

Dim str_sdate : str_sdate = ""              '合約開始日期
Dim str_edate : str_edate = ""                '合約結束日期
Dim str_product_cname : str_product_cname = ""             '購買產品中文名稱
Dim str_client_cname :str_client_cname = ""  '購買人中文名字
Dim str_client_ename :str_client_ename = ""  '購買人英文名字
Dim str_client_sex : str_client_sex = ""	  '購買人性別
Dim str_client_birth : str_client_birth =  ""     '購買人生日
Dim str_client_idno : str_client_idno =  ""        '購買人身份証字號
Dim str_client_id_email : str_client_id_email =  ""  '購買人帳號email
Dim str_client_email : str_client_email	=  ""  		'購買人通訊email
Dim str_client_cphone : str_client_cphone = ""		'購買人手機號碼
Dim str_client_caddr : str_client_caddr = ""		'購買人地址

Dim str_lname : str_lname = ""
Dim str_sulegal : str_sulegal = ""
Dim str_sale_name : str_sale_name = "" '銷售顧問
Dim str_contract_date : str_contract_date = "" '簽約日期
Dim int_fsession : int_fsession = "" 		'贈送堂數
Dim int_point_type : int_point_type = ""
Dim str_count_fsession : str_count_fsession = 0 		'計算過後的贈送堂數
Dim int_osession : int_osession = ""
Dim int_point_fee : int_point_fee = ""
Dim str_count_osession : str_count_osession = ""
Dim str_total_session : str_total_session = 0 			'總堂數
Dim str_ptype : str_ptype = "" 			'付款方式：「一次付清1」、「按月分期3」、「分期付款」
Dim str_discount_type : str_discount_type = "" 
Dim int_total_money : int_total_money = 0 'show總額
Dim str_total_money : str_total_money = "" '中文錢字樣
'Dim str_freemon
Dim str_valid_month : str_valid_month =  "" '合約使用月份
Dim str_try_sdate : str_try_sdate =  "" 
Dim str_try_7date : str_try_7date = "" '七日審閱日期
Dim str_try_8date : str_try_8date = "" '第八日
Dim str_try_30date : str_try_30date = "" '30日
Dim str_is_extending_contract : str_is_extending_contract = "" 
Dim str_payman : str_payman = ""  '付款人
Dim str_freemon : str_freemon = ""
Dim str_tmp_edate : str_tmp_edate = ""
Dim int_account_cycle_max_session : int_account_cycle_max_session = "" '一周幾堂
Dim int_week_count : int_week_count = ""	'共幾周
Dim bol_is_cagree : bol_is_cagree = false '是否已經點過合約
Dim str_fd_agree_name : str_fd_agree_name = "" '開通人員
Dim date_fd_agree_date : date_fd_agree_date = ""
Dim date_contract_start_date : date_contract_start_date = "" '合約開始日期 標準日期格式'
Dim int_account_cycle_days : int_account_cycle_days = "" '循環的天數
'===================================================
'找出開通的財務人員

'===================================================
'算出這筆合約有沒有折讓Begin
Dim str_fd_sql
str_fd_sql = " SELECT dbo.staff_basic.fname + ' ' + dbo.staff_basic.lname AS auditor ,cadate FROM "
str_fd_sql = str_fd_sql&" ( SELECT auditor,cadate FROM client_temporal_contract WHERE sn = @sn ) AS client_temporal_contract1 "
str_fd_sql = str_fd_sql&" LEFT JOIN staff_basic ON staff_basic.sn = client_temporal_contract1.auditor "
var_arr = Array(int_contract_sn)
arr_result = excuteSqlStatementRead(str_fd_sql, var_arr, CONST_VIPABC_RW_CONN)
'response.write g_str_sql_statement_for_debug
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
        str_fd_agree_name = arr_result(0,0)
        date_fd_agree_date = getFormatDateTime(arr_result(1,0),5)
    end if
end if
Dim discount_sql	
discount_sql = " SELECT ISNULL(SUM(amount), 0) AS discount FROM client_payment_detail WHERE (payrecord_sn IN (SELECT sn FROM client_payment_record WHERE (contract_sn = @contract_sn))) AND (payment_mode = 14  or payment_mode = 28)"

	var_arr = Array(int_contract_sn)
	arr_result = excuteSqlStatementRead(discount_sql,var_arr,CONST_VIPABC_RW_CONN)
if ( isSelectQuerySuccess(arr_result) ) then
	if (Ubound(arr_result) >= 0) then
		'Response.write("欄:" & Ubound(arr_result) & "<br />")
		'Response.write("列:" & Ubound(arr_result,2) & "<br />")	
		'response.write g_str_sql_statement_for_debug

		for str_tablerow = 0 To Ubound(arr_result, 2)
			Dim int_discount : int_discount = 0                                         'discount		
			int_discount = arr_result(0, str_tablerow)
		next
	else
			int_discount = 0
	end if
end if
set discount_sql = nothing
set var_arr = nothing
set arr_result = nothing
'算出這筆合約有沒有折讓End		
'==========================================================

'===================================================
'Online Credit Begin

'20120815 阿捨新增 VM12081018 8月行銷活動五款產品 送New iPad 更動直購頁面與產品合約
'找出贈送ipad紀錄
arrParam = array(int_contract_sn)
dim bolNewIPadGift : bolNewIPadGift = false
strSql = " SELECT CfgActivitySn FROM dbo.cfg_contract_activity WITH (NOLOCK) WHERE ContractSn = @ContractSn AND Valid = 1 AND CfgActivitySn = 12 "
Set objCfgProduct = excuteSqlStatementReadEach(strSql , arrParam, CONST_TUTORABC_RW_CONN)
if ( true = bolDeBugMode ) then
    Response.Write("錯誤訊息Sql for 找出贈送ipad紀錄:" & g_str_sql_statement_for_debug & "<br />")
end if
if ( not objCfgProduct.eof ) then
    bolNewIPadGift = true
end if

'if (int_contract_sn <> "") then
	'20090121 Jessica Mod by BD09011223  線上刷卡要可瀏覽合約----- start ---------
	Dim str_online_credit : str_online_credit = "N"   
	Dim str_pay_contract_sn : str_pay_contract_sn = "" '線上刷卡的合約編號 
	str_pay_contract_sn = getClientBrushCardSn(int_contract_sn ,"contract")	
	if ( str_pay_contract_sn <> 0 ) then 
		int_contract_sn = str_pay_contract_sn
		str_online_credit = "Y"
	end if 
	
'end if 
	'20090121 Jessica Mod by BD09011223  線上刷卡要可瀏覽合約----- end ---------
'Online Credit  End
'==========================================================

'==========================================================
'抓取合約相關資料 Begin 
'ex:合約起訖日、產品名稱、應付XX元
'20120725 阿捨修改 短期合約 只能抓valid_days(<30) 算contract_edate
Dim contract_sql
contract_sql = "SELECT DERIVEDTBL_2.contract_sn, /* 0 */ "&_
				"client_temporal_contract.ptype, "&_
				"client_temporal_contract.lead_sn, "&_
				"lead_basic.oaddr, "&_
				"client_temporal_contract.client_sn, "&_
				"cfg_product.cname AS product, /* 5 */"&_
				"cfg_product.product_name AS product_nickname, "&_
				"lead_basic.cname, "&_
				"lead_basic.fname + ' ' + lead_basic.lname AS ename, "&_
				"lead_basic.caddr, "&_
				"lead_basic.htel_code + '-' + lead_basic.htel + (CASE isnull(lead_basic.htel_branch, '') WHEN '' THEN NULL ELSE ' #' + lead_basic.htel_branch END) AS hphone, /* 10 */"&_
				"CONVERT(varchar, client_temporal_contract.inpdate, 111) AS cdate, "&_
				"CAST(YEAR(client_temporal_contract.inpdate) AS varchar(4)) + ' 年 ' + CAST(MONTH(client_temporal_contract.inpdate) AS varchar(4)) + ' 月 ' + CAST(DAY(client_temporal_contract.inpdate) AS varchar(4)) + ' 日' AS sudate, "&_
				"CAST(YEAR(client_temporal_contract.sdate) AS varchar(4)) + ' 年 ' + CAST(MONTH(client_temporal_contract.sdate) AS varchar(4)) + ' 月 ' + CAST(DAY(client_temporal_contract.sdate) AS varchar(4)) + ' 日' AS sdate, "&_
				"case cfg_product.product_status "&_
				"WHEN '2' then "&_
				" CAST(YEAR(DATEADD(d, - 0, DATEADD(d, cfg_product.valid_month * cfg_product.account_cycle_day, client_temporal_contract.sdate))) AS varchar(4)) + ' 年 ' + CAST(MONTH(DATEADD(d, - 0, DATEADD(d, cfg_product.valid_month * cfg_product.account_cycle_day, client_temporal_contract.sdate))) AS varchar(4)) + ' 月 ' + CAST(DAY(DATEADD(d, - 0, DATEADD(d, cfg_product.valid_month * cfg_product.account_cycle_day, client_temporal_contract.sdate))) AS varchar(4)) + ' 日' "&_
                "else " &_
                " case WHEN cfg_product.valid_days < 30 then " &_
                " CAST(YEAR(DATEADD(d, - 0, DATEADD(d, cfg_product.valid_days, client_temporal_contract.sdate))) AS varchar(4)) + ' 年 ' + CAST(MONTH(DATEADD(d, - 0, DATEADD(d, cfg_product.valid_days, client_temporal_contract.sdate))) AS varchar(4)) + ' 月 ' + CAST(DAY(DATEADD(d, - 0, DATEADD(d, cfg_product.valid_days, client_temporal_contract.sdate))) AS varchar(4)) + ' 日' "&_
                "else "&_
				" CAST(YEAR(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + ' 年 ' + CAST(MONTH(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + ' 月 ' + CAST(DAY(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + ' 日' "&_
				"end "&_
                "end "&_
				" AS edate, /* 15 */"&_
				"cfg_product.total_points / 65 AS osession, "&_
				"lead_basic.sex, "&_
				"lead_basic.company, "&_
				"lead_basic.email, "&_
				"lead_basic.client_address AS Expr1, /* 20 */"&_
				"lead_basic.cemail, "&_
				"lead_basic.caddr, "&_
				"lead_basic.legal_name, "&_
				"lead_basic.legal_idno, "&_
				"lead_basic.legal_relations, /* 25 */"&_
				"lead_basic.legal_tel, "&_
				"staff_basic.fname + ' ' + staff_basic.lname AS agent, "&_
				"(CASE lead_basic.sex WHEN 1 THEN '男' ELSE '女' END) AS sex, "&_
				"lead_basic.idno, CAST(YEAR(lead_basic.birth) AS varchar(4)) + ' 年 ' + CAST(MONTH(lead_basic.birth) AS varchar(4)) + ' 月 ' + CAST(DAY(lead_basic.birth) AS varchar(4)) + ' 日' AS birth, "&_
				"lead_basic.cphone_code + '-' + lead_basic.cphone AS cphone, /* 30 */"&_
				"lead_basic.contact_way, "&_
				"lead_basic.company AS Expr2, "&_
				"lead_basic.otel_code + '-' + lead_basic.otel + (CASE isnull(lead_basic.otel_branch, '') WHEN '' THEN NULL ELSE ' #' + lead_basic.otel_branch END) AS ophone, "&_
				"cfg_product.total_cash, "&_
				"cfg_product.pays_off, /* 35 */"&_
				"cfg_product1.first_cash, "&_
				"DERIVEDTBL_2.paytype, "&_
				"DERIVEDTBL_2.payman, "&_
				"lead_basic.ccode, "&_
				"cfg_post.Data1, /* 40 */"&_
				"cfg_city.Data1 AS Expr3, "&_
				"sb.fname + ' ' + sb.lname AS sname, "&_
				"cfg_product.point_type, "&_
				"cfg_session_points.Point_Fee, "&_
				"DERIVEDTBL_1.sum_gift AS fsession, /* 45 */"&_
				"client_temporal_contract.prodbuy , "&_
				"cfg_product.valid_month, "&_
				"client_temporal_contract.discount_type, "&_
				"cfg_product.is_extending_contract, "&_
				"client_freemonth.freegift AS freemon, /* 50 */"&_
				"CAST(YEAR(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, "&_
				"client_temporal_contract.sdate))) AS varchar(4)) + '/' + CAST(MONTH(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, "&_
				"client_temporal_contract.sdate))) AS varchar(4)) + '/' + CAST(DAY(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) AS date_edate , "&_
				"cfg_product.account_cycle_max_session, "&_
				"cfg_product.week_count, "&_
				"client_temporal_contract.cagree , "&_
				"cfg_product.account_cycle_day , /* 55 */"&_
				"lead_basic.c_province, "&_
				"lead_basic.c_area, "&_
				"lead_basic.c_region, "&_
                "client_temporal_contract.prodbuy "&_
                ",client_temporal_contract.ProductMonthReward, /* 60 */"&_
				"cfg_product.contract_type, "&_
                "cfg_product.renew "&_
				"FROM cfg_city "&_
				"RIGHT Outer JOIN cfg_post "&_
				"RIGHT Outer JOIN lead_basic ON cfg_post.Data2 = lead_basic.ccode ON cfg_city.sn = cfg_post.city_sn "&_
				"RIGHT OUTER JOIN staff_basic "&_
				"RIGHT OUTER JOIN client_temporal_contract "&_
				"LEFT OUTER JOIN client_freemonth ON client_temporal_contract.sn = client_freemonth.contract_sn ON staff_basic.sn = client_temporal_contract.inputer ON lead_basic.sn = client_temporal_contract.lead_sn "&_
				"LEFT OUTER JOIN cfg_session_points "&_
				"INNER JOIN cfg_product ON cfg_session_points.Sn = cfg_product.point_type ON client_temporal_contract.prodbuy = cfg_product.sn "&_
				"RIGHT OUTER JOIN ( SELECT contract_sn, SUM(freegift) AS sum_gift "&_
				"FROM client_freegift "&_
				"WHERE (contract_sn = @contract_sn_1) "&_
				"GROUP BY contract_sn ) AS DERIVEDTBL_1 "&_
				"RIGHT OUTER JOIN ( SELECT contract_sn, paytype, payman, handling "&_
				"FROM client_payment_record "&_
				"WHERE (contract_sn = @contract_sn_2) ) AS DERIVEDTBL_2 ON DERIVEDTBL_1.contract_sn = DERIVEDTBL_2.contract_sn ON client_temporal_contract.sn = DERIVEDTBL_2.contract_sn "&_
				"LEFT OUTER JOIN staff_basic AS sb ON DERIVEDTBL_2.handling = sb.sn left outer JOIN cfg_product as cfg_product1 ON client_temporal_contract.old_prodbuy = cfg_product1.sn "
'20120918 阿捨新增 讓退費合約也可顯示(valid=2)
if (str_online_credit <> "Y") then
	contract_sql=contract_sql & "WHERE (client_temporal_contract.valid = 1) OR (client_temporal_contract.valid = 2)"
end if
	var_arr = Array(int_contract_sn,int_contract_sn)
	arr_result = excuteSqlStatementRead(contract_sql,var_arr,CONST_VIPABC_RW_CONN)
    if ( true = bolDeBugMode ) then
       Response.Write("錯誤訊息sql for 抓取合約相關資料:" & g_str_sql_statement_for_debug & "<br />") 
    end if
'response.write g_str_sql_statement_for_debug
if (isSelectQuerySuccess(arr_result)) then
    '有資料
    if (Ubound(arr_result) > 0) then    
		str_sdate = arr_result(13,0)                '合約開始日期
        date_contract_start_date = arr_result(13,0) '合約開始日期
		str_edate = arr_result(14,0)                '合約結束日期
        '20120607 阿捨新增 合約建單日判斷相關條款
        strInputDate = arr_result(11,0)  '合約建單日
		str_product_cname = arr_result(5,0)             '購買產品中文名稱
		str_client_cname = arr_result(7,0)  '購買人中文名字
		str_client_ename = arr_result(8,0)  '購買人英文名字
		str_client_sex = arr_result(27,0)	  '購買人性別
		str_client_birth =  arr_result(29,0)      '購買人生日
		str_client_idno =  arr_result(28,0)         '購買人身份証字號
		str_client_id_email =  arr_result(18,0)  '購買人帳號email
		str_client_email	=  arr_result(20,0)  		'購買人通訊email
		str_client_cphone = arr_result(30,0)		'購買人手機號碼
		str_client_caddr = arr_result(9,0)		'購買人地址
		
		str_lname = arr_result(22,0)
		str_sulegal = ""
		str_sale_name = arr_result(42,0) '銷售顧問
		str_contract_date = arr_result(12,0) '簽約日期
		int_fsession = arr_result(45,0) 		'贈送堂數
		int_point_type = arr_result(43,0)
		str_count_fsession = 0 		'計算過後的贈送堂數
		int_osession = arr_result(15,0)
		int_point_fee = arr_result(44,0)
		str_count_osession = ""
		str_total_session = 0 			'總堂數
		str_ptype = arr_result(37,0) 			'付款方式：「一次付清1」、「按月分期3」、「分期付款」
		str_discount_type = arr_result(48,0) 
		int_total_money = 0 'show總額
		str_total_money = "" '中文錢字樣
		str_freemon = arr_result(50,0)
		str_tmp_edate = arr_result(51,0)
		str_valid_month =  arr_result(47,0) '合約使用月份
		str_try_sdate =  dateadd("d",0,str_sdate) 
		str_try_7date = dateadd("d",6,str_try_sdate) '七日審閱日期
		str_try_8date = dateadd("d",8,str_try_sdate) '第八日
		str_try_30date = dateadd("d",30,str_try_sdate) '30日
		str_is_extending_contract = arr_result(49,0) 
		str_payman = arr_result(38,0)  '付款人
		'定時定額
		int_account_cycle_max_session =  arr_result(52,0)  '一周上幾堂
		int_week_count =  arr_result(53,0) 	'共幾周
        int_account_cycle_days =  arr_result(55,0)
        '是否已經點過合約
		if (arr_result(54,0) = 1)  then    
            bol_is_cagree = true
        end if 	
		strLegalName = arr_result(56,0)

		'------- 贈送堂數計算 ------- start -------
		if ( not isEmptyOrNull(int_fsession) ) then 
		    int_fsession = sCInt(int_fsession)
			if ( isnumeric(int_fsession) and int_fsession > 0 ) then
				if ( int_point_type = 1 ) then
					str_count_fsession =((int_fsession / 65) / 4)
				elseif ( int_point_type = 7 ) then
					str_count_fsession =((int_fsession / 65) / 3)
				else
					str_count_fsession =( int_fsession / 65 )
				end if
			else
				str_count_fsession = 0
			end if	
		end if 
		'------- 贈送堂數計算 ------- end -------

		'------- 總共堂數計算 ------- start -------
		if ( int_point_type = 1 or int_point_type = 7 ) then 
			str_count_osession = ( int_osession /  int_point_fee / 65 )
		else
			str_count_osession = int_osession
		end if 
		
		str_total_session = str_count_osession + str_count_fsession
		'------- 總共堂數計算 ------- end -------
		
		'------ 金額 ------ start ---------
		if ( isEmptyOrNull(str_discount_type) ) then 
            str_discount_type = "" 
        end if 

		if ( ( str_ptype = "1" or str_ptype = "4" ) and str_discount_type <> "S") then 
			int_total_money = arr_result(35,0) 'pays_off
		else
			int_total_money = arr_result(34,0)  'total_cash
		end if 
		
        
		'把總價扣掉折讓就是應出現的產品金額
		int_total_money = int_total_money - int_discount
        str_total_money = int_total_money
		'------ 金額 ------ end ---------
		
		'----- 服務期間 ----- start -------
		
		str_sdate = "自 "&str_sdate&" 起，"	
		str_edate = "至 "&str_edate&" 止。"
		
		if (not isEmptyOrNull(str_freemon) ) then
			str_tmp_edate = dateadd("m",str_freemon,str_tmp_edate)
			str_edate=year(str_tmp_edate)&" 年 "&month(str_tmp_edate)&" 月 "&day(str_tmp_edate)&" 日"	
		else
			str_edate = str_edate
		end if	
		'----- 服務期間 ----- end -------
		
		if ( str_client_idno <> "" ) then 
            str_client_idno = "******"&right(str_client_idno,4) 
        end if 

		'付款人
		if ( isEmptyOrNull(Request("payman"))  or str_payman = "" )  then 
            str_payman = str_client_cname 
        else 
            str_payman = str_payman 
        end if 
		
        '法定代理人
		if ( not isEmptyOrNull(str_lname) )  then 
            str_sulegal = str_lname 
        else 
            str_sulegal="&nbsp;"
		end if
    else
        '無資料
    end if 
	
	'20100705 阿捨修正 大陸三層 地區欄位
	dim intProvinceNumber, intAreaNumber, intRegionNumber '三層碼
	dim strValueName
	intProvinceNumber = arr_result(56,0) 
	intAreaNumber = arr_result(57,0) 
	intRegionNumber = arr_result(58,0) 
	
	if (  Not isEmptyOrNull(intProvinceNumber) AND Not isEmptyOrNull(intAreaNumber) AND Not isEmptyOrNull(intRegionNumber) ) then
		strPath = Server.Mappath("/xml/ChinaProvinceAndArea.xml")
		strValueName = findXmlValueName(strPath, intProvinceNumber, intAreaNumber, intRegionNumber)
		str_client_caddr =  strValueName & str_client_caddr
	end if
	
    '20120316 阿捨新增 無限卡產品判斷	
    intProductSn = arr_result(59,0) 
    dim bolLifeUnlimited : bolLifeUnlimited = false
    dim bolUnlimited : bolUnlimited = false
   
    if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (intProductSn & ",")) > 0 ) then
        bolUnlimited = true 
    end if
    if ( Instr(CONST_NO_LIMIT_FOREVER_PRODUCT_SN, (intProductSn & ",")) > 0 ) then
        bolLifeUnlimited = true  '終生無限卡
    end if

    '20120419 阿捨新增 無限卡贈送期數
    dim intProductMonthReward : intProductMonthReward = arr_result(60,0)
    dim bolProductMonthReward : bolProductMonthReward = false
    if ( Not isEmptyOrNull(intProductMonthReward) ) then
        if ( sCInt(intProductMonthReward) > 0 ) then
            bolProductMonthReward = true
        end if
    end if

    dim bolComboProduct : bolComboProduct = false
    '20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493                    
    if ( Instr(CONST_DIANPING_GROUPBUY_PRODUCT_SN, (intProductSn & ",")) > 0 ) then
            bolComboProduct = true
    end if

	'20120606 VM12053006 學習儲值卡官網+IMS功能 Penny
	Dim intContractType : intContractType = arr_result(61,0)
	if ( Not isEmptyOrNull(intContractType) ) then
        if( 2 = sCInt(intContractType) ) then
			'response.redirect "http://"&const_vipabc_webhost_name&"/program/member/ajax_short_contract.asp?ContractType="&intContractType & "&ClientSn=" & arr_result(4,0) & "&ContractSn=" & int_contract_sn
            response.redirect "/program/member/ajax_short_contract.asp?ContractType="&intContractType & "&ClientSn=" & arr_result(4,0) & "&ContractSn=" & int_contract_sn
			response.end
		end if
    end if
	
    '20120420 阿捨新增 業績結算日顯示 (找出下個月1號減去一天)
    dim strNowMaxDate : strNowMaxDate = dateadd("d", -1, (left(getFormatDateTime(dateadd("m", 1, getFormatDateTime(str_contract_date,5)),5),8) & "01"))
    dim arrNowMaxDate : arrNowMaxDate = split(getFormatDateTime(strNowMaxDate,5),"/")

     '20120515 tree 確認產品為新產品還是舊產品
    Dim strProductNewOrOld : strProductNewOrOld = "" '取得產品是否為舊(1)還是新(2) ，VIPABC 舊產品客戶訂課時間為8小時前(1對1扣2堂) 新產品客戶訂課時間為24小時前(1對1扣3堂)
    strProductNewOrOld = getProductNewOrOld(intProductSn)
    dim bolNewProduct : bolNewProduct = false
    if ( "2" = strProductNewOrOld ) then
        bolNewProduct = true
    end if

    '20120816 阿捨新增 續約判斷
    dim intRenewStatus :  intRenewStatus = arr_result(62, 0)
    dim bolRenewStatus : bolRenewStatus = false
    if ( 1 = sCInt(intRenewStatus) ) then
        bolRenewStatus = true
    end if

    dim bolScholarship : bolScholarship = false
     '20120821 阿捨新增 VG12081712 VIPABC獎學金活動
    if ( datediff("d","2012/08/22",strInputDate) >= 0 AND datediff("d","2012/11/01",strInputDate) <= 0 ) then
        if ( Instr(CONST_VIPABC_SCHOLARSHIP_PRODUCT_SN, (intProductSn & ",")) > 0  AND false = bolNewIPadGift ) then
                bolScholarship = true
        end if
    end if

    dim bolNewContract : bolNewContract = false
    '20120927 Tree GM12092503 更新VIPABC的合約內容 要20120928當天候漸單的才要做修改
    'if ( datediff("d","2012/09/28",strInputDate) >= 0 ) then TREE 要恢復
    if ( datediff("d","2012/10/09",strInputDate) >= 0 ) then
        bolNewContract = true
    end if
end if 

if ( true = bolDeBugMode ) then
    response.Write "strInputDate: " & strInputDate & "<br />"
    response.Write "bolNewContract: " & bolNewContract & "<br />"
    response.Write "datediff(d,2012/08/18,strInputDate): " & datediff("d","2012/08/18",strInputDate) & "<br />"
end if

'20121031 阿捨新增 上海移除改為北京 順便新增html合約呈現方法
dim bolChangeContractView : bolChangeContractView = false
if ( datediff("d","2012/11/05",strInputDate) >= 0 ) then
    bolChangeContractView = true
end if

'20121026 阿捨新增 公關產品 不導頁
dim bolPRFreeClass : bolPRFreeClass = false
if ( Instr(CONST_VIPABC_PR_FREE_PRODUCT_SN, (intProductSn & ",")) > 0 ) then
    bolPRFreeClass = true 
end if

'20121026 阿捨新增 員工免費課程 不導頁
dim bolStaffFreeClass : bolStaffFreeClass = false
if ( Instr(CONST_VIPABC_STAFF_FREE_PRODUCT_SN, (intProductSn & ",")) > 0 ) then
    bolStaffFreeClass = true 
end if

if ( true = bolChangeContractView AND false = bolPRFreeClass AND false = bolStaffFreeClass ) then
    response.redirect "/program/member/ajax_contract_new.asp?contract_sn=" & int_contract_sn
    response.end
end if

dim arrCountChinaWord : arrCountChinaWord = array("", "一", "二", "三", "四", "五", "六", "七")
%>
<!-- Menber Content Start -->
<div class="main_con" id="main_content">
<!--內容start-->
<div class="terms_main2">
	<form lang='zh' name="f1" id="f1" method="post" action='contract_agree.asp'>
	
    <div cla ss="line01">
    </div>
    <div class="title2 t-center big">
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
        会员权益概要
        <% else %>
        VIPABC会员权益概要
        <% end if %>
        </div>
    <div>
        <div class="check">
            <input type="checkbox" name="checkbox" id="checkbox" <%=strCheckedBox&" "&strDisabled%> /></div>
        <div class="law">
            本会员了解并同意：<br />
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td width="4%" align="left" valign="top">
                        一、
                    </td>
                    <td width="96%" align="left">
                        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                            麦奇教育集团提供
                        <% else %>
                            所购买的产品内容为
                        <% end if %>
                        1~3人精致小班制
                        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                            在线
                        <% end if %>
                        教学服务
                        <% '20120316 阿捨新增 無限卡產品判斷合約內容 %>
                        <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                            。
                        <% else %>
                            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                                ，每次扣抵一课时。但一对一在线教学服务，计费是以三次小班制课程兑换一次一对一课程（即扣抵三课时）。
                            <% else %>
                                <% if ( true = bolNewProduct ) then %><%'新舊產品判斷 %>
                                    ，每次咨询将扣抵一课时。但一对一教学服务，计费是以三次小班制课程兑换一次一对一课程（即扣抵三课时）。
                                <% else %>
                                    ，每次咨询将扣抵一课时。但一对一教学服务，计费是以二次小班制课程兑换一次一对一课程（即扣抵二课时）。
                                <% end if %>
                            <% end if %>
                        <% end if %>
                    </td>
                </tr>
                <tr>
                    <td align="left" valign="top">
                        二、
                    </td>
                    <td align="left">
                        <% if ( true = bolNewContract ) then %>
                            麦奇教育集团亦提供单向教学为主之英语学习在线大会堂服务（每次使用抵扣一堂课程）。注：大会堂是以专业课程为主（例如商业谈判、商业简报、发音、语法课程等等）的单向教学式之30人以上大讲堂。
                        <% else %>
                            本公司并有提供单向教学为主之英语学习大会堂服务（每次使用扣抵一堂课程）。注：大会堂是以专业课程为主（例如商业谈判、商业简报、发音、语法课程等等）的单向教学式之30人以上大讲堂。
                        <% end if %>
                    </td>
                </tr>
                <tr>
                    <td align="left" valign="top">
                        三、
                    </td>
                    <td align="left">
                        <% if ( true = bolNewContract ) then %>
                            每次在线课程时间为45分钟
                        <% else %>
                            每次上课的时间为45分钟
                        <% end if %>
                        <% if ( true = bolNewProduct ) then %><%'新舊產品判斷 %>
                            <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                                （小班制、大会堂课程皆是）。
                            <% else %>
                                <% if( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                                    （小班制、大会堂课程皆是）。麦奇教育集团所提供学习课程方式为每固定期间排定固定之课时数，会员若未能于当期使用完毕，未使用课时数将自动归零，不会累计至下一期。惟经麦奇教育集团同意，会员得预支使用其预先排定之课时数。
                                <% else %>
                                    （小班制、一对一、大会堂课程皆是）。本公司所提供学习课程方式为每周期排定固定之课时数，会员若未能于当周期使用完毕，未使用课时数将自动归零，不会累计至次周期。惟经本公司同意，会员得预支使用其预先排定之课时数。
                                <% end if %>
                            <% end if %>
                        <% else %>
                            <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                                （小班制、一对一、大会堂课程皆是）。<%'以前合約本來就錯了，但客戶已確認合約，所以一對一在合約中不能移除，避免客戶拷貝頁面那就會和官網合約對不起來 %>
                            <% else %>
                                （小班制、一对一、大会堂课程皆是）。本公司所提供学习课程方式为每周期排定固定之课时数，会员若未能于当周期使用完毕，未使用课时数将自动归零，不会累计至次周期。惟经本公司同意，会员得预支使用其预先排定之课时数。
                            <% end if %>
                        <% end if %>
                    </td>
                </tr>
                <tr>
                    <td align="left" valign="top">
                        四、
                    </td>
                    <td align="left">
                        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                            麦奇教育集团提供的在线教学服务由北京创意麦奇教育信息咨询有限公司通过其运营的网站(www.vipabc.com)提供。会员无法指定特定顾问为其提供服务。
                        <% else %>
                            本公司提供多国籍语言顾问进行语言咨询服务，但会员无法指定特定顾问进行咨询服务。
                        <% end if %>
                    </td>
                </tr>
            </table>
        </div>
        <div class="clear">
        </div>
    </div>
    <div>
        <div class="check">
            <input type="checkbox" name="checkbox1" id="checkbox1" <%=strCheckedBox&" "&strDisabled%> /></div>
        <div class="law">
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                本会员了解并同意会员使用麦奇教育集团会籍权限账号须于全数缴清其与上海麦奇教育信息咨询有限公司签订的咨询服务合同书的咨询服务费后始开通。
            <% else %>
                本会员了解并同意会员使用权限账号须全数缴清所购买产品款项后始开通。
            <% end if %>
        </div>
        <div class="clear">
        </div>
    </div>
    <div>
        <div class="check">
            <input type="checkbox" name="checkbox2" id="checkbox2" <%=strCheckedBox&" "&strDisabled%> /></div>
        <div class="law">
            <% if ( true = bolNewProduct ) then %><%'新舊產品判斷 %>
                <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                    <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                        本会员了解并同意使用咨询服务及麦奇教育集团在线教学服务时，最少需在课程服务开始前24小时通过网络进行课程预约，如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程。若未于课前四小时取消课程，则该课程仍计入已使用课时数。
                    <% else %>
                        本会员了解并同意最少需在课程服务开始前24小时通过网络进行课程预约，如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程。若未于课前四小时取消课程，则该课程仍计入已使用课时数。
                    <% end if %>
                <% else %>
                    <%if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                        本会员了解并同意使用咨询服务及麦奇教育集团在线教学服务时，最少需在课程服务开始前24小时通过网络进行课程预约，若于课前12至24(不含)小时内预约者，该课程将以1.25课时数计算；若于课前4至12(不含)小时内预约者，该课程将以1.5课时数计算；若于课前2至4(不含)小时内预约者，该堂课将以2课时数计算。如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程。若未于课前四小时取消课程，则该课程仍计入已使用课时数。
                    <% else %>
                        本会员了解并同意最少需在课程服务开始前24小时透过网络进行课程预约，若于课前12至24(不含)小时内预约者，该课程将以1.25课时数计算；若于课前4至12(不含)小时内预约者，该课程将以1.5课时数计算；若于课前2至4(不含)小时内预约者，该堂课将以2课时数计算。如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程，若未于课前四小时取消课程，则该课程仍计入已使用课时数。
                    <% end if %>
                <% end if %>
            <% else %>
                <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                    本会员了解并同意使用咨询服务及麦奇教育集团在线教学服务时，最少需在课程服务开始前八小时通过网络进行课程预约，如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程。若未于课前四小时取消课程，则该课程仍计入已使用课时数。
                <% else %>
                    本会员了解并同意最少需在课程服务开始前八小时通过网络进行课程预约，如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程。若未于课前四小时取消课程，则该课程仍计入已使用课时数。
                <% end if %>
            <% end if %>
            </div>
        <div class="clear">
        </div>
    </div>
    <div>
        <div class="check">
            <input type="checkbox" name="checkbox3" id="checkbox3" <%=strCheckedBox&" "&strDisabled%> /></div>
        <div class="law">
            <% if ( true = bolRenewStatus ) then %><%'續約判斷 %>
                本会员了解此系以续约方式购买本产品，因此本合约开通后，即不得退费。
            <% else %>
                本会员了解并同意于
                <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                    支付咨询服务费后如于七日内办理退费，除扣除行政管理费人民币500元，并应扣除已使用之课程的费用；第八日起至第三十日止办理退费，除扣除行政管理费人民币3,000元，并应扣除此期间之课程费用。支付咨询服务费逾一个月者，不得退费。无论因何原因导致退款，本会员同意支付退款赔偿金，该退款赔偿金的金额不应少于《会员服务使用条款》第十二条约定的利息补偿金。
                <% else %>
                    购买产品后如于七日内办理退费，除扣除行政管理费500元，并应扣除已使用之课程费用；第八日起至第三十日止办理退费，除扣除行政管理费3,000元，并应扣除此期间之课程费用。购买产品逾一个月者，则不得退费。无论因何原因导致退款，本会员同意支付退款赔偿金，该退款赔偿金的金额不应少于合同书第十二条约定的利息补偿金。
                <% end if %>
                <% if ( true = bolNewIPadGift ) then %>且会员若于合约期间内申请退费且已收取赠品者，则须按赠品之价格返回款项予本公司(New iPad以人民币3,688元整计算)。<% end if %>
            <% end if %>
            </div>
        <div class="clear">
        </div>
    </div>
    <div>
        <div class="check">
            <input type="checkbox" name="checkbox4" id="checkbox4" <%=strCheckedBox&" "&strDisabled%> /></div>
        <div class="law">
            本会员了解并同意会员不得将
            <% if ( true = bolNewContract ) then%><%'更新VIPABC的合約內容%>
                麦奇教育集团
            <% end if %>
            会籍出售、转让给予他人使用，非经
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                麦奇教育集团
            <% else %>
                本公司
            <% end if %>
            同意私下出售、转让会造成丧失会员资格且不得请求退还未使用课程之费用；本会员确因情况特殊有会籍移转
            之必要，应依照
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                《会员服务使用条款》向麦奇教育集团申请
            <% else %>
                合同内载明之服务条款向本公司
            <% end if %>
            办理。</div>
        <div class="clear">
        </div>
    </div>
    <div>
        <div class="check">
            <input type="checkbox" name="checkbox5" id="checkbox5" <%=strCheckedBox&" "&strDisabled%> /></div>
        <div class="law">
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                本会员了解并同意会员进行在线课程学习，是依会员个人需求将课时数与时间搭配，并已预留相当时间供弹性调整，会员购买前应详细评估判断，选择最合适自己方案及期限。因此视为会员同意咨询之服务期及会籍，不得延期。
            <% else %>
                本会员了解并同意会员购买之产品，是依会员个人需求将咨询课时数与时间搭配，并已预留相当时间供弹性调整，会员购买前应详细评估判断，选择最合适自己方案及期限。因此视为会员同意咨询之服务期，不得延期。
            <% end if %>
        </div>
        <div class="clear">
        </div>
    </div>
    <br />
    <div class="dode">
    </div>
    <div class="title2 big t-center">
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            咨询服务合同书
        <% else %>
            VIPABC会员合同书
        <% end if %>
    </div>
    <div style="margin-top: 15px;">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td class="t1">
                    中文姓名
                </td>
                <td class="t2">
                    <%=str_client_cname%>
                </td>
                <td class="t1">
                    英文姓名
                </td>
                <td class="t2">
                    <%=str_client_ename%>
                </td>
            </tr>
            <tr>
                <td class="t1">
                    出生日期
                </td>
                <td class="t2">
                    <%=str_client_birth%>
                </td>
                <td class="t1">
                    <!--身份证号-->
                </td>
                <td class="t2">
                    <%'str_client_idno%>
                </td>
            </tr>
            <tr>
                <td class="t1">
                    会员账号
                </td>
                <td class="t2">
                    <%=str_client_id_email%>
                </td>
                <td class="t1">
                    联络电话
                </td>
                <td class="t2">
                    <%=str_client_cphone%>
                </td>
            </tr>
            <tr>
                <td class="t1">
                    电子邮箱
                </td>
                <td class="t2">
                    <%=str_client_email%>
                </td>
                <td class="t1">
                    通讯地址
                </td>
                <td class="t2">
                    <%=str_client_caddr%>
                </td>
            </tr>
        </table>
    </div>
    <p style="margin: 20px 0 0 8px;">
        立合同书人：
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            上海麦奇教育信息咨询有限公司（受Tutor Group Limited的委托为其所在的麦奇教育集团会员提供咨询服务，为其在中国地区的委托经销商，简称本公司）；
        <% else %>
            TutorGroup麦奇教育集团（即VIPABC.com，为香港注册且运营网站，并于中国地区指定上海麦奇教育信息咨询有限公司为授权经销商，简称本公司），
        <% end if %>
        正式会员：<span class="bold color_3"><%=str_client_cname%></span>（即注册付费会员，以下简称会员），兹就会员加入
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                麦奇教育集团网络数位学习系统，并于会籍有效期限内，依据本合同书所订的条款接受网络课程的咨询学习服务，双方同意就相关合同内容约定如下：
            <% else %>
                本公司VIPABC网络数位学习系统，并于会籍有效期限内，依据本合同书所订的条款参与网络咨询课程及服务，双方同意就相关合同内容约定如下：
            <% end if %>
    </p>
    <p style="margin: 20px 0 0 8px; padding-left: 10px; border-left: 5px #f0ede8 solid;">
        <% '20120316 阿捨新增 無限卡產品判斷合約內容
        dim intCountNum : intCountNum = 1
        if ( true = bolUnlimited ) then 
        %><%'無限卡產品判斷 %>
        <%'============================== 無限卡產品 Start ==============================%>
        一、
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            本公司为会员提供成为麦奇教育集团会员以及行使会员资格的咨询服务。会员将以新入会资格成为麦奇教育集团的正式<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;会员，在
            <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷%>
                本合同
            <% else %>
                <span class="bold color_3"><%=str_valid_month%>个月</span>的
            <% end if %>
                咨询服务期间
            <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷%>
                得依照本合同条件获得在线英语学习<span class="bold color_3"><%=str_product_cname%></span>产品（以下简称产品），可不限<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;次数使用小班制及大会堂的在线咨询学习服务。
            <% else %>
                获得为期<span class="bold color_3"><%=str_valid_month%>个月</span>的在线英语学习无限卡产品（以下简称产品），可不限次数使用小班<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;制及大会堂的在线咨询学习服务。
            <% end if %>
            <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷%>
                <br />
            <% else %>
                咨询服务期间<span class="bold color_3"><%=replace(replace(str_sdate&str_edate," ",""),"。","")%></span>。<br />
            <% end if %>
        <% else %><% '新合約內容以上 for 無限卡產品判斷%>
            会员系以新入会资格成为VIPABC正式会员，选定购买服务为<span class="bold color_3"><%=str_product_cname%></span>产品，产品内容为
            <span class="bold color_3"><%=str_valid_month%>個月</span>
            的咨询服务期间，可不限<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;次数使用小班制及大会堂咨询服务，产品价格为人民币<span class="bold color_3"><%=str_total_money%></span>元整。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;咨询服务期间<span class="bold color_3"><%=replace(replace(str_sdate&str_edate," ",""),"。","")%></span>。<br />
        <% end if %><% '舊合約內容以上 for 無限卡產品判斷%>
        <% intCountNum = intCountNum + 1 %>

        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容 for 無限卡產品判斷%>
             <%=arrCountChinaWord(intCountNum)%>、咨询服务费为人民币<span class="bold color_3"><%=str_total_money%></span>元整。
             <br />
             <% intCountNum = intCountNum + 1 %>
        <% end if %><% '新合約內容以上 for 無限卡產品判斷%>

        <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷 %>
            <%=arrCountChinaWord(intCountNum)%>、会员得依照本合同条款使用在线英语学习产品及服务:<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)&nbsp;本产品及服务限定为会员个人使用，具有专属性，不得被继承或以任何方式转让或提供给任何第三人使用。会员同意<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本公司如合理怀疑会员违反本约定者，本公司有权暂时停止会员账户之使用，且若经确认违反本约定者，本公司有权<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;解除或终止本合约，会员并须同时赔偿本公司以咨询服务费用为基准计算的十倍惩罚性违约金。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(2)&nbsp;为确保会员个人使用之专属性，会员同意本公司于会员使用之网络平台上设置相关的保护措施，以确保会员账号之安<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;全性。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(3)&nbsp;会员须有持续使用本咨询产品及服务之行为。本合同签署日起算满五年后，会员若连续一年期间未使用任何咨询服务<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;者，则终身卡之使用期间将于满一年未使用之日自动终止。终止后，会员即不得再以任何方式使用本咨询服务及在线<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;英语学习产品。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(4)&nbsp;终身无限卡产品无法使用一对一咨询服务及临订咨询课程之服务，会员须于咨询课程前24小时间预定咨询课程。 <br />
            <% intCountNum = intCountNum + 1 %>
        <% end if %>

        <% '20120419 阿捨新增 贈送月時數 %>
        <% if ( true = bolProductMonthReward ) then %>
            <%=arrCountChinaWord(intCountNum)%>、本公司同意会员若于完成語言分析當天付清全部款项并点选接受「会员入会契约书」者，则本公司同意免费延长本产<br />
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;品<span class="bold color_3"><%=intProductMonthReward%></span>个月之咨询期间。会员了解并同意该免费延长之期间，不得转赠予任何其他第三人或要求折换价金。<br />
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;且会员若中途解约者，该赠送之期间即自动失效。<br />
            <% intCountNum = intCountNum + 1 %>
        <% end if %>

        <% '贈送New IPad %>
        <% if ( true = bolNewIPadGift ) then %>
            <%=arrCountChinaWord(intCountNum)%>、VIPABC同意会员若于完成语言分析当天付清全部款项并点选接受「会员入会契约书」者，VIPABC将于合约起始日起算30天<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;后，赠送会员赠品<span class="bold color_3">New iPad壹台</span>，但若会员于合约期间内解除合约，或有违反入会契约书之内容情节重大、经VIPABC取消<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;其会籍资格者，则会员应按赠品价格计人民币<span class="bold color_3">3,688</span>元整赔偿VIPABC，且VIPABC有权自应退还予会员之款项中直接扣抵。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;赠送之New iPad为<span class="bold color_3">APPLE New iPad 16G Wi-Fi </span>版本之平板计算机(VIPABC保留变更产品型号、规格，改以其他同等级产品<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;取代之权利)，送交产品之型号规格以实际出货之实物为准，且VIPABC不负安装、设定、保固、维修及产品瑕疵之责。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;VIPABC仅提供中国地区境内之寄送，且不负寄送时之迟延、遗失、错误、无法辨识或毁损之责任。<br />
            <% intCountNum = intCountNum + 1 %>
        <% end if %>

        <%=arrCountChinaWord(intCountNum)%>、会员了解并同意确实遵守下列约定:<br />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)&nbsp;会员应妥善保管会员账号及密码，且不得将会员账号密码提供给其他第三人使用，或以任何方式提供第三人使用会员的<br />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;咨询服务
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容 for 無限卡產品判斷%>
            包括产品的在线学习(以下简称课程)
        <% end if %><% '新合約內容以上 for 無限卡產品判斷%>
        。
        <% '20121018 阿捨新增 新增無限卡條例 %>
         <% if ( true = bolUnlimited AND false = bolLifeUnlimited ) then %><%'無限卡產品判斷 且 非終身無限卡%>
            会员同意<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本公司如合理怀疑会员违反本约定者，本公司有权暂时停止会员账户之使用，且若经确认违反本约定者，本公司有权<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;解除或终止本合约，会员并须同时赔偿本公司以咨询服务费用为基准计算的十倍惩罚性违约金。<br />
         <% else %><%'無限卡產品判斷 以上%>
            <br />
        <% end if %><%'非無限卡產品判斷 以上%>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(2)&nbsp;
        <% '20121018 阿捨新增 新增無限卡條例 %>
         <% if ( true = bolUnlimited AND false = bolLifeUnlimited ) then %><%'無限卡產品判斷 且 非終身無限卡%>
            为确保会员个人使用之专属性，会员同意本公司于会员使用之网络平台上设置相关的保护措施，以确保会员账号之安<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;全性。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(3)&nbsp;
         <% else %><%'無限卡產品判斷 以上%>
         <% end if %><%'非無限卡產品判斷 以上%>
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容 for 無限卡產品判斷%>
            会员承诺若无法出席已预约之咨询服务或课程，应于所预约之咨询或课程时间开始前4小时取消；若于咨询服务期间内<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;，发生一周内已预约却未出席且无事先取消之咨询服务或课程超过20堂时，则本公司有权取消会员无限卡资格，届时<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;会员剩余之咨询服务期间内，可使用之咨询服务堂数则以每周20堂为限。
        <% else %><% '新合約內容以上 for 無限卡產品判斷%>
            会员承诺已预约之咨询服务若无法出席时，应于所预约之咨询时间开始前4小时取消；若于咨询服务期间内，发生一周内
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;已预约却未出席且无事先取消之咨询服务或课程超过20堂时，则本公司有权取消会员无限卡资格，届时会员剩余之咨询服<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;务期间内，可使用之咨询服务堂数则以每周20堂为限。<br />
        <% end if %><% '舊合約內容以上 for 無限卡產品判斷%>
        <% intCountNum = intCountNum + 1 %>
        <br />
		
        <% '20121018 阿捨新增 無限卡 條例 %>
        <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷%>
			<%=arrCountChinaWord(intCountNum)%>、特别约定:<br />
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容 for 無限卡產品判斷%>
            <% else %>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本条约定为会员与VIPABC就本产品之特别约定，本条约定内容若与VIPABC会员服务条款内<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;容不相一致者，则以本条款之内容为准。<br />
            <% end if %>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)&nbsp;
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容 for 無限卡產品判斷%>
                产品无法使用一对一咨询服务及临时订课服务。
            <% else %><% '新合約內容以上 for 無限卡產品判斷%>
                本产品无法使用一对一咨询服务。
            <% end if %><% '舊合約內容以上 for 無限卡產品判斷%>
            <br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(2)&nbsp;由于会员购买产品之特殊性(不限堂数使用)，会员了解并同意本产品一经使用后即不得以任何方式或理由要求将会籍转<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;让、移转予任何其他第三人，且不适用
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容 for 無限卡產品判斷%>
                《会员服务使用条款》
            <% else %><% '新合約內容以上 for 無限卡產品判斷%>
                VIPABC会员服务条款
            <% end if %><% '舊合約內容以上 for 無限卡產品判斷%>
            第十五条第二项付费移转之规定。<br />
            <% intCountNum = intCountNum + 1 %>
            <%=arrCountChinaWord(intCountNum)%>、終止合約:<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)&nbsp;会员了解并同意签订本合约后若欲申请解除或终止合约及退费者，应依《会员服务使用条款》第十七条规定办理。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(2)&nbsp;除因会员违反本合约规定、法律、法规或因可归责于会员之事由而解除或终止者外，双方同意本公司可随时解除或终止<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;合约，然意本公司应将会员已经支付之费用全额无息退还予会员，且意本公司无须负赔偿或补偿之责。<br />
            <% intCountNum = intCountNum + 1 %>
        <% else %><% '無限卡判斷 以上%>
            <%'矛盾點 以下應該不會再出現了%>
            <%'矛盾點 以上應該不會再出現了%>
        <% end if %><% '非無限卡判斷 以上 %>

        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容 for 無限卡產品判斷%>
            <%=arrCountChinaWord(intCountNum)%>、其他:<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)&nbsp;本合同书所附的《会员服务使用条款》适用于本合同书，本合同书条款若与《会员服务使用条款》内容不相一致者，<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;则以本合同书之内容为准。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(2)&nbsp;本合同书受中国法律管辖。如果有争议，应提交至北京仲裁委员会按其适用的仲裁规则仲裁解决。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;仲裁裁决对双方均具有约束力。<br />
            <% intCountNum = intCountNum + 1 %>
        <% end if %><%'新合約內容以上 for 無限卡產品判斷%>
        </p>
        <%'============================== 無限卡產品 End ==============================%>
        <% else %><%'無限卡產品判斷以上%>
        <%'============================== 一般產品 Start ==============================%>
            <%
            '20120723 阿捨新增 套餐產品堂數 明細判斷
            intMaxSession = ""
            strClassTypeName = ""
            srtCurriculum = ""
            if ( true = bolComboProduct ) then
                dim arrClassType : arrClassType = array(7, 3, 99)
                dim arrClassTypeName : arrClassTypeName = array("1对1课程", "1对3课程", "大会堂课程")
                if ( isArray(arrClassType) ) then
                    intTotal = Ubound(arrClassType)
                    for intI = 0 to intTotal
                        intMaxSession = 0
                        intClassType = arrClassType(intI)
                        intMaxSession = getComboProductMaxSession(intProductSn, intClassType, 0, CONST_TUTORABC_R_CONN)
                        if (  intMaxSession > 0 ) then
                            strClassTypeName = arrClassTypeName(intI)
                        end if
                    next 
                end if ' if ( isArray(arrClassType) ) then
            end if 'if ( true = bolComboProduct ) then

            if ( not isEmptyOrNull(intMaxSession) ) then
                srtCurriculum = intMaxSession
            else
                srtCurriculum = str_count_osession
            end if
            %>
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
                一、本公司为会员提供成为麦奇教育集团会员以及行使会员资格的咨询服务。会员将以<% if ( true = bolRenewStatus ) then %>續約方式繼續<% else %>新入会资格<% end if %>成为麦奇教育集团的正式会员，
                    在<%=str_valid_month%>个月的咨询服务期间获得为期<%=str_valid_month%>个月的在线英语学习<span class="bold color_3"><%=str_product_cname%></span>产品（以下简称产品），
                    在每30天的固定周期内，会员可使用之在线咨询学习服务堂数为<span class="bold color_3"><%=srtCurriculum%></span>课时。
                    咨询服务期间为<span class="bold color_3"><%=replace(replace(str_sdate&str_edate," ",""),"。","")%></span>。另本公司同意赠送会员<span class="bold color_3"><%=str_count_fsession%></span>课程。<br />
                <% intCountNum = intCountNum + 1 %>     
                <%=arrCountChinaWord(intCountNum)%>、咨询服务费为人民币<%=str_total_money%>元整。<br />
                <% intCountNum = intCountNum + 1 %>     
                <%=arrCountChinaWord(intCountNum)%>、其他<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;（1）本合同书所附的《会员服务使用条款》适用于本合同书，本合同书条款若与《会员服务使用条款》内容不相一致者，则以本合同书之内容为准。<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;（2）本合同书受中国法律管辖。如果有争议，应提交至北京仲裁委员会按其适用的仲裁规则仲裁解决。仲裁裁决对双方均具有约束力。
                <% intCountNum = intCountNum + 1 %>
            <% else %><% '新合約內容以上 for 一般產品判斷%>
                一、会员系以<% if ( true = bolRenewStatus ) then %>續約方式繼續<% else %>新入会资格<% end if %>成为VIPABC正式会员，选定购买服务为<span class="bold color_3"><%=str_product_cname%></span>，<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;咨询服务期间<span class="bold color_3"><%=replace(replace(str_sdate&str_edate," ",""),"。","")%></span>，<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;总共可使用
                <% if ( not isEmptyOrNull(intMaxSession) ) then %>
                    <%=strClassTypeName%>
                <% end if %>
                <span class="bold color_3"><%=srtCurriculum%></span>课时，
                <% '如果有這些資訊在秀出來 %>
                <% if ( not isEmptyOrNull(str_count_fsession) ) then %>
                    另并赠送<span class="bold color_3"><%=str_count_fsession%></span>课时，
                <% end if %>
                <%'如果有這些資訊在秀出來 %>
                <% if ( not isEmptyOrNull(int_account_cycle_max_session) ) then %>
                    <span class="bold color_3">每<%=int_account_cycle_days%>天</span>最少需使用<span class="bold color_3"><%=int_account_cycle_max_session%></span>课时，
                <% end if %>
                总购买费用计人民币<span class="bold color_3"><%=str_total_money%></span>元。
            <% intCountNum = intCountNum + 1 %>
            <% end if %><% '舊合約內容以上 for 一般產品判斷%>

             <% if ( true = bolNewIPadGift ) then %><% '贈送New IPad %>
                <br /><%=arrCountChinaWord(intCountNum)%>、VIPABC同意会员若于完成语言分析当天付清全部款项并点选接受「会员入会契约书」者，VIPABC将于合约起始日起算30天<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;后，赠送会员赠品<span class="bold color_3">New iPad壹台</span>，但若会员于合约期间内解除合约，或有违反入会契约书之内容情节重大、经VIPABC取消<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;其会籍资格者，则会员应按赠品价格计人民币<span class="bold color_3">3,688</span>元整赔偿VIPABC，且VIPABC有权自应退还予会员之款项中直接扣抵。<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;赠送之New iPad为<span class="bold color_3">APPLE New iPad 16G Wi-Fi </span>版本之平板计算机(VIPABC保留变更产品型号、规格，改以其他同等级产品<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;取代之权利)，送交产品之型号规格以实际出货之实物为准，且VIPABC不负安装、设定、保固、维修及产品瑕疵之责。<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;VIPABC仅提供中国地区境内之寄送，且不负寄送时之迟延、遗失、错误、无法辨识或毁损之责任。<br />
                <% intCountNum = intCountNum + 1 %>
            <% end if %>

            <% if ( true = bolRenewStatus ) then %><%'續約判斷%>
                <br /><%=arrCountChinaWord(intCountNum)%>、会员了解本公司所提供学习课程方式为每周期排定固定之课时数，会员若未能于当周期使用完毕，未使用课时数将自动归零<br />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;，不会累计至次周期。
                <% intCountNum = intCountNum + 1 %>
            <% end if %>
        <% end if %><%'非無限卡產品判斷以上%>
        <%'============================== 一般產品 End ==============================%>

        <%'獎學金活動%>
        <% if ( true = bolScholarship ) then 
            dim strCfgScholarshipProduct : strCfgScholarshipProduct = CONST_VIPABC_SCHOLARSHIP_PRODUCT_SESSION_SETTING
            dim arrCfgScholarshipProduct : arrCfgScholarshipProduct = split(strCfgScholarshipProduct, ";")
            dim intTotalNum : intTotalNum = 0
            dim intSubTotalNum : intSubTotalNum = 0
            dim objScholarshipProductSession : objScholarshipProductSession = null
            dim objScholarshipProductMoney : objScholarshipProductMoney = null
            Set objScholarshipProductSession = Server.CreateObject("Scripting.Dictionary")
            Set objScholarshipProductMoney = Server.CreateObject("Scripting.Dictionary")

            if ( isArray(arrCfgScholarshipProduct) ) then
                intTotalNum = Ubound(arrCfgScholarshipProduct)
                for intI = 0 to intTotalNum
                    strScholarshipProductData = arrCfgScholarshipProduct(intI)
                    arrScholarshipProductData = split(strScholarshipProductData, "#")
                    intScholarshipProductSn = arrScholarshipProductData(0)
                    arrScholarshipData0 = split(arrScholarshipProductData(1), "-")
                    arrScholarshipData1 = split(arrScholarshipProductData(2), "-")
                    objScholarshipProductSession(intScholarshipProductSn & "_0") = arrScholarshipData0(0)
                    objScholarshipProductSession(intScholarshipProductSn & "_1") = arrScholarshipData1(0)
                    objScholarshipProductMoney(intScholarshipProductSn & "_" & arrScholarshipData0(0)) = arrScholarshipData0(1)
                    objScholarshipProductMoney(intScholarshipProductSn & "_" & arrScholarshipData1(0)) = arrScholarshipData1(1)
                next
                'strScholarshipProductSn = strScholarshipProductSn & "," 
            end if
        %>
            <br /><%=arrCountChinaWord(intCountNum)%>、本公司同意会员若于<%=arrNowMaxDate(0) &"年"& arrNowMaxDate(1) &"月"& arrNowMaxDate(2) &"日"%>前首次购买【<%=str_product_cname%>】方案之咨询课程，且于上开期限前付清全部款项<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;并点选接受「会员入会契约书」者，则可参加奖学金活动方案，活动内容及条件如下：<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1.&nbsp;参加本活动之会员于原合约期限期满(不含延长期限及增购期限者)，若实际使用之课程堂数达到指定堂数(「达标堂数」)<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;时，得向本公司申请奖学金，相关约定如下：<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)&nbsp;「达标堂数」之计算方式如下：<br /> 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(a)&nbsp;可计入「达标堂数」之课堂型态：大会堂、一对多及一对一课程。惟一对多及大会堂课程使用堂数须为「达标堂<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;数」的50%以上，若不足50%时，则请领的奖学金数额以一半计算。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(b)&nbsp;不计入「达标堂数」之课堂型态：观看录像文件，使用Express Services加扣堂数，或使用后因故归还堂数者，皆<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;不计入「达标堂数」。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(c)&nbsp;限会员预约并且会员本人实际出席课程者始能计入「达标堂数」，预约但未出席上课之堂数，则不予计入。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(d)&nbsp;仅限于本合约付费购买之堂数数额内计算(不含赠送、受让、整并及已转让给他人之堂数等。)。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(2)&nbsp;奖学金数额如下表所示：<br />
            <table width="400" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td align="center">
                    <table width="300" border="0" cellpadding="2" cellspacing="1" bgcolor="#999999">
                        <tr>
                        <td bgcolor="#FFFFFF" align="center">达标堂数</td>
                        <td bgcolor="#FFFFFF" align="center">得申领之奖学金(人民币)</td>
                      </tr>
                      <% 
                      for intJ = 0 to 1
                        intScholarshipProductSession = objScholarshipProductSession(intProductSn & "_" & intJ)
                        intScholarshipProductMoney = objScholarshipProductMoney(intProductSn & "_" & intScholarshipProductSession)
                        if ( true = bolDeBugMode ) then
                            response.Write "intProductSn & _ & intJ: " & intProductSn & "_" & intJ & "<br/>"
                            response.Write "intScholarshipProductSession: " & intScholarshipProductSession & "<br/>"
                            response.Write "intProductSn & _ & intScholarshipProductSession: " & intProductSn & "_" & intScholarshipProductSession & "<br/>"
                            response.Write "intScholarshipProductMoney: " & intScholarshipProductMoney & "<br/>"
                        end if 
                      %>
                      <tr>
                        <td bgcolor="#FFFFFF" align="center"><%=intScholarshipProductSession%> 堂</td>
                        <td bgcolor="#FFFFFF" align="center"><%=intScholarshipProductMoney%> 元</td>
                      </tr>
                       <% next %>
                    </table>
                    </td>
                </tr>
            </table>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2.&nbsp;奖学金之申请以一次为限，且于合约期限届满后始可申请。符合条件之会员应主动向本公司提出申请，并缴交相关文件(申<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;请书及身分证复印件)予本公司。本公司审核通过后，将于计算认定完毕后一个月内支付奖学金予会员。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3.&nbsp;发送之奖学金将依中华人民共和国税法之相关规定进行申报与扣缴。<br />
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4.&nbsp;会员于契约期限内若有提前终止契约或转让契约、课程之情事者，会员或其受让人均不适用本活动。<br />
        <% intCountNum = intCountNum + 1 %>
        <% end if %><%'獎學金活動%>
        </p>
    <div align="right" style="margin-top: 20px;">
        <table width="74%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td width="38%" align="left">
                <% if ( true = bolNewContract ) then %>
                    麦奇教育集团
                <% else %>
                    TutorGroup麦奇教育集团
                <% end if %>
                </td>
                <td width="62%" rowspan="4" align="left">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="43%" align="right">
                                会员签名：
                            </td>
                            <td width="57%" align="left">
                                <span class="bold">
                                    <%=str_client_cname%></span>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                法定代理／监护人：
                            </td>
                            <td align="left">
                                <span class="bold">
                                    <%=str_sulegal%></span>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                付款人：
                            </td>
                            <td align="left">
                                <span class="bold">
                                    <%=str_payman%></span>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                销售顾问：
                            </td>
                            <td align="left">
                                <span class="bold">
                                    <%=str_sale_name%></span>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="left">
                    <% if ( true = bolNewContract ) then %>
                        中国区委托经销商
                    <% else %>
                        中国区授权经销商
                    <% end if %>
                </td>
            </tr>
            <tr>
                <td align="left">
                    上海麦奇教育信息咨询有限公司
                </td>
            </tr>
            <tr>
                <td align="left">&nbsp;
                </td>
            </tr>
        </table>
    </div>
    <br />
    <div class="dode">
    </div>
    <div class="t-center">
        <div class="title2 big t-center">
            <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            会员服务条款
            <% else %>
            VIPABC会员服务条款
            <% end if %>
            </div>
            <% if ( true = bolNewContract ) then %><%'20120927 Tree GM12092503 更新VIPABC的合約內容 要20120928當天候漸單的才要做修改%>
                <img src="/images/images_cn/contract_01_new2.gif"><br />
                <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷 %>
                    <img src="/images/images_cn/contract_02_life_limited.gif"><br />
                <% else %><%'終身無限卡產品判斷 以上%>
                    <img src="/images/images_cn/contract_02_new2.gif"><br />
               <% end if %><%'非終身無限卡產品判斷 以上%>
                <% if ( true = bolRenewStatus ) then %><%'續約判斷%>
                    <img src="/images/images_cn/contract_03_new3.gif"><br />
                <% else %>
                    <img src="/images/images_cn/contract_03_new2.gif"><br />
                <% end if %>
                <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷 %>
                    <img src="/images/images_cn/contract_04_life_limited.gif"><br />
                <% else %><%'終身無限卡產品判斷 以上%>
                    <img src="/images/images_cn/contract_04_new2.gif"><br />
               <% end if %><%'非終身無限卡產品判斷 以上%>
                <img src="/images/images_cn/contract_05_new2.gif"><br />
                <img src="/images/images_cn/contract_06_new2.gif"><br />
                <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                    <img src="/images/images_cn/contract_07_new2.gif"><br /><%'新產品無限卡顯示%>
                <% else %>
                    <img src="/images/images_cn/contract_07_new3.gif"><br /><%'新產品定時定額顯示%>
                <% end if %>
                <img src="/images/images_cn/contract_08_new2.gif"><br />
                <img src="/images/images_cn/contract_09_new2.gif"><br />
            <% else %>
                <img src="/images/images_cn/contract_01.gif"><br />
                <img src="/images/images_cn/contract_02.gif"><br />
                <% if ( true = bolRenewStatus ) then %><%'續約判斷%>
                    <img src="/images/images_cn/contract_03_renew.gif"><br />
                    <img src="/images/images_cn/contract_04_renew.gif"><br />
                <% else %>
                    <img src="/images/images_cn/contract_03.gif"><br />
                    <img src="/images/images_cn/contract_04.gif"><br />
                <% end if %>
                <img src="/images/images_cn/contract_05.gif"><br />
                <img src="/images/images_cn/contract_06.gif"><br />
                <% if ( true = bolNewProduct ) then %><%'新舊產品判斷 %>
                    <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                        <img src="/images/images_cn/contract_07_3.gif"><br /><%'新產品無限卡顯示%>
                    <% else %>
                        <img src="/images/images_cn/contract_07_2.gif"><br /><%'新產品定時定額顯示%>
                    <% end if %>
                <% else %>
                    <img src="/images/images_cn/contract_07.gif"><br /><%'新產品顯示%>
                <% end if %>
                <% if ( true = bolNewProduct ) then %><%'新舊產品判斷 %>
                    <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                        <img src="/images/images_cn/contract_08_3.gif"><br /><%'新產品無限卡顯示%>
                    <% else %>
                        <img src="/images/images_cn/contract_08_2.gif"><br /><%'新產品定時定額顯示%>
                    <% end if %>
                <% else %>
                    <img src="/images/images_cn/contract_08.gif"><br /><%'新產品顯示%>
                <% end if %>
                <img src="/images/images_cn/contract_09.gif"><br />
            <% end if %>
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            <img src="/images/images_cn/contract_10_new.gif"><br />
            <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷 %>
                <img src="/images/images_cn/contract_11_life_limited.gif"><br />
            <% else %><%'終身無限卡產品判斷 以上%>
                <img src="/images/images_cn/contract_11_new.gif"><br />
            <% end if %><%'非終身無限卡產品判斷 以上%>
        <% else %>
            <img src="/images/images_cn/contract_10.gif"><br />
            <img src="/images/images_cn/contract_11.gif"><br />
        <% end if %>
        <img src="/images/images_cn/contract_12.gif"><br />
        <%'20120607 阿捨新增 合約建單日判斷相關條款%> 
        <% if ( datediff("d", "2012/06/07", strInputDate) < 0 ) then %> 
            <img src="/images/images_cn/contract_13.gif"><br />
            <img src="/images/images_cn/contract_14.gif"><br />
        <% else %><%'20120607 阿捨新增 合約建單日判斷相關條款 以上%> 
            <% if ( true = bolRenewStatus ) then %><%'續約判斷%>
                <img src="/images/images_cn/contract_13_renew.gif"><br />
            <% else %><%'續約判斷 以上%>
                <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷 %>
                    <img src="/images/images_cn/contract_13_life_limited.gif"><br />
                    <img src="/images/images_cn/contract_14_life_limited.gif"><br />
                <% else %><%'終身無限卡產品判斷 以上%>
                    <img src="/images/images_cn/contract_13_new.gif"><br />
                    <img src="/images/images_cn/contract_14_new.gif"><br />
                <% end if %><%'非終身無限卡產品判斷 以上%>
            <% end if %><%'非續約判斷 以上%>      
        <% end if %><%'非合約建單日判斷相關條款 以上%> 
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            <img src="/images/images_cn/contract_15_new.gif"><br />
            <img src="/images/images_cn/contract_16_new.gif"><br />
        <% else %>
            <img src="/images/images_cn/contract_15.gif"><br />
            <img src="/images/images_cn/contract_16.gif"><br />
        <% end if %>
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            <% if ( true = bolUnlimited ) then %><%'無限卡產品判斷 %>
                <img src="/images/images_cn/contract_17_new2.gif"><br />
            <% else %>
                <img src="/images/images_cn/contract_17_new3.gif"><br />
            <% end if %>
        <% else %>
            <img src="/images/images_cn/contract_17.gif"><br />
        <% end if %>
        <% if ( true = bolLifeUnlimited ) then %><%'終身無限卡產品判斷 %>
            <img src="/images/images_cn/contract_18_life_limited.gif"><br />
        <% else %><%'終身無限卡產品判斷 以上%>
            <img src="/images/images_cn/contract_18.gif"><br />
        <% end if %><%'非終身無限卡產品判斷 以上%>
        <img src="/images/images_cn/contract_19.gif"><br />
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            <img src="/images/images_cn/contract_20_new.gif"><br />    
        <% else %>
            <img src="/images/images_cn/contract_20.gif"><br />
        <% end if %>
        <% if ( true = bolNewContract ) then %><%'更新VIPABC的合約內容%>
            <img src="/images/images_cn/contract_21_new2.gif">
        <% else %>
            <img src="/images/images_cn/contract_21.gif">
        <% end if %>
    </div>
    <div class="dode">
    </div>
    <p class="t-center">
        <input type="button" value="+ 打印合约" id="print_contract" class="btn_1 m-left10" />
        <%
    'if ( str_client_cname <> "" and str_client_birth <> ""  and str_client_idno <> ""  and str_client_email <> ""  and str_client_caddr <> ""  and str_client_cphone <> "" and str_client_cphone <> "-" and str_online_credit <> "Y" and bol_is_all_paid and (not bol_is_cagree) ) then 
    if ( str_client_cname <> "" and str_client_birth <> ""  and str_client_email <> ""  and str_client_caddr <> ""  and str_client_cphone <> "" and str_client_cphone <> "-" and (not bol_is_cagree) ) then     
        %>
        <input type="button" name="agree_or_yet" id="agree" value="+ <%=getWord("CONTRACT_116")%>"
            class="btn_1 m-left10" onclick="return checkdata(<%=Session("contract_sn")%>);" />
        <input type="button" name="agree_or_yet" id="disagree" value="+ <%=getWord("CONTRACT_117")%>"
            class="btn_1 m-left10" />
        <input type="button" name="print_or_prepage" style="display: none" id="print_contract"
            value="+ <%=getWord("CONTRACT_118")%>" class="btn_1 m-left10" onclick="print();" />
        <input type="hidden" name="contract_sn" id="contract_sn" value="<%=int_contract_sn%>" />
        <%
	'BD09030903 電子合約增加提示訊息功能(若基本資料不完全) raychien 20090310
	elseif (str_client_cname = "" or str_client_birth = "" or str_client_idno = "" or str_client_email = "" or str_client_caddr = "" or str_client_cphone = "" or str_client_cphone = "-" ) then
		 response.write "<script type=""text/javascript"">alert('"&getWord("CONTRACT_119")&"');</script>"
		 response.end
	end if 
        %>
    </p>
    </form>
</div>
<!--內容end-->
</div>

<!-- Main Content End -->