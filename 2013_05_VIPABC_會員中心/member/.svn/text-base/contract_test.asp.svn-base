<% response.write "<script>alert('Page does not exist');</script>"
response.write "<script>location.href='/index.asp';</script>" 
%>
<!--#include virtual="/lib/include/global.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>VIPABC</title>
<link media="screen" type="text/css" href="../css/css_cn.css" rel="stylesheet"/>
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script type="text/javascript">
//======check 頁面上的checkbox是否全部勾選 ======
function checkdata(contract){
	//return false;
	for (var i=0;i<document.forms[0].elements.length;i++) {
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
</script>
</head>

<body>
<%

Dim int_contract_sn : int_contract_sn = Session("contract_sn")             '接Session-->合約編號

if ( int_contract_sn = "" ) then int_contract_sn = Request("contract_sn")	'無奈啊,我要嘆三聲了!!
'if ( int_contract_sn = "" ) then int_contract_sn = Request("contract_sn")

'response.write "contract_sn......."&int_contract_sn&"<br>"
'response.write "client_sn......."&request("client_sn")&"<br>"
Dim str_tmp_go_page

'塞資料的相關設定值
Dim var_arr 																	'傳excuteSqlStatementRead的陣列
Dim arr_result																'接回來的陣列值
Dim str_tablecol 																'table欄位
Dim str_tablerow 																'table列數

'---------- 重新找電子合約 ------- start -----
Dim str_contract_sql : str_contract_sql = ""

if ( isEmptyOrNull(int_contract_sn) ) then 

	str_contract_sql="SELECT sn FROM client_temporal_contract WHERE (client_sn = @client_sn) AND (valid=1) AND (ctype = 1)"
	var_arr = Array(g_var_client_sn)
	arr_result = excuteSqlStatementRead(str_contract_sql,var_arr,CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then
			int_contract_sn = arr_result(0, 0)
		else
			'Response.write "no data!!"
		end if
	end if
	set str_contract_sql = nothing
	set var_arr = nothing
	set arr_result = nothing
	
	'response.write g_str_sql_statement_for_debug&"<br>"
end if 
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
'===================================================
'算出這筆合約有沒有折讓Begin

Dim discount_sql	
discount_sql="SELECT ISNULL(SUM(amount), 0) AS discount FROM client_payment_detail WHERE (payrecord_sn IN (SELECT sn FROM client_payment_record WHERE (contract_sn = @contract_sn))) AND (payment_mode = 14  or payment_mode = 28)"

	var_arr = Array(int_contract_sn)
	arr_result = excuteSqlStatementRead(discount_sql,var_arr,CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
		'Response.write("欄:" & Ubound(arr_result) & "<br>")
		'Response.write("列:" & Ubound(arr_result,2) & "<br>")	
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
Dim contract_sql	
contract_sql= " SELECT DERIVEDTBL_2.contract_sn, "
contract_sql= contract_sql & " client_temporal_contract.ptype, "
contract_sql= contract_sql & " client_temporal_contract.lead_sn, "
contract_sql= contract_sql & " lead_basic.oaddr, "
contract_sql= contract_sql & " client_temporal_contract.client_sn, "
contract_sql= contract_sql & " cfg_product.cname AS product, "
contract_sql= contract_sql & " cfg_product.product_name AS product_nickname, "
contract_sql= contract_sql & " lead_basic.cname, lead_basic.fname + ' ' + lead_basic.lname AS ename, "
contract_sql= contract_sql & " lead_basic.client_address, "
contract_sql= contract_sql & " lead_basic.htel_code + '-' + lead_basic.htel + (CASE isnull(lead_basic.htel_branch, '') WHEN '' THEN NULL ELSE ' #' + lead_basic.htel_branch END) AS hphone, "
contract_sql= contract_sql & " CONVERT(varchar, client_temporal_contract.inpdate, 111) AS cdate, "
contract_sql= contract_sql & " CAST(YEAR(client_temporal_contract.inpdate) AS varchar(4)) + ' 年 ' + CAST(MONTH(client_temporal_contract.inpdate) AS varchar(4)) + ' 月 ' + CAST(DAY(client_temporal_contract.inpdate) AS varchar(4)) + ' 日' AS sudate, "
contract_sql= contract_sql & " CAST(YEAR(client_temporal_contract.sdate) AS varchar(4)) + ' 年 ' + CAST(MONTH(client_temporal_contract.sdate) AS varchar(4)) + ' 月 ' + CAST(DAY(client_temporal_contract.sdate) AS varchar(4)) + ' 日' AS sdate, "
contract_sql= contract_sql & " CAST(YEAR(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + ' 年 ' + CAST(MONTH(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + ' 月 ' + CAST(DAY(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + ' 日' AS edate, "
contract_sql= contract_sql & " cfg_product.total_points / 65 AS osession, "
contract_sql= contract_sql & " lead_basic.sex, "
contract_sql= contract_sql & " lead_basic.company, lead_basic.email, "
contract_sql= contract_sql & " lead_basic.client_address AS Expr1, "
contract_sql= contract_sql & " lead_basic.cemail, lead_basic.caddr, "
contract_sql= contract_sql & " lead_basic.legal_name, lead_basic.legal_idno, "
contract_sql= contract_sql & " lead_basic.legal_relations,lead_basic.legal_tel, "
contract_sql= contract_sql & " staff_basic.fname + ' ' + staff_basic.lname AS agent, "
contract_sql= contract_sql & " (CASE lead_basic.sex WHEN 1 THEN '男' ELSE '女' END) AS sex,lead_basic.idno, CAST(YEAR(lead_basic.birth) AS varchar(4)) + ' 年 ' + CAST(MONTH(lead_basic.birth) AS varchar(4)) + ' 月 ' + CAST(DAY(lead_basic.birth) AS varchar(4)) + ' 日' AS birth, "
contract_sql= contract_sql & " lead_basic.cphone_code + '-' + lead_basic.cphone AS cphone, "
contract_sql= contract_sql & " lead_basic.contact_way, lead_basic.company AS Expr2, "
contract_sql= contract_sql & " lead_basic.otel_code + '-' + lead_basic.otel + (CASE isnull(lead_basic.otel_branch, '') WHEN '' THEN NULL ELSE ' #' + lead_basic.otel_branch END) AS ophone, "
contract_sql= contract_sql & " cfg_product.total_cash, cfg_product.pays_off,cfg_product1.first_cash, "
contract_sql= contract_sql & " DERIVEDTBL_2.paytype, DERIVEDTBL_2.payman, lead_basic.ccode, "
contract_sql= contract_sql & " cfg_post.Data1, cfg_city.Data1 AS Expr3, "
contract_sql= contract_sql & " sb.fname + ' ' + sb.lname AS sname,cfg_product.point_type, "
contract_sql= contract_sql & " cfg_session_points.Point_Fee, DERIVEDTBL_1.sum_gift AS fsession, "
contract_sql= contract_sql & " client_temporal_contract.prodbuy , "
contract_sql= contract_sql & " cfg_product.valid_month, client_temporal_contract.discount_type, "
contract_sql= contract_sql & " cfg_product.is_extending_contract, "
contract_sql= contract_sql & " client_freemonth.freegift AS freemon,"
contract_sql= contract_sql & " CAST(YEAR(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + '/' + CAST(MONTH(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4)) + '/' + CAST(DAY(DATEADD(d, - 1, DATEADD(m, cfg_product.valid_month, client_temporal_contract.sdate))) AS varchar(4))  AS date_edate , "
'定時定額 (暫不考慮一周幾天的設定) pipi 20100303
contract_sql= contract_sql & " cfg_product.account_cycle_max_session, " '一周上幾堂
contract_sql= contract_sql & " cfg_product.week_count "	'共幾周

contract_sql= contract_sql & " FROM cfg_city "
contract_sql= contract_sql & " INNER JOIN cfg_post "
contract_sql= contract_sql & " INNER JOIN lead_basic "
contract_sql= contract_sql & " ON cfg_post.Data2 = lead_basic.ccode "
contract_sql= contract_sql & " ON cfg_city.sn = cfg_post.city_sn "
contract_sql= contract_sql & " RIGHT OUTER JOIN staff_basic "
contract_sql= contract_sql & " RIGHT OUTER JOIN client_temporal_contract "
contract_sql= contract_sql & " LEFT OUTER JOIN client_freemonth  "
contract_sql= contract_sql & " ON client_temporal_contract.sn = client_freemonth.contract_sn "
contract_sql= contract_sql & " ON staff_basic.sn = client_temporal_contract.inputer "
contract_sql= contract_sql & " ON lead_basic.sn = client_temporal_contract.lead_sn "
contract_sql= contract_sql & " LEFT OUTER JOIN cfg_session_points "
contract_sql= contract_sql & " INNER JOIN cfg_product "
contract_sql= contract_sql & " ON cfg_session_points.Sn = cfg_product.point_type "
contract_sql= contract_sql & " ON client_temporal_contract.prodbuy = cfg_product.sn "
contract_sql= contract_sql & " RIGHT OUTER JOIN "
contract_sql= contract_sql & " ( "
contract_sql= contract_sql & " 		SELECT "
contract_sql= contract_sql & " 		contract_sn, "
contract_sql= contract_sql & " 		SUM(freegift) AS sum_gift "
contract_sql= contract_sql & " 		FROM client_freegift "
contract_sql= contract_sql & " 		WHERE (contract_sn = @contract_sn_1) "
contract_sql= contract_sql & " 		GROUP BY contract_sn "
contract_sql= contract_sql & " ) AS DERIVEDTBL_1 "
contract_sql= contract_sql & " RIGHT OUTER JOIN "
contract_sql= contract_sql & " ( "
contract_sql= contract_sql & " 		SELECT contract_sn, paytype, payman, handling "
contract_sql= contract_sql & " 		FROM client_payment_record WHERE (contract_sn = @contract_sn_2) "
contract_sql= contract_sql & " ) AS DERIVEDTBL_2 "
contract_sql= contract_sql & " ON DERIVEDTBL_1.contract_sn = DERIVEDTBL_2.contract_sn "
contract_sql= contract_sql & " ON client_temporal_contract.sn = DERIVEDTBL_2.contract_sn "
contract_sql= contract_sql & " LEFT OUTER JOIN staff_basic AS sb "
contract_sql= contract_sql & " ON DERIVEDTBL_2.handling = sb.sn  "
contract_sql= contract_sql & " left outer JOIN cfg_product as cfg_product1 "
contract_sql= contract_sql & " ON client_temporal_contract.old_prodbuy = cfg_product1.sn "

if (str_online_credit <> "Y") then
	contract_sql=contract_sql&"WHERE (client_temporal_contract.valid = 1)"
end if

	var_arr = Array(int_contract_sn,int_contract_sn)
	arr_result = excuteSqlStatementRead(contract_sql,var_arr,CONST_VIPABC_RW_CONN)
	
'response.write "int_contract_sn........."&int_contract_sn&"<br>"
'response.write g_str_sql_statement_for_debug&"<br>"
'response.end
if (isSelectQuerySuccess(arr_result)) then
    '有資料
    if (Ubound(arr_result) > 0) then    
		str_sdate = arr_result(13,0)                '合約開始日期
		str_edate = arr_result(14,0)                '合約結束日期
		str_product_cname = arr_result(5,0)             '購買產品中文名稱
		str_client_cname = arr_result(7,0)  '購買人中文名字
		str_client_ename = arr_result(8,0)  '購買人英文名字
		str_client_sex = arr_result(27,0)	  '購買人性別
		str_client_birth =  arr_result(29,0)      '購買人生日
		str_client_idno =  arr_result(28,0)         '購買人身份証字號
		str_client_id_email =  arr_result(18,0)  '購買人帳號email
		str_client_email	=  arr_result(20,0)  		'購買人通訊email
		str_client_cphone = arr_result(30,0)		'購買人手機號碼
		str_client_caddr = arr_result(21,0)		'購買人地址
		
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
		
		'response.write "<br>"&str_freemon&"............."&int_point_fee&"............."&str_discount_type
		
		'------- 贈送堂數計算 ------- start -------
		if ( not isEmptyOrNull(int_fsession) ) then 
		
			if ( isnumeric(int_fsession) and int_fsession > 0 ) then
				if ( int_point_type = 1 ) then
					str_count_fsession =((int_fsession / 65) / 4)
				elseif ( int_point_type = 7 ) then
					str_count_fsession =((int_fsession / 65) / 3)
				else
					str_count_fsession =( int_fsession / 65 )
				end if
				'freegift="<span lang=""EN-US"" style=""font-size: 12.0pt; font-family: 新細明體"">本公司並同意會員另有免費贈送 <span lang=""en-us"">"&fs&"</span> 堂之優惠。</span>"	
			else
				'freegift="&nbsp;"
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
		if ( isEmptyOrNull(str_discount_type) ) then str_discount_type = "" end if 
		if ( str_ptype = "1" and str_discount_type <> "S") then 
			int_total_money = arr_result(35,0) 'pays_off
		else
			int_total_money = arr_result(34,0)  'total_cash
		end if 
			
		'把總價扣掉折讓就是應出現的產品金額
		int_total_money = int_total_money - int_discount
		if ( isnumeric(int_total_money) ) then str_total_money = convertNumber(int_total_money,CONST_CN_STD_NUMBER )&"元" end if
		'------ 金額 ------ end ---------
		
		'----- 服務期間 ----- start -------
		
		
		str_sdate = "自 "&str_sdate&" 起，"	
		str_edate = "至 "&str_edate&" 止。"
		
		
		if(not isEmptyOrNull(str_freemon)) then
				str_tmp_edate = dateadd("m",str_freemon,str_tmp_edate)
				str_edate=year(str_tmp_edate)&" 年 "&month(str_tmp_edate)&" 月 "&day(str_tmp_edate)&" 日"	
			else
			
				str_edate = str_edate
			end if	
		'----- 服務期間 ----- end -------
		
		if ( str_client_idno <> "" ) then str_client_idno = "******"&right(str_client_idno,4) end if 
		'付款人
		if ( isEmptyOrNull(Request("payman"))  or str_payman = "" )  then str_payman = str_client_cname else str_payman = str_payman end if 
		'法定代理人
		if ( not isEmptyOrNull(str_lname) <> "" )  then str_sulegal = str_lname else str_sulegal="&nbsp;"
		
		else
		'無資料
		
	end if 
end if 

%>


<!-- Menber Content Start -->
       <!--內容start-->
       <div class="terms_main2">
       <div class="title2 t-center big">会员权益概要</div>
    
    <div>
    <div class="check"><input type="checkbox" name="checkbox" id="checkbox" /></div>
    <div class="law">本会员了解所购买之产品为VIPABC <span class="bold color_3">xxxxx</span> 产品，每周排定可使用 <span class="bold color_3">xxxxx</span> 课时，本服务期间自 <span class="bold color_3">xx</span> 年 <span class="bold color_3">xx</span> 月 <span class="bold color_3">xx</span> 日起，至 <span class="bold color_3">xx</span> 年 <span class="bold color_3">xx</span> 月 <span class="bold color_3">xx</span> 日止，总共可使用 <span class="bold color_3">xxx</span> 课时，总购买费用计人民币（下同） <span class="bold color_3">xxxxx</span> 元。 </div>
    <div class="clear"></div>
    </div>
    
    <div>
    <div class="check"><input type="checkbox" name="checkbox" id="checkbox1" value='y' <%if request("Checkbox1")="y" then %>checked<%end if %>/></div>
    <div class="law">本会员了解并同意：<br>
    1. 所购买的产品内容为1~3人精致小班制教学服务，每次咨询将扣抵一课时。但一对一教学服务，计费是以二次小班制课程兑换一次一对一课程（即扣抵二课时）。<br>
    2. 本公司并有提供单向教学为主之英语学习大会堂服务(每次使用扣抵一堂课程)。<br>注：大会堂是以专业课程为主（例如商业谈判、商业简报、发音、语法课程等等）的单向教学式之30人以上大讲堂。<br>
    3. 每次上课时间为45分钟（小班制、一对一、大会堂课程皆是）。本公司所提供学习课程方式为每周排定固定之课时数，会员若未能于当周使用完毕，未使用课时数将自动归零，不会累计至次周。惟经本公司同意，会员得预支使用其预先排定之课时数。<br>
    4. 本公司提供多国籍语言顾问进行语言咨询服务，但会员无法指定特定顾问进行咨询服务。 
    </div>
    <div class="clear"></div>
    </div>
    
    <div>
    <div class="check"><input type="checkbox" name="checkbox" id="checkbox" /></div>
    <div class="law">本会员了解并同意会员使用权限账号须全数缴清所购买产品款项后始开通。</div>
    <div class="clear"></div>
    </div>
    
    <div>
    <div class="check"><input type="checkbox" name="checkbox" id="checkbox" /></div>
    <div class="law">本会员了解并同意最少需在课程服务开始前八小时通过网络进行课程预约，如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程。若未于课前四小时取消课程，则该课程仍计入已使用课时数。</div>
    <div class="clear"></div>
    </div>
    
    <div>
    <div class="check"><input type="checkbox" name="checkbox" id="checkbox" /></div>
    <div class="law">本会员了解并同意于购买产品后如于七日内办理退费，除扣除行政管理费500元，并应扣除已使用之课程费用；第八日起至第三十日止办理退费，除扣除行政管理费3,000元，并应扣除此期间之课程费用。购买产品逾一个月者，则不得退费。</div>
    <div class="clear"></div>
    </div>
    
    <div>
    <div class="check"><input type="checkbox" name="checkbox" id="checkbox" /></div>
    <div class="law">本会员了解并同意会员不得将会籍出售、转让给予他人使用，非经本公司同意私下出售、转让会造成丧失会员资格且不得请求退还未使用课程之费用；本会员确因情况特殊有会籍移转之必要，应依照合同内载明之服务条款向本公司办理。</div>
    <div class="clear"></div>
    </div>
    
    <div>
    <div class="check"><input type="checkbox" name="checkbox" id="checkbox" /></div>
    <div class="law">本会员了解并同意会员购买之产品，是依会员个人需求将咨询课时数与时间搭配，并已预留相当时间供弹性调整，会员购买前应详细评估判断，选择最合适自己方案及期限。因此视为会员同意咨询之服务期，不得延期。</div>
    <div class="clear"></div>
    </div>
    
    
    <div class="dode"></div>
    <br>
    <div class="title2 big t-center">VIPABC会员合同书</div>
    
    <div>
    <div class="name1">中文姓名：黃曉明</div>
    <div class="name2">英文姓名：Happy</div>
    <div class="name3">■ 女</div>
    <div class="clear"></div>
    <div class="name1">出生日期：1979年3月5日</div>
    <div class="name4">身份证号：T*****8502</div>
    <div class="clear"></div>
    <div class="name1">会员账号：happyman@yahoo.com.tw</div>
    <div class="name4">联络电话：0921-051415</div>
    <div class="clear"></div>
    <div class="name1">电子邮箱：happyman54652@yahoo.com.tw</div>
    <div class="name4">通讯地址：xxxxxxxxx</div>
    <div class="clear"></div>
    </div>
    <p>立合同书人：TutorGroup麦奇教育集团（即VIPABC.com，为香港注册且运营网站，并于中国地区指定上海麦奇教育信息咨询有限公司为授权经销商，简称本公司），正式会员：（即注册付费会员，以下简称会员），兹就会员加入本公司VIPABC网络数位学习系统，并于会籍有效期限内，依据本合同书所订的条款参与网络咨询课程及服务，双方同意就相关合同内容约定如下：</p>
    <p>会员系以新入会资格成为VIPABC正式会员，选定购买服务为<span class="bold color_3">xxxxx</span>，每周排定可使用<span class="bold color_3">xx</span>课时，咨询服务期间自 <span class="bold color_3">xx</span> 年 <span class="bold color_3">xx</span> 月 <span class="bold color_3">xx</span> 日起，至 <span class="bold color_3">xx</span> 年 <span class="bold color_3">xx</span> 月 <span class="bold color_3">xx</span> 日止，总共可使用 <span class="bold color_3">xxx</span> 课时，总购买费用计人民币（下同） <span class="bold color_3">xxxxx</span> 元。</p>
    <br>
    <p><span class="bold">第一条 接受条款</span><br>
      本公司以「动态课程生成系统」，为会员提供语言咨询及相关专业英语服务，<span style="text-decoration:underline">会员对本公司提供之网络数位学习咨询服务及方式，已于本公司网站上充分阅读了解并同意遵守本约之规范，并可随时阅览，就其内容亦有相当时间详尽审阅后，在自由意愿下签订本合同</span>。当会员按下「同意」键或正式签署后，即表示会员在使用本公司所提供的任何服务时，愿意遵守本合同书之条款、会员服务规范及相关规定。若会员无法接受本约内容，请离开或无须有任何签署或入会之行为。</p>
    <p><span class="bold">第二条 监护权之行使</span><br>本公司会员之年龄，如系未满十八岁之会员时，除已了解本约说明外，并应于会员之法定代理人（或监护人）阅读、了解并同意本合同书所有内容后，才可以使用VIPABC网站提供相关之服务。当会员使用或继续使用VIPABC，即表示会员之法定代理人（或监护人），已阅读、了解并同意接受合同书所有内容。</p>
    <p><span class="bold">第三条 会员注册义务</span><br>
会员依注册程序加入新会员时，应详实登录个人正确、完整之资料（包括真实姓名及各项相关资料），并随时更新已过时或异动之资料。如果会员提供的个人资料不实，或未及时更新资料而有任何延误或损失时，应由会员自行负责，本公司亦可随时终止会员的会员资格及使用各项服务之权利。</p>

<p><span class="bold">第四条 资格及有效期限</span><br>一、会员以新入会资格成为VIPABC正式会员，所选定购买之课程及合同咨询服务期限，详见「VIPABC会员入会基本资料」所载。本会员了解并同意会员使用权限账号须全数缴清后，所购买产品款项后始开通。<br>
二、会员已了解并同意，本公司于开放咨询时段内，会员可随时享有咨询服务。<span style="text-decoration:underline">本公司亦得因应会员人数增加而随之增聘语言顾问，以满足会员使用需求，所以会员于合同有效期内，无论课时数使用之情形如何，合同时间与课时数等同重要。合同期满或虽未期满或终止但课时数于期限内使用完毕时，本公司对于会员之服务即视为终止，会员不得再有所主张或请求</span>。<br>
三、会员于合同有效期内观摩其它会员之录像文件，观看大会堂录像，均以一次课计算。 <br>
</p>
<p><span class="bold">第五条 使用义务与责任</span><br>会员使用本公司网站服务前，须自行配备上网所需之各项计算机软硬件设备，并负担连接因特网及电话等相关费用。使用网站服务应遵守香港、中国及会员联机所在地相关法令及因特网之国际使用惯例与礼节，并不得侵害任何第三人之权益。<br>
</p>
<p><span class="bold">第六条 系统安装及提供</span><br>本公司于合同有效期限内，应提供会员VIPABC在线系统之使用安装及该系统维护服务，使会员得以正常使用网络数位学习系统。如系会员个人计算机之软硬件或电信业者因特网系统等障碍因素，则不在本公司提供维护服务范围之内。会员亦不得以网络带宽、其它环境或会员个人因素，致使其学习质量不佳时，向本公司要求或主张任何之权益。<br>
</p>
<p><span class="bold">第七条 系统暂停或中断</span><br>本公司有例行性或因故需将本网站相关系统设备进行迁移、更换或维护，而须暂停或中断全部之服务时，本公司将于72小时前在系统内或网站上，以张贴告示之方式通知会员。本公司如未依约定方式通知致影响会员之权益时，会员可请求将合同服务日数延长二日。但本公司遇有临时或紧急或不可抗力之事由，或于事后并未关闭服务主机者，不在此限。<br>
</p>
<p><span class="bold">第八条 咨询系统维护服务</span><br>本公司对VIPABC相关之系统应负维护之责，会员如遇系统无法正常使用或有异动之情形，应迅速与客服人员联系。本公司提供服务项目包括下列；
<br />一、对于网站咨询服务系统之故障排除及维护更新。
<br />二、会员对于咨询服务系统之使用建议或申诉。
<br />三、会员个人资料申请、变更之协助。
<br />四、其它与会员权益有关并依约定提供服务之事宜。
<br>
</p>
<p><span class="bold">第九条 会员基本规则</span><br>一、会员同意并遵守关于使用咨询服务之规则，包括现行实施或相关之公告。会员同意本公司得就<span style="text-decoration:underline">本合同内容及相关信息单方为不定期修订，并于公告后正式生效，敬请随时查询VIPABC网站公告。若会员于本合同内容或相关信息修订后仍继续使用</span>咨询服务，则视为其同意该等内容及信息之修订。<br />
二、<span style="text-decoration:underline">会员必须于课程开始前八小时在线预约，以确认咨询上课时间，如课程当天因故无法出席，最迟须在四个小时前来电或通过网络取消课程。若未于课前四小时取消课程，则该课程仍计入已使用课时数。</span><br />
三、为尊重自身及其它会员权益，预约后请务必准时进入在线咨询室。如有逾时之情形，会员将无法进入该次课程，该预约课时数仍会计入使用课时数。<br />
四、<span style="text-decoration:underline">本公司所提供学习课程方式为每周排定固定之课时数，会员若未能于当周使用完毕，未使用课时数将自动归零，不会累计至次周，若由此造成会员课程未能如期完成由会员自行负责。</span>惟会员得经本公司同意预支使用其预先排定之课时数。<br />
五、网络预约：www.vipabc.com　客服信箱：<%=const_mail_address_services_vipabc%><br />
本公司中国区授权经销商客户服务专线：（+86）4006-15-66-15 
<br>
</p>
<p><span class="bold">第十条 禁止行为</span><br>下列行为系本公司VIPABC网站及在线咨询服务禁止之行为，但不以这些行为为限；<span style="text-decoration:underline">会员如有不当行为或违反相关服务条款时，本公司视违反情节轻重或影响大小，而予以暂停或撤销其会籍，会员除不得要求退还未使用课时数之费用外，本公司并保留该损害赔偿之请求权利。</span><br>
一、会员入侵或破坏本公司网站内部各项设施，或任何有损及本公司名誉之行为。<br>
二、利用在线视频系统施行：暴露、猥亵、性暗示或其它不雅或足以影响本公司咨询服务进行之行为。<br>
三、利用在线系统进行言语或文字骚扰、谩骂、攻击，或其它不受欢迎之行为。<br>
四、通过文件传输软件，发送非法软件或具破坏或影响他人权益之情形。<br>
五、在咨询进行中或服务系统上，有销售营利、政治目的、宗教宣扬或其它与课程内容无关之行为。<br>
六、为维护双方之权益，本公司禁止职员及语言顾问与会员在非咨询服务时间或利用其它以外之方式互动（例如；实时通讯等软硬件设施）。如在此规范外之一切行为或意外发生，概与本公司无涉。<br>
七、利用公开竞标、拍卖或以其它方式租、借、转卖本公司会籍资格及服务项目。<br>
八、私自提供账号、密码予他人使用本公司服务项目。<br>
九、会员咨询之课程内容与教材相关权利属于本公司所有，会员不得擅将咨询课程之录音、录像或教材内容，公开播放或销售。<br>
十、其它经由本公司经营者认定之不当行为。
<br></p>
<p><span class="bold">第十一条 咨询费用之缴付</span><br><span style="text-decoration:underline">会员应视个人实际之咨询所需与经济能力，选择适当产品类型及课时数</span>。会员于签订本合同时，应以现金或转账汇款或信用卡刷卡为主要支付方式，并将款项壹次付清予本公司或由授权经销商代收。会员选择现金支付、汇款或由金融机构办理分期方式支付时，请先与客户服务专员联系，以便提供相关信息资料。本公司于会员付清款项及本合同签署完成后，始为新会员正式启用其咨询资格。<br></p>
<p><span class="bold">第十二条 融资方式缴付</span><br>会员咨询费用款项若以金融机构之融资方式办理支付时，除可自行与金融机构办理外，亦可选择向本公司往来金融机构申请办理，由金融机构代会员壹次付清费用，会员则以分期方式向金融机构缴纳期付金。但有关会员之审核权限及金融机构与会员间权利义务关系，仍由该双方自行约定之，与本公司无涉，本公司亦不为任何之担保。<br>
会员使用前述本公司往来之金融机构申贷本约应付款项时，该笔金融机构贷款所收取之手续相关费用部分，本会员同意予以负担。
</p>
<p><span class="bold">第十三条 解除合同</span><br>会员自合同使用生效日起七日内，如对本公司所提供服务不愿再继续使用时，应以书面方式通知本公司或授权经销商解除合同，本公司就所缴付金额扣除行政管理费500元，并就使用部分，按每一单次课程300元计算费用后，退回剩余金额（但一对一课程以600元计，观看会员个人咨询录像文件或其它会员之录像文件课程时，每课时以300元计）。<br>
会员自合同使用生效日第八日起至第三十日止，如不愿再继续使用时，应以书面方式通知本公司或授权经销商解除合同，本公司就所缴付金额扣除行政管理费3,000元，并扣除此期间之课程费用后，退回剩余金额。<br>
自合同生效日三十一日后，会员则不得通知解除合同或要求退费。
</p>
<p><span class="bold">第十四条 会籍之专属性</span><br>会员了解并同意，本公司提供语言咨询服务之理念，系使用全世界独创并荣获多国发明专利之「动态课程生成系统」，不仅按每位会员语言能力、个人背景、兴趣及学习偏好与学习比重，将会员特性、需求做适合之配对，咨询后再辅以学习追踪系统，以此做为多位语言顾问评估会员升级之依据；该具体且完整追踪学习效果，并为个人专属的管理学习资源而不可替代。<span style="text-decoration:underline">因此，本公司对每位会员的前置投资作业及专属性，并非可按时间及课时数量化其价值。因而会员加入顾问服务后，不能以个人因素任意解除、终止合同，亦不得将本服务提供予他人使用，若有解约或终止会籍事宜，应依本约第十八条规定办理之。</span></p>
<p><span class="bold">第十五条 会籍之转让规范</span><br>
一、 会员了解并同意，会员在行使咨询服务时，系以本公司研发之数位学习系统，采取计算机比对方式，将会员职业、兴趣及背景等输入，配比出合适之语言顾问提供指导，咨询过程中并由语言顾问对会员学习过程记录分析，以完整追踪学习成效。<span style="text-decoration:underline">因此个人会籍不得转卖，会员亦不得将会籍资格、可使用课时数以出售、公开竞标或拍卖或出借等方式予他人使用，因此会影响本公司操作系统及其它会员权益，将视为无效行为</span>。如经发现后，本公司可终止合同并停止提供会员服务，会员不得请求费用之退还或补偿，其与受让人间之权利义务关系概与本公司无涉。 <br>
二、会员确因情形特殊有移转会籍之必要时，<span style="text-decoration:underline">仅限转让予三亲等内亲属及本公司会员，并应提出书面申请，本公司有权审核受让会员之资格，经本公司同意受让会员之资格，于扣除（取消）原所赠课时数及相关优惠后，会员应另支付1,000元之会籍移转手续费。</span><br>会员同意会籍之移转，须全数课时数移转且以壹次为限，且该受让之5新会员，不得有终止合同或再次移转之情形。 
</P>
<p><span class="bold">第十六条 会员会籍不得延期</span><br>
本公司提供服务之设计，是依会员个人需求将咨询课时数与时间搭配，并已预留相当时间供弹性调整，会员聘任前应详细评估判断，选择最合适自己方案及期限。<span style="text-decoration:underline">因此会员同意咨询之服务期间，不得延期</span>。
</p>
<p><span class="bold">第十七条 违约处理方式</span><br>
一、<span style="text-decoration:underline">会员有本约第十一条各款情形之一者不予退费，本公司并得要求赔偿损失。</span><br> 
二、<span style="text-decoration:underline">会员若有按本约第十三条向金融机构融资缴付之情形，而致本公司有损失金融机构之相关手续费用时，本公司将要求赔偿损失。</span><br>
三、<span style="text-decoration:underline">本公司若有赠送会员课程课时数，须以课程课时数使用完毕且合同未到期时方始生效</span>，若会员有提前解除或终止合同，或有本约第十一条各款情形之一者，不得请求给付赠与课时数或请求以之折换现金。
</p>
<p><span class="bold">第十八条 会员续约权利</span><br>
会员购买之产品及课时数使用完毕后，可采续购方式继续行使会员权益。该相关续约手续及办法，依本公司规定办理。</p>
<p><span class="bold">第十九条 影像文件之使用</span><br>
会员同意其于咨询期间相关之录像、录音及文字、图片文件等内容，可无偿提供本公司系列产品或服务之编辑与研发、会员教学观摩、业务推广、产品发表或销售之使用。</p>
<p><span class="bold">第二十条 合同效力</span><br>
双方同意本约内容之全部或部分判定无效时，并未影响合同其它条款之效力。<span style="text-decoration:underline">为免生争议及认定上之歧异，任何口头非文字载明之承诺、条件，均视为无效约定。</span></p>
<p><span class="bold">第廿一条 未尽事项</span><br>
就本约未尽事宜，双方同意依本公司内部或网站上有关之公告、说明、服务条款补充，并依法律与其它相关法令解释并适用之。</p>
<p><span class="bold">第廿二条 准据法</span><br>
本公司与会员同意本合同之解释，依据中华人民共和国法律之规定。</p>
<p><span class="bold">第廿三条 争议解决</span><br>
双方同意因本合同发生争议时，应本着诚信原则善尽最大努力以协商方式处理。协商不成，双方均得于本公司授权经销商所在地之人民法院提起诉讼。</p>
<p><span class="bold">第廿四条 合同之保存</span><br>
本「入会合同书」及「VIPABC会员入会基本资料」于会员确认并同意后，请会员自行将该源文件保留并打印纸本，兹为双方合同签署之凭据。</p>
<div class="dode"></div>
<div>
    <div class="name1">TutorGroup麦奇教育集团中国区授权经销商<br>
    上海麦奇教育信息咨询有限公司<br>销售顾问：xxxxxxxxx<br>銷售主管：xxxxxxxxx<br>财务人员确认：<br>日期 xx / xx / xx<br>
    </div>
    <div class="name4"><br><br>会员签名：xxx<br>
      法定代理 / 监护人：xxx<br>
      付款人：xxx<br>签约日期 xx / xx / xx</div>    
    <div class="clear"></div>       
    </div>
<div class="dode"></div>
<p class="t-right"><input type="submit" name="button" id="button" value="同意" class="btn_1"/><input type="submit" name="button" id="button" value="不同意" class="btn_1 m-left10"/><input type="submit" name="button" id="button" value="打印合约书" class="btn_1 m-left10"/><input type="submit" name="button" id="button" value="回上页" class="btn_1 m-left10"/></p>

    </div>
       <!--內容end-->
    <!--end-->
</div>
<div class="clear"></div>
</div>
<!-- Main Content End -->
