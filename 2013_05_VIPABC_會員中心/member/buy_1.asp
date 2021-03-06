<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type='text/javascript'>
var buy_1_2 = "<%=getWord("buy_1_2")%>";
</script>
<script type='text/javascript' src="/lib/javascript/framepage/js/buy_1.js"></script>
<%
'定義變數；

Dim str_client_fullname : str_client_fullname = session("client_fullname")   '接下session值-->client_fullname

Dim bol_isbuy : bol_isbuy = false                                           '判斷是否有購買紀錄,預設false
Dim bol_payment_record_history : bol_payment_record_history = false 		'歷史付款紀錄,判斷是否曾經有付款紀錄,預設false
Dim str_paytype : str_paytype = ""                                          '繳款類別
Dim str_paymode : str_paymode = ""                                          '付款方式
Dim str_paydate : str_paydate = ""                                          '付款日期
Dim int_damount : int_damount = 0                                           '付款金額

'購買明細資料
Dim str_purchase_sql : str_purchase_sql = ""
Dim str_total_cash :  str_total_cash = 0                                   'total_cash
Dim str_amount :  str_amount = 0                                           '产品总额
Dim str_contract_sn : str_contract_sn = ""                                  '合約編號(6位數)
Dim str_product_cname : str_product_cname = ""                              '产品名称
Dim str_kdate : str_kdate = ""                                              '订购日期
Dim str_porderno : str_porderno = ""                                        '订单编号
Dim str_sales : str_sales = ""                                              '专员
Dim str_total_points : str_total_points = ""                                '使用课时数
Dim str_sdate :  str_sdate = ""                                             '课程开始日期
Dim str_edate :  str_edate = ""                                             '课程结束日期			
Dim str_valid_month :  str_valid_month = ""                                 '服务期限			
Dim str_aext : str_aext = ""                                                '分机號碼
Dim i : i = 0
Dim arr_sales '专员 陣列

'付款明細資料開始
Dim str_payment_detail_sql : str_payment_detail_sql = ""
Dim str_payment_detail_tablerow												'購買明細筆數
Dim str_dsn 

'歷史付款明細紀錄結束
Dim str_pay_amount : str_pay_amount = 0                                     '已付金額
Dim str_payment_his_sql: str_payment_his_sql = ""
Dim bol_payment_record_history_tablerow										'購買明細筆數
Dim str_invoice : str_invoice = ""                                          '发票种类
Dim str_invoiceno : str_invoiceno = ""                                      '发票号码

Dim str_new_invoice : str_new_invoice = 0
Dim str_new_invoiceno : str_new_invoiceno = ""
											
Dim str_new_invoice_sql
Dim var_invoice_arr
Dim arr_invoice_result
Dim str_invoice_tablerow
															

'塞資料的相關設定值
Dim var_arr 																	'傳excuteSqlStatementRead的陣列
Dim arr_result																'接回來的陣列值
Dim str_tablecol 																'table欄位
Dim str_tablerow 																'table列數

'20120316 阿捨新增 無限卡產品判斷	
dim bolUnlimited : bolUnlimited = false
dim strUnlimited : strUnlimited = ""
'20121025 阿捨新增 終身無限卡產品判斷	
dim bolLifeUnlimited : bolLifeUnlimited = false
dim strLifeUnlimited : strLifeUnlimited = ""
str_contract_sql = " SELECT prodbuy FROM client_temporal_contract WHERE (client_sn = @client_sn) AND (valid=1) AND (ctype = 1) order by sn DESC "
var_arr = Array(g_var_client_sn)
arr_result = excuteSqlStatementRead(str_contract_sql, var_arr, CONST_VIPABC_RW_CONN)
'response.Write g_str_sql_statement_for_debug&"<br>"
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
		intProductSn = arr_result(0, 0)
	
        if ( Not isEmptyOrNull(intProductSn) ) then
            if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (intProductSn & ",")) > 0 ) then
                strUnlimited = "unlimited"
                bolUnlimited = true
            end if

            if ( Instr(CONST_NO_LIMIT_FOREVER_PRODUCT_SN, (intProductSn & ",")) > 0 ) then
                strLifeUnlimited = "終身"
                bolLifeUnlimited = true
            end if
        end if
	else
		'response.Write "no data!!"
	end if
end if
set str_contract_sql = nothing
set var_arr = nothing
set arr_result = nothing

'20120606 VM12053006 學習儲值卡官網+IMS功能 Penny
Dim bolLearningCard : bolLearningCard = false
if ( Not isEmptyOrNull(intProductSn) ) then
	if ( Instr(CONST_VIPABC_LEARNING_CARD_PRODUCT_SN, ("," & intProductSn & ",")) > 0 ) then
		bolLearningCard = true
	end if
end if
Dim intPeriod : intPeriod = 0
if ( true = bolLearningCard ) then
	str_contract_sql = " SELECT period FROM RandomKey WITH (NOLOCK) WHERE (ClientSn = @client_sn)"
	var_arr = Array(g_var_client_sn)
	arr_result = excuteSqlStatementRead(str_contract_sql, var_arr, CONST_VIPABC_RW_CONN)
	'response.Write g_str_sql_statement_for_debug&"<br>"
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then
			intPeriod = arr_result(0, 0)
		else
			'response.Write "no data!!"
		end if
	end if
end if
        
'================================================================
'購買明細資料開始
str_purchase_sql="SELECT  TOP (1) contract_pay.sn, contract_pay.cname, CONVERT(varchar, contract_pay.kdate, 111) AS kdate, "
str_purchase_sql = str_purchase_sql & " contract_pay.porderno, contract_pay.amount1, contract_pay.sales, "
str_purchase_sql = str_purchase_sql & " CASE isnull(client_purchase.add_points, 0) WHEN 0 THEN (  "
str_purchase_sql = str_purchase_sql & " CASE total_points2 WHEN 0 THEN ( "
str_purchase_sql = str_purchase_sql & " 	 CASE isnull(client_freegift.freegift, 0) WHEN 0 THEN contract_pay.total_points "
str_purchase_sql = str_purchase_sql & " 	 ELSE (contract_pay.total_points + client_freegift.freegift) END"
str_purchase_sql = str_purchase_sql & "  	 ) ELSE total_points2 END "
str_purchase_sql = str_purchase_sql & "	) ELSE client_purchase.add_points + client_purchase.access_points END AS total_points,"
str_purchase_sql = str_purchase_sql & "  CONVERT(varchar, contract_pay.sdate, 111) AS sdate, "
str_purchase_sql = str_purchase_sql & " CASE isnull(contract_pay.product_edate, '') WHEN '' THEN CONVERT(varchar, contract_pay.edate, 111)  "
str_purchase_sql = str_purchase_sql & " ELSE CONVERT(varchar, contract_pay.product_edate, 111)  END AS edate,"
str_purchase_sql = str_purchase_sql & " contract_pay.valid_month, contract_pay.total_cash, contract_pay.aext, "
str_purchase_sql = str_purchase_sql & " client_purchase.add_points,isnull(client_purchase.pay_money,'') as pay_money  "
str_purchase_sql = str_purchase_sql & " FROM  contract_pay  "
str_purchase_sql = str_purchase_sql & " LEFT OUTER JOIN client_purchase ON contract_pay.sn = client_purchase.contract_sn "
str_purchase_sql = str_purchase_sql & " LEFT OUTER JOIN  ( "
str_purchase_sql = str_purchase_sql & "		select contract_sn, sum(freegift) as freegift from client_freegift group by contract_sn "
str_purchase_sql = str_purchase_sql & " ) as client_freegift ON contract_pay.sn = client_freegift.contract_sn "
str_purchase_sql = str_purchase_sql & " WHERE  (contract_pay.client_sn = @client_sn_1) AND (contract_pay.valid = 1) "
str_purchase_sql = str_purchase_sql & " OR (contract_pay.client_sn = @client_sn_2) AND (contract_pay.valid = 0) "
str_purchase_sql = str_purchase_sql & " AND (contract_pay.paymode = '客戶線上刷卡') AND (contract_pay.isbuy IS NULL OR contract_pay.isbuy = 0 ) "
str_purchase_sql = str_purchase_sql & " ORDER BY contract_pay.sn DESC, contract_pay.dsn "

var_arr = Array(g_var_client_sn,g_var_client_sn)

arr_result = excuteSqlStatementRead(str_purchase_sql,var_arr,CONST_VIPABC_RW_CONN)
'response.Write("錯誤訊息" & g_str_sql_statement_for_debug & "<br />")
if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then
		'response.Write("欄:" & Ubound(arr_result) & "<br />")
        'response.Write("列:" & Ubound(arr_result,2) & "<br />")
		for str_tablerow = 0 To Ubound(arr_result, 2)
		
			str_contract_sn = arr_result(0, str_tablerow)
			str_product_cname = arr_result(1,str_tablerow)
			str_kdate = arr_result(2,str_tablerow)
			str_porderno = arr_result(3,str_tablerow)
			str_amount = arr_result(4,str_tablerow)
			str_sales = arr_result(5,str_tablerow)
			str_total_points = arr_result(6,str_tablerow)
			str_sdate = arr_result(7,str_tablerow)
			str_edate = arr_result(8,str_tablerow)
			str_valid_month = arr_result(9,str_tablerow)
			str_total_cash = arr_result(10,str_tablerow)
			str_aext = arr_result(11,str_tablerow)
			'判斷是否為大陸客服，暫時解，主要還是要從資料庫下手
			'如果是則分機要用本來的號碼不能加8
			'如果分機不為空
			if str_aext<>"" then
				'如果分機長度為四，且第一個字為8（大陸分機）
				if Len(str_aext) = 4 and left(str_aext,1) = 8 then
					str_aext = right(str_aext,Len(str_aext)-1)
				end if
			end if 	
    	next
	else
		Call alertGo(getWord("buy_1_1"), "", CONST_ALERT_THEN_GO_BACK )
		response.end 
	end if		
else
	Call alertGo(getWord("buy_1_1"), "", CONST_ALERT_THEN_GO_BACK )
	response.end 
end if
set str_purchase_sql = nothing
set var_arr = nothing
set arr_result = nothing
'購買明細資料結束
'================================================================

%>



<div id="temp_contant">
	<!--內容start-->
    <div class='page_title_5'><h2 class='page_title_h2'>产品购买</h2></div>
    <div class="box_shadow" style="background-position:-80px 0px;"></div> 
    <div class="main_membox">
    <div class="buy">
        <div class="w"><%=getWord("buy_1_3")%><span class="bold"><%=str_client_fullname%></span><%=getWord("buy_1_4")%></div>
        <!--訂單 start-->
        <div class="box">
        	<div class="line"></div>
        	<div class="title"><%=getWord("BUY_1_5")%></div>
            
            <!--订购日期 date start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_6")%></div>
                <div class="show"><%=str_kdate%></div>
                <div class="clear"></div>
            </div>
            <!--订购日期 date end-->
            
            <!--订单编号 no start-->
            <%if (not isEmptyOrNull(str_porderno)) then %>
            <div class="con">
                <div class="name"><%=getWord("buy_1_7")%></div>
                <div class="show"><%=str_porderno%></div>
                <div class="clear"></div>
            </div>
            <%end if%>        
            <!--订单编号 no end-->
            
            <!--消费总额 total start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_8")%></div>
                <div class="show"><%=getWord("buy_1_9")%><%=replace("123456789","123456789","<b>"&formatnumber(str_amount,0)&"</b>")%><%=getWord("buy_1_10")%></div>
                <div class="clear"></div>
            </div>
            <!--消费总额 total end-->
            
            <!--服务专线tel start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_11")%></div>
                <%
                'if str_aext<>"" then 
                '	str_sales = replace(replace("02-23679811 分機 123／專員Jessic为您服务","123",str_aext),"Jessica",str_sales)
                'else 
                '	str_sales = replace(replace("02-23679811 分机 123／专员Jessic为您服务","分机 123",""),"Jessica",str_sales)
                'end if	
				
				if ( not isEmptyOrNull(str_aext)) then 
					str_aext = getWord("EXT") & str_aext
				end if 
				
                %>
                <div class="show"><%=getWord("buy_1_12")%><%=str_aext%>／<%=getWord("buy_1_13")%><%=str_sales%><%=getWord("buy_1_14")%></div>
                <div class="clear"></div>
            </div>
            <!--服务专线 tel end-->
            
            <!--服务信箱 mail start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_15")%></div>
                <%
                
                if lcase(left(str_sales,4))="eric" then
                    arr_sales = split(str_sales," ")
                    str_sales = arr_sales(1)&" "&arr_sales(0)
                end if 
                %>
                <div class="show"><%=const_mail_address_services_vipabc%></div>
                <div class="clear"></div>
            </div>
            <!--服务信箱 mail end-->
        </div>
        <!--訂單 end-->
       
        <!--商品start-->
        <div class="box">
            <div class="line"></div>
            <div class="title"><%=getWord("buy_1_16")%></div>
            
            <!--产品名称 name start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_17")%></div>
                <div class="show"><%=str_product_cname%></div>
                <div class="clear"></div>
            </div>
            <!--产品名称 name end-->
        
            <!--使用课时数 堂數 start-->
            <div class="con">
            <div class="name"><%=getWord("buy_1_18")%></div>
                <div class="show">
                <% 
                '20120316 阿捨新增 無限卡產品判斷合約內容
                if ( true = bolUnlimited ) then 
					response.Write strUnlimited
				'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493
				elseif ( 493 = sCInt(intProductSn) ) then 
					response.Write "付费大会堂2堂，一对三真人授课3堂"
                else 
				    'response.Write str_total_points
                    if ( str_total_points <> "" and sCLng(str_total_points) > 0 ) then 
				        response.Write sCLng(str_total_points) / 65 
				    else 
				        response.Write 0
				    end if 
                end if 
                %>
				</div>
            <div class="clear"></div>
            </div>
            <!--使用课时数 堂數 end-->
            
            <!--服务期限 deline start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_19")%></div>
                <div class="show">
					<%
						if ( true = bolLearningCard ) then
							response.Write intPeriod & " days"
						elseif ( 493 = sCInt(intProductSn) ) then
							response.Write "7 days"
                        elseif ( true = bolLifeUnlimited ) then 
					        response.Write strLifeUnlimited
						else
							response.Write str_valid_month & " months"
						end if					
					%>
				</div>
                <div class="clear"></div>
            </div>
            <!--服务期限 deline end-->
            
            <!--课程开始日期 start start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_20")%></div>
                <div class="show"><%=str_sdate%></div>
                <div class="clear"></div>
            </div>
            <!--课程开始日期 start end-->
            
            <!--课程结束日期 stop start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_21")%></div>
                <div class="show">
                <%
                '20121025 阿捨新增 終身無限卡產品判斷合約內容
                if ( true = bolLifeUnlimited ) then 
					response.Write strLifeUnlimited
                else 
				    response.Write str_edate
                end if 
                %>
                </div>
                <div class="clear"></div>
            </div>
            <!--课程结束日期 stop end-->
            
            <!--产品总额 start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_22")%></div>
                <div class="show"><%=formatnumber(str_total_cash,0)%></div>
                <div class="clear"></div>
            </div>
            <!--产品总额 end-->
            
            <!--总计（人民币） start-->
            <div class="con">
                <div class="name"><%=getWord("buy_1_23")%></div>
                <div class="show"><%=formatnumber(str_total_cash,0)%></div>
                <div class="clear"></div>
            </div>
            <!--总计（人民币） end-->
        </div>
        <!--商品end-->
       
        <Form name="buy" method="post" action="/program/member/buy_2.asp">
        <input type="hidden" name="dsn" value='' />
	    <input type="hidden" name="mh_usersn" value='<%=g_var_client_sn%>'/>
		<input type="hidden" name="str_product_name" value='<%=str_product_cname%>'/>
		<input type="hidden" name="str_sales" value='<%=str_sales%>'/>
        <%
		'================================================================
		'付款明細資料開始
		str_payment_detail_sql="select  paytype , paymode, damount, convert(varchar,paydate,111) as paydate,dsn  "
		str_payment_detail_sql = str_payment_detail_sql & " from contract_pay "
		str_payment_detail_sql = str_payment_detail_sql & " where client_sn=@client_sn and sn =@contract_sn and paymode in ('客戶線上刷卡','線上刷卡') and ( isbuy is null or isbuy = 0) "
		str_payment_detail_sql = str_payment_detail_sql & " order by dsn "
		
		'var_arr = Array("149532","814800")
		var_arr = Array(g_var_client_sn,str_contract_sn)

		arr_result = excuteSqlStatementRead(str_payment_detail_sql,var_arr,CONST_VIPABC_RW_CONN)
		'response.Write g_str_sql_statement_for_debug
		
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then
				'response.Write("欄:" & Ubound(arr_result) & "<br />")
				'response.Write("列:" & Ubound(arr_result,2) & "<br />")
				str_payment_detail_tablerow = Ubound(arr_result,2)            '付款資訊的列數(可能會有多筆)
				bol_isbuy = true
		%> 
        <!--說明start(本次付款项目说明)-->
        <div class="box">
            <div class="line"></div>
            <div class="title"><%=getWord("buy_1_24")%></div>
        	<div class="con2">
            <!--项目 title-->
            <div class="name1">
            	<div class="pt1"><%=getWord("buy_1_25")%></div>
                <div class="pt"><%=getWord("buy_1_26")%></div>
                <div class="pt"><%=getWord("buy_1_27")%></div>
                <div class="pt"><%=getWord("buy_1_28")%></div>
                <div class="pt"><%=getWord("buy_1_29")%></div>
                <div class="clear"></div>
            </div>
            <!--项目 title-->
			
            <!--缴清全额 show-->
            <div class="show1">
            <%
			for str_tablerow = 0 To Ubound(arr_result, 2)			
				str_dsn = ""                                                 '付款流水編號
				str_paytype = arr_result(0, str_tablerow)
				str_paymode = arr_result(1,str_tablerow)
				int_damount = arr_result(2,str_tablerow)
				str_paydate = arr_result(3,str_tablerow)
				str_dsn = arr_result(4,str_tablerow)
			%>
            	<div class="pt1"><input type="radio" name="check" id="rad_paydamount<%=str_dsn%>" onClick="chgdata('<%=str_dsn%>')" /></div>
                <div class="pt"><%=getWord("buy_1_30")%></div>
                <div class="pt"><%=getWord("buy_1_31")%></div>
                <div class="pt"><%=str_paydate%></div>
                <div class="pt"><%=formatnumber(int_damount,0) %></div>
                <div class="clear"></div>
                 <input type="hidden" name="damount_<%=str_dsn%>" value='<%=int_damount%>'/>
            <% next %>
                <div class="word"><%=getWord("buy_1_32")%><p><%=getWord("buy_1_33")%>
                    <input type="text" name="payamount" value='' class="login_box3 width13 m-right5 m-left5" disabled/>
                    <%=replace("100元整","100","")%></p>
                </div>
            </div>
            <!--缴清全额 show-->
            
        	</div>
        </div>
        <!--說明end-->
        <%
			end if
		end if
		set str_payment_detail_sql = nothing
		set var_arr = nothing
		set arr_result = nothing
		'付款明細資料結束
		'================================================================
		%> 
 
		<%
		'================================================================
		'歷史付款明細紀錄結束
		str_payment_his_sql="select distinct contract_pay.dsn, contract_pay.paytype , contract_pay.paymode, "
		str_payment_his_sql = str_payment_his_sql & " case isnull(client_money_back_record1.mb_type,'') when '' then contract_pay.damount  "
		str_payment_his_sql = str_payment_his_sql & " else contract_pay.damount*(-1) end as damount , "
		str_payment_his_sql = str_payment_his_sql & " convert(varchar,contract_pay.dadate,111) as paydate,contract_pay.invoice,contract_pay.invoiceno  "
		str_payment_his_sql = str_payment_his_sql & " from contract_pay "
		str_payment_his_sql = str_payment_his_sql & " left outer join (  "
		str_payment_his_sql = str_payment_his_sql & " 	select mb_sn,sn from client_payment_detail  "
		str_payment_his_sql = str_payment_his_sql & " ) as client_payment_detail1 "
		str_payment_his_sql = str_payment_his_sql & " on contract_pay.dsn = client_payment_detail1.sn "
		str_payment_his_sql = str_payment_his_sql & " left outer join ( "
		str_payment_his_sql = str_payment_his_sql & " 	select sn, mb_type from client_money_back_record where mb_type='換付款方式' or mb_type='多付退回' "
		str_payment_his_sql = str_payment_his_sql & " ) as client_money_back_record1 "
		str_payment_his_sql = str_payment_his_sql & " on client_money_back_record1.sn = client_payment_detail1.mb_sn "
		str_payment_his_sql = str_payment_his_sql & " where contract_pay.client_sn=@client_sn "
		str_payment_his_sql = str_payment_his_sql & " and contract_pay.sn =@contract_sn and ((contract_pay.paymode='線上刷卡' and contract_pay.isbuy=1 ) "
		str_payment_his_sql = str_payment_his_sql & " or ( contract_pay.paymode<>'線上刷卡' and contract_pay.paytype<>'退費' and contract_pay.dauditor <>'' and (contract_pay.padate is not null "
		str_payment_his_sql = str_payment_his_sql & " or contract_pay.dadate is not null or (datediff(y,contract_pay.paydate,getdate())>0) )) "
		str_payment_his_sql = str_payment_his_sql & " or contract_pay.paymode='折讓' or contract_pay.paymode='現金折扣'"
		str_payment_his_sql = str_payment_his_sql & " or (contract_pay.paytype='退費' and client_money_back_record1.sn is not null) ) "
		str_payment_his_sql = str_payment_his_sql & " order by contract_pay.dsn "
		
		var_arr = Array(g_var_client_sn,str_contract_sn)
		arr_result = excuteSqlStatementRead(str_payment_his_sql,var_arr,CONST_VIPABC_RW_CONN)
		'response.Write g_str_sql_statement_for_debug
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then				
			'response.Write("欄:" & Ubound(arr_result) & "<br />")
			'response.Write("列:" & Ubound(arr_result,2) & "<br />")
			bol_payment_record_history_tablerow = Ubound(arr_result,2)            	    '付款資訊的列數(可能會有多筆)
			bol_payment_record_history = true	
				if bol_payment_record_history= true then
		%>
        <!--力史start-->
        <div class="box">
        	<div class="line"></div>
            <div class="title"><%=getWord("buy_1_34")%></div>
            <div class="con3">
            <!--缴款类别 title-->
            <div class="name1">
            	<div class="pt1"><%=getWord("BUY_1_26")%></div>
                <div class="pt"><%=getWord("BUY_1_27")%></div>
                <div class="pt"><%=getWord("BUY_1_28")%></div>
                <div class="pt"><%=getWord("BUY_1_29")%></div>
                <div class="pt"><%=getWord("BUY_1_35")%></div> 
                <div class="pt"><%="发票代码"%></div>
                <div class="pt"><%=getWord("BUY_1_36")%></div>
                <div class="clear"></div>
            </div>
            <!--缴款类别 title-->
                            
            <%
                            for str_tablerow = 0 To Ubound(arr_result, 2)
                                str_dsn = arr_result(0, str_tablerow)
                                str_paytype = arr_result(1, str_tablerow)
                                str_paymode = arr_result(2,str_tablerow)
                                int_damount = arr_result(3,str_tablerow)
                                str_paydate = arr_result(4,str_tablerow)
                                str_invoice = arr_result(5, str_tablerow)
                                str_invoiceno = arr_result(6,str_tablerow)
								
								
								
                                if str_invoiceno="" or isnull(str_invoiceno) then str_invoice=""
								
                                str_pay_amount = str_pay_amount + int_damount
                                '=================================================================================
                                'For 新的發票 Bgein
								'Dim str_new_invoice	: str_new_invoice = ""
								'Dim str_new_invoiceno : str_new_invoiceno = ""
												
                                'str_invoice         = ""
                                'str_invoiceno       = ""
                                'str_invoice = str_invoice
								
                                if (not isEmptyOrNull(str_invoiceno)) then str_new_invoiceno =  str_invoiceno&"<br />"
                                if ( isEmptyOrNull(str_invoiceno)) then str_invoice = ""
								 
                                str_new_invoice_sql="SELECT * FROM client_invoice_record where dsn=@dsn and breakway IS NULL"
                                var_invoice_arr = Array(str_dsn)
                                arr_invoice_result = excuteSqlStatementRead(str_new_invoice_sql,var_invoice_arr,CONST_VIPABC_RW_CONN)
								'response.Write g_str_sql_statement_for_debug
                                if (isSelectQuerySuccess(arr_invoice_result)) then
                                    if (Ubound(arr_invoice_result) >= 0) then
                                        str_invoice_tablerow = Ubound(arr_invoice_result,2)       
                                        for str_invoice_tablerow = 0 To Ubound(arr_invoice_result, 2)
                                           
                                            str_new_invoice = arr_invoice_result(5, str_invoice_tablerow)
                                            str_new_invoiceno =  arr_invoice_result(4, str_invoice_tablerow)
                                            invoice_num = arr_invoice_result(19, str_invoice_tablerow)
                                            'if str_new_invoice = 0 then
                                                'str_new_invoice = getWord("BUY_1_37") '二聯式發票
                                            'else
                                                'str_new_invoice = getWord("BUY_1_38")  '三聯式發票
                                            'end if
                                                
                                            'if ( str_new_invoice <> str_invoice )  then '當發票種類不一樣的時候
                                            	'if str_invoice <>"" then str_invoice = str_invoice &"<br />"
                                                'str_invoice = str_invoice & str_new_invoice
                                            'end if
                                                
                                            'str_invoiceno = str_invoiceno& str_new_invoiceno&"<br />"
                                        next
                                    end if
                                end if
                                'For 新的發票 End
								
								'20120319 Ian [RD12031515]VIPABC官網繁體中文字樣更正 Bgein
								if( str_paytype = "繳清全額" ) then
									str_paytype = "缴清全额"
								elseif( str_paytype = "繳部份金額" ) then
									str_paytype = "缴部份金额"
								elseif( str_paytype = "補齊尾款" ) then
									str_paytype = "补齐尾款"
								elseif( str_paytype = "退費" ) then
									str_paytype = "退费"
								else
									str_paytype = str_paytype
								end if
								
								if ( str_paymode = "匯款" ) then
									str_paymode = "汇款"
								elseif ( str_paymode = "線上刷卡" ) then
									str_paymode = "线上刷卡"
								elseif ( str_paymode = "折讓" ) then
									str_paymode = "折让"
								elseif ( str_paymode = "現金" ) then
									str_paymode = "现金"
								elseif ( str_paymode = "POS機" ) then
									str_paymode = "POS机"	
								else 
									str_paymode = str_paymode
								end if
								
								'20120319 Ian [RD12031515]VIPABC官網繁體中文字樣更正 End
                                '=================================================================================
								%>
                            <!--show-->
                            <div class="show1">
                            	<div class="pt1" style="height:45px; line-height:45px;"><%=str_paytype%></div>
                                <div class="pt" style="height:45px;"><%=str_paymode%></div>
                                <div class="pt" style="height:45px; line-height:45px;"><%=str_paydate%></div>
                                <div class="pt" style="height:45px; line-height:45px;"><%=formatnumber(int_damount,0)%></div>
                                <div class="pt" style="height:45px; line-height:45px;"><%="普通发票"%></div>
                                <div class="pt" style="height:45px; line-height:45px;"><%="231001070052"%></div>
                                <div class="pt" style="height:45px; line-height:45px;"><%=invoice_num%></div>
                                <div class="clear"></div>
                            </div>
                            <!--show-->
                            <%next%> 
            </div>
        </div>
        <!--力史start-->
        <%
				end if
			end if
		'=========ryanlin add by 2010/05/14====================================
		 Dim str_amount_sql : str_amount_sql="" '檢查付款明細
		 Dim arr_amount_money '金額陣列
		 Dim str_amount_tablerow '長度
		 Dim int_total_amount_paymoney : int_total_amount_paymoney = 0 '總已付款金額
		 str_amount_sql = str_amount_sql&"select damount, convert(varchar,paydate,111) as paydate,dsn"  
		 str_amount_sql = str_amount_sql&" from contract_pay "
		 str_amount_sql = str_amount_sql&" where client_sn="&g_var_client_sn&" and sn ="&str_contract_sn&" and paytype<>'退費' and " 
         str_amount_sql = str_amount_sql&" ((paymode in ('客戶線上刷卡','線上刷卡','按月分期付款/信用卡代扣','按月分期付款/借記卡代扣') and isbuy = 1) or (paymode in ('POS機','匯款','ATM轉帳') and dauditor<>'')) "
		 str_amount_sql = str_amount_sql&" order by dsn "
		 'response.Write str_amount_sql
		 arr_amount_money = excuteSqlStatementReadQuick(str_amount_sql,CONST_VIPABC_RW_CONN)
		 if (isSelectQuerySuccess(arr_amount_money)) then
			if (Ubound(arr_amount_money) >= 0) then      
				for str_amount_tablerow = 0 To Ubound(arr_amount_money, 2)
					int_total_amount_paymoney = int_total_amount_paymoney + arr_amount_money(0,str_amount_tablerow)
				next
			end if
		 end if
		'=========ryanlin add by 2010/05/14====================================
		%>
        <div align="left">已付金额(人民币)<span id="pay1"><%=formatnumber(int_total_amount_paymoney,0) %></span></div>
        <div align="left">未付金额(人民币)<span id="pay2"><%=formatnumber(str_amount - int_total_amount_paymoney,0) %></span></div>
        <%
		end if
		set str_payment_his_sql = nothing
		set var_arr = nothing
		set arr_result = nothing
        '歷史付款明細紀錄結束
		'===========================================================================
        %>
        
		<%if (bol_isbuy = true) then%>
        <div class="w color_1"><%=getWord("BUY_1_39")%><br />
        <font color="red">
        <br />
		&nbsp;&nbsp;&nbsp;&nbsp;交易提醒:<br />
		1. 请使用IE6、IE7、IE8及IE内核的浏览器进行支付。<br />
		2. 避免使用IE9、360、TT、遨游、谷歌Chrome、苹果电脑、火狐Firefox等浏览器。<br />
		3. 点选后请留意浏览器上方是否出现”安全登录控件”请点选安装 (<a href="https://img.99bill.com/jt/u/i/securitycenter/pic/q1.jpg" target="_blank">详细说明</a>)。<br />
		4. 请关闭360安全卫士、360网购保镖。<br />
		5. 预先安装 <a href="https://www.99bill.com/b/0.html" target="_blank">块钱安全登录控件</a>。<br />
		6. 预先安装 <a href="http://www.cmbchina.com/personalbank/netpay/Faq.htm" target="_blank">招商银行安全登录控件</a> (第20条)。<br /></font>
        	<div class="t-right m-top5">
        		<input type="submit" name="button" id="button" value="+ <%=getWord("BUY_1_40")%>" class="btn_1" onClick="return chkdata()"/>
            </div>
        </div>
       <%end if%>  
        </Form>
    </div>
    </div>
    <!--內容end-->
	<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
