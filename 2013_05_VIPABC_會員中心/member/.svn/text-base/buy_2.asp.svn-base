<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type="text/javascript" src="/lib/javascript/verify.js"></script>
<script type="text/javascript" src="/lib/javascript/error_handle.js"></script>
<script type="text/javascript">
function chglayer(index)
{
	if(index==2){
		div_invoice_1.style.display = '';
	}else{
		div_invoice_1.style.display = 'none';
	}
}

function chk_data()
{
	// 取出各節點
    var str_cname = $("#txt_cname").val();
	var str_tel_code = $("#txt_tel_code").val();
	var str_tel = $("#txt_tel").val();
	var str_email = $("#txt_email").val();
	var str_address = $("#txt_client_address").val();
	var str_invoice_num = $("#txt_invoice_num").val();
	var str_company = $("#txt_company").val();
	//var str_pay_cash = $("#pay_cash").val();
	var str_cphone = str_tel_code+str_tel;
	
	//檢驗姓名格式
	/*if (g_obj_verify.isChineseNameValid(str_cname) == g_obj_error_handle.error_code.input_data_is_empty)
	{
		alert("<%=getWord("BUY_2_2")%>");
		$("#txt_cname").focus();
		return false;
	}
	if (g_obj_verify.isChineseNameValid(str_cname) != g_obj_error_handle.success_code ) 
	{
		alert("<%=getWord("CONTACT_US_NAMEERROR")%>");
		$("#txt_cname").focus();
		return false;
	}*/
	
	//檢驗手機格式
	/*if (g_obj_verify.isPhone(str_cphone,10,20) == g_obj_error_handle.error_code.input_data_is_empty )
	{
		alert("<%=getWord("BUY_2_4")%>");
		$("#txt_tel_code").focus();
		return false;     
	}
	if (g_obj_verify.isPhone(str_cphone,10,20) != g_obj_error_handle.success_code )
	{
		alert("<%=getWord("CONTACT_US_MOBILEERROR")%>");
		$("#txt_tel_code").focus();
		return false;     
	}*/	    
	
	//檢驗email格式    
	/*if (g_obj_verify.isEmailAddressValid(str_email) == g_obj_error_handle.error_code.input_data_is_empty)
	{
		alert("<%=getWord("BUY_2_4")%>"); //请输入刷卡人电子邮件
		$("#txt_email_addr").focus();
		return false;
	}
	if (g_obj_verify.isEmailAddressValid(str_email) != g_obj_error_handle.success_code)
	{
		alert("<%=getWord("CONTACT_US_EMAILERROR")%>");
		$("#txt_email_addr").focus();
		return false;
		
	}*/
	//checkEmailIsMember(str_email);
	
	////統一發票
//	if ( $("input[name='invoice']:checked").length == 0)
//	{
//		alert("<%=getWord("BUY_2_8")%>");
//		$("#invoice").focus();
//		return false;
//	}
//	
//	//判斷三聯式發票
//	if (document.getElementById('invoice1').checked)
//	{
//		//TODO：發票號碼規則未定
//		//if (g_obj_verify.isNumberOrEnglish(str_invoice_num) == g_obj_error_handle.error_code.input_data_is_empty )
//		if (str_invoice_num == "" )
//		{
//			alert("<%=getWord("BUY_2_33")%>");
//			$("#txt_invoice_num").focus();
//			return false;
//		}
//		
//		if (str_company == "" )
//		{
//			alert("<%=getWord("BUY_2_32")%>");
//			$("#txt_company").focus();
//			return false;
//		}
//	}
	
	//檢驗地址
	/*if (str_address == "")
	{
		alert("<%=getWord("BUY_2_31")%>");
		$("#txt_client_address").focus();
		return false;
	}
	*/
	//驗證我已阅读并退貨條款
	if ( $("input[name='chk_agree']:checked").length == 0)
	{
		alert("<%=getWord("BUY_2_9")%>");////请确认是否同意使用条款之内容！
		$("#chk_agree").focus();
		return false;
	}
		

}

</script>
<%
'定義變數；

Dim str_client_fullname : str_client_fullname = session("client_fullname")   '接下session值-->client_fullname
Dim int_product_sn : int_product_sn = ""                  '購買產品product_sn
Dim str_product_name : str_product_name = request("str_product_name")   '購買產品name
Dim str_sales : str_sales = request("str_sales")   'str_sales
Dim int_dsn : int_dsn = request("dsn")

'response.write "dsn......"&int_dsn&"<br>"

'塞資料的相關設定值
Dim var_arr 																	'傳excuteSqlStatementRead的陣列
Dim arr_result																'接回來的陣列值
Dim str_tablecol 																'table欄位
Dim str_tablerow 																'table列數

Dim int_paymoney : int_paymoney = 0

Dim str_client_address : str_client_address = "" '客戶住址

'客戶的基本資料
Dim client_personal_sql : client_personal_sql = ""
Dim str_email : str_email = ""                                              '刷卡人E-Mail
Dim str_cphone_code : str_cphone_code = ""                                  '刷卡人联络电话--cphone_code
Dim str_cphone : str_cphone = ""                                            '刷卡人联络电话--cphone
Dim str_cname : str_cname = ""                                              '刷卡人姓名
Dim str_amount : str_amount = ""                                            '本次应付款金额
Dim str_paymode : str_paymode = ""                                          '产品购买方式
Dim str_paytype : str_paytype = ""                                          '本次缴款类别
Dim str_lead_sn : str_lead_sn = ""                                          'lead_sn
Dim str_invoice : str_invoice = ""                                          '发票种类
Dim str_invoiceno : str_invoiceno = ""                                      '发票号码
Dim str_invoice_header : str_invoice_header = ""                            '发票抬頭
Dim int_valid_month : int_valid_month = ""                                  '產品期間(月)，用來判斷兩年期的產品才能刷卡分期
Dim client_address_sql : client_address_sql = ""							'捉客戶地址 lead_basic.client_address

Dim bol_iscustom : bol_iscustom = false                                     '判斷是否有購買
Dim int_detail_isbuy : int_detail_isbuy = 0 '是否付款client_payment_detail.isbuy
Dim int_contract_valid : int_contract_valid = 0	'client_temporal_contract.valid

'暫時先寫死作測試用的 Begin
'if (g_var_client_sn = "") then g_var_client_sn= "149532" else  g_var_client_sn = g_var_client_sn end if    '102334
'if (str_client_fullname = "") then str_client_fullname= "DaisyTseng" else  str_client_fullname = str_client_fullname end if 
'if (int_product_sn = "") then int_product_sn= "224" else  int_product_sn = int_product_sn end if 
'暫時先寫死作測試用的 End



if ( not isEmptyOrNull(g_var_client_sn) ) then
	'===================================================
	'客戶基本資料抓取Begin
	client_personal_sql="SELECT client_payment_record.contract_sn, client_temporal_contract.client_sn,       "
	client_personal_sql = client_personal_sql & " client_payment_detail.amount,  "
	client_personal_sql = client_personal_sql & " cfg_product.product_temp_name, client_basic.email,"
	client_personal_sql = client_personal_sql & " client_basic.cphone_code, client_basic.cphone,  "
	client_personal_sql = client_personal_sql & " client_basic.cname, client_temporal_contract.prodbuy, "
	client_personal_sql = client_personal_sql & " CASE client_temporal_contract.ptype WHEN 1 THEN '一次付清' ELSE '分期付款' END AS pm, "
	client_personal_sql = client_personal_sql & " cfg_pay_type.data1 AS pt ,client_temporal_contract.lead_sn, "
	client_personal_sql = client_personal_sql & " client_payment_record.invoice,client_payment_record.invoice_num,  "
	client_personal_sql = client_personal_sql & " client_payment_record.invoice_header,"
	client_personal_sql = client_personal_sql & " client_payment_detail.isbuy,client_temporal_contract.valid, cfg_product.valid_month "
	client_personal_sql = client_personal_sql & " FROM client_temporal_contract "
	client_personal_sql = client_personal_sql & " INNER JOIN client_payment_record  "
	client_personal_sql = client_personal_sql & " ON client_temporal_contract.sn = client_payment_record.contract_sn "
	client_personal_sql = client_personal_sql & " INNER JOIN client_payment_detail  "
	client_personal_sql = client_personal_sql & " ON client_payment_record.sn = client_payment_detail.payrecord_sn "
	client_personal_sql = client_personal_sql & " LEFT OUTER JOIN cfg_pay_type ON client_payment_record.paytype = cfg_pay_type.sn "
	client_personal_sql = client_personal_sql & " LEFT OUTER JOIN client_basic "
	client_personal_sql = client_personal_sql & " ON client_temporal_contract.client_sn = client_basic.sn  "
	client_personal_sql = client_personal_sql & " LEFT OUTER JOIN cfg_product "
	client_personal_sql = client_personal_sql & " ON client_temporal_contract.prodbuy = cfg_product.sn  "
	client_personal_sql = client_personal_sql & " LEFT OUTER JOIN client_purchase  "
	client_personal_sql = client_personal_sql & " ON client_temporal_contract.sn = client_purchase.contract_sn "
	client_personal_sql = client_personal_sql & " WHERE (client_payment_detail.payment_mode in (23,34))   "
	client_personal_sql = client_personal_sql & " AND (client_payment_detail.sn =@dsn ) "
	client_personal_sql = client_personal_sql & " AND (client_temporal_contract.client_sn = @client_sn) "
	
	var_arr = Array(int_dsn,g_var_client_sn)
	arr_result = excuteSqlStatementRead(client_personal_sql,var_arr,CONST_VIPABC_RW_CONN)
	'response.write g_str_sql_statement_for_debug
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then
			'Response.write("欄:" & Ubound(arr_result) & "<br>")
			'Response.write("列:" & Ubound(arr_result,2) & "<br>")	
			
			'for str_tablerow = 0 To Ubound(arr_result, 2)
				str_amount = arr_result(2, str_tablerow)     
				str_email = arr_result(4, str_tablerow)                                       
				str_cphone_code = arr_result(5, str_tablerow)                                           
				str_cphone = arr_result(6, str_tablerow)                                              
				str_cname = arr_result(7, str_tablerow)                                   
				str_paymode = arr_result(9, str_tablerow)                                                 
				str_paytype = arr_result(10, str_tablerow)                                                      
				str_lead_sn = arr_result(11, str_tablerow)                                                      
				str_invoice = arr_result(12, str_tablerow)                                                    
				str_invoiceno = arr_result(13, str_tablerow)                                          
				str_invoice_header = arr_result(14, str_tablerow)  
				int_product_sn =  arr_result(8, str_tablerow) 
				int_detail_isbuy =  arr_result(15, str_tablerow) 
				int_contract_valid = arr_result(16, str_tablerow)   
                int_valid_month = arr_result(17, str_tablerow) 	
                if (isEmptyOrNull(int_valid_month) ) then				
				    int_valid_month = 0
				end if
			'next
		end if
	end if

	set client_personal_sql = nothing
	set var_arr = nothing
	set arr_result = nothing
	
		'客戶基本資料抓取End
	'==========================================================
	
	'-------------- 判斷是否已付款成功 --------- start --------------
	if ( int_detail_isbuy = 1 and int_contract_valid = 1 ) then
	'getWord("BUY_2_1")
		Call alertGo("此筆款項已付款!!", "/program/member/buy_1.asp", CONST_ALERT_THEN_REDIRECT )
		response.end 
	end if 
	'-------------- 判斷是否已付款成功 --------- end --------------

    '----------add by ryanlin 20100401 修改地址BUG-------start------
	Dim var_arr_client,str_address_sql
	Dim int_province_number,int_area_number '地址，地區
	'開始去XML裡面撈相關的地區資料
	Dim obj_addr_xml, obj_prov_node
	Dim obj_province_node : obj_province_node = null '省份的節點
	Dim obj_area_node : obj_area_node = null '省份的節點
	Dim str_province_name : str_province_name ="" '省份
	Dim str_area_name : str_area_name = "" '地區
	'======================================
	'抓取客戶地址Begin
	if ( not isEmptyOrNull(str_lead_sn) ) then		
		client_address_sql="SELECT client_address,clb_province,clb_area from lead_basic where sn=@lead_sn"
		var_arr = Array(str_lead_sn)
		arr_result = excuteSqlStatementRead(client_address_sql,var_arr,CONST_VIPABC_RW_CONN)
		'response.write g_str_sql_statement_for_debug
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then
				str_client_address=Trim(arr_result(0, 0))	'address
                int_province_number = arr_result(1,0)   
		        int_area_number = arr_result(2,0)
			end if
		end if
	end if
	set client_address_sql = nothing
	set var_arr = nothing
	set arr_result = nothing
	'抓取客戶地址End
	'=========================================
end if

'if(request("CustomMoney")="1trueon100") then bol_iscustom = true
'導到產品介紹頁
if( isEmptyOrNull(int_product_sn) ) then
	'Call alertGo(getWord("BUY_2_1"), "/program/member/member.asp", CONST_ALERT_THEN_REDIRECT )
	'response.end 
end if

	Set obj_addr_xml = New XML
	If (obj_addr_xml.load(Server.Mappath("/xml/china_province_and_area.xml"))) Then
		'每一省434					
		'path="/bookstore/book/title[@lang='eng']" 				
		if int_province_number <> "" and int_area_number <> "" then
			Set obj_province_node = obj_addr_xml.getSingleNodeObj("/Location/Province[@value='"&int_province_number&"']")
			str_province_name = obj_province_node.getAttribute("name")
			Set obj_area_node = obj_addr_xml.getSingleNodeObj("/Location/Province[@value='"&int_province_number&"']/Area[@value='"&int_area_number&"']")
			str_area_name = obj_area_node.getAttribute("name")
		end if
	end if
	Set obj_addr_xml = Nothing
	'----------add by ryanlin 20100401 修改地址BUG--------end-----

%>
<div id="temp_contant">
<form name="fr_down_payment" method="post" action="buy_3.asp" >
<input type='hidden' name='productsn' value='<%=int_product_sn%>'>
<input type='hidden' name='productname' value='<%=str_product_name%>'>
<input type='hidden' name='sales' value='<%=str_sales%>'>
<input type='hidden' name='dsn' value='<%=int_dsn%>'>
<input type='hidden' name='iscustom' value='<%if (bol_iscustom) then response.write "1" else response.write "0" end if%>'>
<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
<input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
<input type="hidden" name="txt_email" value="<%=str_email%>" />
<input type="hidden" name="pay_cash" value="<%=str_amount%>" />
<input type="hidden" name="valid_month" value="<%=int_valid_month%>" />

    <!--內容start-->
    <div class="main_membox">
    <div class="buy">
   		<div class="w"><%=getWord("BUY_1_3")%>
        	<span class="bold"><%=str_client_fullname%></span><%=getWord("BUY_2_10")%><span class="color_4">（<%=getWord("BUY_2_11")%>）</span>
        </div>
    	<!--明細start-->
    	<div>
    		<div class="box">
    			<!--刷卡人姓名 start-->
    			<div class="con">
    				<div class="name"><%=getWord("BUY_2_12")%></div>
    				<div class="show"><%=str_cname%>
                    </div>
    				<div class="clear"></div>
    			</div>
    			<!--刷卡人姓名 end-->
                
    			<!--刷卡人联络电话 start-->
    			<div class="con">
    				<div class="name"><%=getWord("BUY_2_13")%></div>
    				<div class="show">
   						<%=str_cphone_code&str_cphone%>
    				</div>
    				<div class="clear"></div>
    			</div>
    			<!--刷卡人联络电话 end-->
                
    			<!--刷卡人E-Mail start-->
    			<div class="con">
    				<div class="name"><%=getWord("BUY_2_14")%></div>
   	 				<div class="show"><%=str_email%>
    				</div>
    				<div class="clear"></div>
    			</div>
    			<!--刷卡人E-Mail end-->
                
                <!--产品购买方式 start-->
                <div class="con">
                    <div class="name"><%=getWord("BUY_2_15")%></div>
                    <div class="show"><%=str_paymode%></div>
                    <div class="clear"></div>
                </div>
                <!--产品购买方式end-->
                
                <!--本次缴款类别 start-->
                <div class="con">
                    <div class="name"><%=getWord("BUY_2_16")%></div>
                    <div class="show"><%=str_paytype%></div>
                    <div class="clear"></div>
                </div>
                <!--本次缴款类别 end-->
                
                <!--本次应付款金额 start-->
                <div class="con">
                    <div class="name"><%=getWord("BUY_2_17")%></div>
                    <div class="show"><%=str_amount%></div>
                    <div class="clear"></div>
                </div>
                <!--本次应付款金额 end-->
                
                <!--发票类别 start-->
<!--                <div class="con">
                    <div class="name"><%=getWord("BUY_2_18")%></div>
                    <div class="show">
                        <input type="radio" name="invoice" id="invoice0" value="0" onclick="chglayer(1);" <%if str_invoice="0" then%>checked<%end if %> /><%=getWord("BUY_2_19")%>
                        <input type="radio" name="invoice" id="invoice1" value="1" onclick="chglayer(2);" <%if str_invoice="1" then%>checked<%end if %> /><%=getWord("BUY_2_20")%></div>
                    <div class="clear"></div>
                </div>
-->                <!--发票类别 end-->
               
                <!--三聯式发票 end-->
<!--                <div id="div_invoice_1" <%if (str_invoice = "1" ) then %>style="display:'';" <%else%>style="display:none;"<%end if %>>
                    <div class="con">
                        <div class="name"><%=getWord("BUY_2_29")%></div>
                        <div class="show">
                        <input type='text' name='txt_invoice_num' id="txt_invoice_num" class="login_box3 width1 m-top5" value="<%=str_invoiceno%>" maxlength='8' ></div>
                        <div class="clear"></div>
                        
                    </div>
                    
                    <div class="con">
                    <div class="name"><%=getWord("BUY_2_30")%></div>
                        <div class="show">
                        <input type='text' name='txt_company' id="txt_company" value="<%=str_invoice_header%>" class="login_box3 width1 m-top5" maxlength='100' ></div>
                        <div class="clear"></div>
                    </div>
                </div>-->
                <!--三聯式发票 end-->
                
                <!--发票邮寄地址 start-->
                <div class="con">
                    <div class="name"><%=getWord("BUY_2_21")%></div>
                    <div class="show"><%=str_province_name&str_area_name&str_client_address%>
                    </div>
                    <div class="clear"></div>
                </div>
                <!--发票邮寄地址 end-->
                
                <!--服务条款 start-->
                <div class="con">
                    <div class="name"><%=getWord("BUY_2_22")%></div>
                    <div class="show">
                        <input type="checkbox" name="chk_agree" value='1'/><%=getWord("BUY_2_23")%>
                            <a target="_blank" href="/program/terms/service.asp"><%=getWord("BUY_2_24")%></a><%=getWord("BUY_2_25")%><span class="m-left10 bold"><!--<a href="#"><%=getWord("BUY_2_26")%></a>--></span>
                    </div>
                    <div class="clear"></div>
                </div>
                <!--服务条款 end-->
        	</div>
        	<div class="w2"><input type="submit" name="next_step_button" id="next_step_button" value="+ <%=getWord("BUY_2_27")%>" class="btn_1" onclick="return chk_data();return false;"></div>
        </div>
        <!--明細end-->
       
        <!--說明start-->
        <div class="box2" style="display:none;" >
        <%=getWord("BUY_2_28")%>
        </div>
        <!--說明end-->
        
    </div>
    </div>
    <!--內容end-->
</form>        
<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>       
</div>
