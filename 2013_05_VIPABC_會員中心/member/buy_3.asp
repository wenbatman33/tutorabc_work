<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
Dim str_cname : str_cname = "" '中文姓名
Dim str_client_address : str_client_address = "" '客戶地址
Dim str_tel_code : str_tel_code = ""  '區碼電話
Dim str_txt_email : str_txt_email = ""  'email
Dim str_txt_tel : str_txt_tel = ""  '電話
Dim str_iscustom  : str_iscustom = "" '是否購買
Dim str_pay_cash : str_pay_cash = ""  '付款金額
Dim str_dsn : str_dsn = ""  'payment_detail_序號
Dim str_productsn : str_productsn = ""  '產品序號
Dim str_valid_month : str_valid_month = ""  '產品期間(月)，用來判斷兩年期的產品才能刷卡分期

str_cname = trim(request("txt_cname"))
str_client_address = trim(request("txt_client_address"))
str_tel_code = trim(request("txt_tel_code"))
str_txt_email = trim(request("txt_email"))
str_txt_tel = trim(request("txt_tel"))
str_iscustom = trim(request("iscustom"))
str_pay_cash = trim(request("pay_cash"))
str_dsn = trim(request("dsn"))
str_productsn = trim(request("productsn"))
str_productname = trim(request("productname"))
str_productname = replace(str_productname, "：","") '去除產品特殊符號:(全形)
str_productname = replace(str_productname, ":","") '去除產品特殊符號:
str_sales = trim(request("sales"))
str_valid_month = trim(request("valid_month"))

if (g_var_client_sn = "" or str_cname = "" or str_txt_email = "" or str_iscustom = "" or str_dsn = "" or str_pay_cash = "" or str_productsn = "") then
	'如果session 掉了 導回首頁
	alertGo "网页逾时，请重新登入","http://"&CONST_VIPABC_WEBHOST_NAME&"/index.asp",CONST_ALERT_THEN_REDIRECT
end if
%>
<%
dim qucikpayway : qucikpayway = "http://www.vipabc.com/program/member/QuickPay/QuickPayService/SendToBank"
dim cardconstructionpayway : cardconstructionpayway = "http://www.vipabc.com/program/member/CardConstruction/QuickPayService/SendToBank"
dim agriculturalbankpayway : agriculturalbankpayway = "http://www.vipabc.com/program/member/agriculturalbankpay/CustomerPay/SendToBank"
dim cardcommercepayway : cardcommercepayway = "http://www.vipabc.com/program/member/CardCommerce/QuickPayService/SendToBank"
dim cardCommunicationspayway : cardCommunicationspayway = "http://www.vipabc.com/program/member/CardCommunications/QuickPayService/SendToBank"
dim cardCommunicationsAPIpayway : cardCommunicationsAPIpayway = "http://www.vipabc.com/program/member/CardCommunicationsAPI/QuickPayService/SendToBank"
dim cardMerchantspayway : cardMerchantspayway = "http://www.vipabc.com/program/member/CardMerchants/QuickPayService/SendToBank"
dim CardAlipaypayway : CardAlipaypayway = "http://www.vipabc.com/program/member/CardAlipay/QuickPayService/SendToBank"
%>
<style>
body{ margin:0;}
#pay{ width:698px; margin:0 auto; font-size:12px;}
.fontbuy3{ font-size:14px; font-weight:bold;height:50px;}
.paytitle{ width:688px; height:20px; border-bottom:dotted 1px #cfd2d7; padding:10px 0px 0px 10px; font-size:14px; color:#4d4d4d;}
.kjpay{ height:150px; background:#fffbe1; border-top:dotted 1px #febf90; border-bottom:dotted 1px #febf90; margin-top:20px;}
.STYLE1 {color: #FF0000}
.linepay{margin-top:25px;}
.STYLE2 {font-size: 14px}
.btn{ margin-top:30px; padding-left:30px;}
.btnclass{ width:98px; height:25px; background:#ed8114; color:#fff; cursor:pointer; border:solid 1px #ec7701;}
.hid{ display:none; margin-top:10px; margin-left:35px; }
.btn_submit{
    background-image:url(/images/loading2.gif);
	background-repeat:no-repeat;
	background-position:0 0;
	width:50px;
	height:50px;
	margin:0 auto;
	padding-top:10px;
}
</style>
<script type="text/javascript">
<!--
//將支付寶刷卡相關資料寫進資料庫中
function writeDataToHitrustRecord(p_str_cname, p_int_product_sn, p_int_tel_code, p_int_tel, p_str_email, p_bol_iscustom, p_int_pay_cash, p_int_dsn)
{
    //支付寶相關的刷卡資料寫進資料庫
    $.ajax(
    {
        type: "POST",
        url: "ajax_write_payment_record.asp",
        cache: false,
        async: false,
        data: { txt_cname: escape(p_str_cname), productsn: escape(p_int_product_sn), txt_tel_code: escape(p_int_tel_code), txt_tel: escape(p_int_tel), txt_email: escape(p_str_email), iscustom: escape(p_bol_iscustom), pay_cash: escape(p_int_pay_cash), dsn: escape(p_int_dsn) },
        error: function(xhr)
        {
            alert("錯誤" + xhr.responseText);
        },
        success: function(result)
        {
            if (result != "")
            {
                    //如果有成功，導向支付寶頁面  
                $(location).attr("href",result);                     
            }
            else
            {                    
                //如果沒有值
                alert("资料错误，请重新登入");
                    $(location).attr("href","/program/member/member_login.asp");          
            }
        },
        complete: function()
        {
            $.unblockUI();
        }
    });
}
//-->

function gotopay()
{
	document.getElementById("submitbutton1").style.display="none";
	document.getElementById("submitbutton2").style.display="block";
	
    //alert($("#paywayvalue").val());
	if ($("#paywayvalue").val() == 1) //銀聯
	{
	    document.fr_QuickPay_payment.submit();
	}
	else if ($("#paywayvalue").val() == 2) //交通分期
	{
		//一次性付款才能選擇分期
		<%if Session("direct_contract_sn") = "Y" then%>
	    var e = document.getElementById("periodselect").options[document.getElementById("periodselect").selectedIndex].value;
	    //alert(e);
		$("#period").val(e);
		<%end if%>
        document.fr_CardCommunications_payment.submit();
	}
	else if ($("#paywayvalue").val() == 3) //招商
	{
	    document.fr_cmbchina_down_payment.submit();
	}
	else if ($("#paywayvalue").val() == 4) //建設
	{
	    document.fr_construction_payment.submit();
	}
	else if ($("#paywayvalue").val() == 5) //農業
	{
	    document.fr_agriculturalBankPay_payment.submit();
	}
	else if ($("#paywayvalue").val() == 6) //交通直聯
	{
	    document.fr_CardCommunicationsAPI_payment.submit();
	}
	else if ($("#paywayvalue").val() == 7) //工商
	{
		     //一次性付款才能選擇分期
		<%if Session("direct_contract_sn") = "Y" then%>
		var e = document.getElementById("periodselect2").options[document.getElementById("periodselect2").selectedIndex].value;
	    //alert(e);
		$("#period2").val(e);
		<%end if%>
	    document.fr_commerce_payment.submit();
	}
	else if ($("#paywayvalue").val() == 8) //民生
	{
		alert("系统维护中，请使用其它付款方式");
	}
	else if ($("#paywayvalue").val() == 9) //快錢
	{
	    document.fr_99bill_down_payment.submit();
	}
	else if ($("#paywayvalue").val() == 10) //國際
	{
	    document.fr_down_payment.submit();
	}
	else if ($("#paywayvalue").val() == 11) //支付寶
	{
	    //呼叫ajax寫入資料，如果成功導向支付寶頁面，失敗alert連線預時
        writeDataToHitrustRecord("<%=str_cname%>", "<%=str_productsn%>", "<%=str_tel_code%>", "<%=str_txt_tel%>", "<%=str_txt_email%>", "<%=str_iscustom%>", "<%=str_pay_cash%>", "<%=str_dsn%>");
		//document.fr_alipay.submit();
	}
	else if ($("#paywayvalue").val() == 12) //招商測試
	{
	    document.fr_CardMerchants_payment.submit();
	}
	else if ($("#paywayvalue").val() == 13) //支付寶測試
	{
	    document.fr_CardAlipay_payment.submit();
	}
	else
	{
	    alert("请选择付款方式");
		document.getElementById("submitbutton1").style.display="block";
	    document.getElementById("submitbutton2").style.display="none";
	}
}

function setPaywayValue(obj)
{
    var str_pay_cash = "<%=str_pay_cash%>";
    $("#paywayvalue").val(obj.value);
	
	//alert("<%=session("client_sn")%>__" + str_pay_cash);
	    //一次性付款才能選擇分期
	    <%if Session("direct_contract_sn") = "Y" then%>
	    if (obj.value == "2" && parseInt(str_pay_cash) >= 1500)
		{
		    document.getElementById("periodselectshow").style.display="block";
		}
        else
		{
	        document.getElementById("periodselectshow").style.display="none";
		}	
		if (obj.value == "7" && parseInt(str_pay_cash) >= 600)
		{
		    document.getElementById("periodselectshow2").style.display="block";
		}
        else
		{
	        document.getElementById("periodselectshow2").style.display="none";
		}
		<%end if%>
}

</script>
<!--內容start-->
<div class="main_membox" id="main_payment_content">
<table width="100%">
<tr>
    <td>
	<div class="buy">
        <div class="w">
        <font color="red">
		&nbsp;&nbsp;&nbsp;&nbsp;交易提醒:<br />
		1. 请使用IE6、IE7、IE8及IE内核的浏览器进行支付。<br />
		2. 避免使用IE9、360、TT、遨游、谷歌Chrome、苹果电脑、火狐Firefox等浏览器。<br />
		3. 点选后请留意浏览器上方是否出现”安全登录控件”请点选安装 (<a href="https://img.99bill.com/jt/u/i/securitycenter/pic/q1.jpg" target="_blank">详细说明</a>)。<br />
		4. 请关闭360安全卫士、360网购保镖。<br />
		5. 预先安装 <a href="https://www.99bill.com/b/0.html" target="_blank">快钱安全登录控件</a>。<br />
		6. 预先安装 <a href="http://www.cmbchina.com/personalbank/netpay/Faq.htm" target="_blank">招商银行安全登录控件</a> (第20条)。<br /></font>
		</div>
	</div>		
    </td>
</tr>
</table>

<br />	
<table border="0">
<tr><td>
                <script>//交通银行通道(分期) - start</script>
					<form name="fr_CardCommunications_payment" method="post" action="<%=cardCommunicationspayway%>">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
					<input type="hidden" name="period" Id="period" value="" />
		            </form>
				<script>//交通银行通道(分期) - end</script>		
				<script>//交通银行通道(直联) - start</script>
					<form name="fr_CardCommunicationsAPI_payment" method="post" action="<%=cardCommunicationsAPIpayway%>">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
		            </form>
				<script>//交通银行通道(直联) - end</script>	
				<script>//农业银行通道 - start</script>
					<form name="fr_agriculturalBankPay_payment" method="post" action="<%=agriculturalbankpayway%>">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
		            </form>
				<script>//农业银行通道 - end</script>
			    <script>//建设银行通道 - start</script>
					<form name="fr_construction_payment" method="post" action="<%=cardconstructionpayway%>">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
		            </form>
				<script>//建设银行通道 - end</script>		
				<script>//工商银行通道 - start</script>
					<form name="fr_commerce_payment" method="post" action="<%=cardcommercepayway%>">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
					<input type="hidden" name="period2" Id="period2" value="" />
		            </form>
				<script>//工商银行通道 - end</script>		
				<script>//国际卡(外卡) - start</script>
                    <form name="fr_down_payment" id="fr_down_payment" method="post" action="down_payment.asp">
                    <input type="hidden" name="change_bank_for_chinapay" value="2" checked />
                    <input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
					</form>
				<script>//国际卡(外卡) - end</script>		
				<script>//银联快捷支付 - start</script>
					<form name="fr_QuickPay_payment" method="post" action="<%=qucikpayway%>">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
		            </form>
				<script>//银联快捷支付 - end</script>		
				<script>//快钱支付 - start</script>
					<form name="fr_99bill_down_payment" id="fr_99bill_down_payment" method="post" action="n99bill/99bill_payment.asp">
                    <input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
	                </form>
				<script>//快钱支付 - end</script>		
				<script>//招商银行 - start</script>
					<form name="fr_cmbchina_down_payment" id="fr_cmbchina_down_payment" action="cmbchina/cmbchina_payment.asp" method="post">
                    <input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
	                </form>
				<script>//招商银行 - end</script>
				<script>//招商银行測試 - start</script>
					<form name="fr_CardMerchants_payment" id="fr_CardMerchants_payment" action="<%=cardMerchantspayway%>" method="post">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
	                </form>
				<script>//招商银行測試 - end</script>				
				<script>//支付宝 - start</script>
					<form name="fr_alipay" id="fr_alipay" method="post" action="down_payment.asp">
        		    <input type="hidden" name="txt_cname" value="<%=str_cname%>" />
        		    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
        		    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
        		    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
        		    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
        		    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
        		    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
        		    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
        		    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
                    </form>
				<script>//支付宝 - end</script>
				<script>//支付宝測試 - start</script>
					<form name="fr_CardAlipay_payment" id="fr_CardAlipay_payment" action="<%=CardAlipaypayway%>" method="post">
					<input type="hidden" name="txt_cname" value="<%=str_cname%>" />
                    <input type="hidden" name="txt_client_address" value="<%=str_client_address%>" />
                    <input type="hidden" name="txt_tel_code" value="<%=str_tel_code%>" />
                    <input type="hidden" name="txt_tel" value="<%=str_txt_tel%>" />
                    <input type="hidden" name="txt_email" value="<%=str_txt_email%>" />
                    <input type="hidden" name="iscustom" value="<%=str_iscustom%>" />
                    <input type="hidden" name="pay_cash" value="<%=str_pay_cash%>" />
                    <input type="hidden" name="dsn" value="<%=str_dsn%>" />
                    <input type="hidden" name="productsn" value="<%=str_productsn%>" />
		            <input type="hidden" name="productname" value="<%=str_productname%>" />
		            <input type="hidden" name="clientsn" value="<%=g_var_client_sn%>" />
		            <input type="hidden" name="sales" value="<%=str_sales%>" />
	                </form>
				<script>//支付宝測試 - end</script>
</td>
</tr>
<tr><td><input type="hidden" name="paywayvalue" id="paywayvalue" value=""></td></tr>
<tr>
<td>
<div id="pay">
  <div class="paytitle">选择您的付款方式</div>
    <div class="kjpay">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" height="36" class="fontbuy3">推荐支付方式</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="3%"><input type="radio" name="payway" id="payway" value="1" onclick="setPaywayValue(this);" /></td>
              <td width="31%"><img src="/images/BuyPage/1.jpg" /></td>
			  <td width="3%"><input type="radio" name="payway" id="payway" value="7" onclick="setPaywayValue(this);" /></td>
              <td width="31%"><img src="/images/BuyPage/pay_05.jpg" /></td>
              <td width="3%"><input type="radio" name="payway" id="payway" value="2" onclick="setPaywayValue(this);" /></td>
              <td width="31%"><img src="/images/BuyPage/pay_12.jpg" /></td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td valign="top"><br /><span class="STYLE1">快捷支付</span>（无需网银/单笔2万）</td>
			  <td>&nbsp;</td>
              <td valign="top"><br />
			  <span class="STYLE1">信用卡分期</span>（限工行信用卡）<br />
				  <div id="periodselectshow2" style="display:none;">
				  请选择期数
				  <select name="periodselect2" id="periodselect2">
				  <option value="0">一次付</option>
				  <option value="3">&nbsp;&nbsp;3&nbsp;期</option>
				  <option value="6">&nbsp;&nbsp;6&nbsp;期</option>
				  <option value="12">12&nbsp;期</option>
				  <%if str_valid_month = "24" then%>
				  <option value="24">24&nbsp;期</option>
				  <%end if%>
				  </select>
				  </div>
			  </td>
			  <td>&nbsp;</td>
              <td valign="top"><br />
			  <span class="STYLE1">信用卡分期</span>（限交银信用卡）<br />
		      	  <div id="periodselectshow" style="display:none;">
				  请选择期数
				  <select name="periodselect" id="periodselect">
				  <option value="0">一次付</option>
				  <option value="3">&nbsp;&nbsp;3&nbsp;期</option>
				  <option value="6">&nbsp;&nbsp;6&nbsp;期</option>
				  <option value="12">12&nbsp;期</option>
				  <%if str_valid_month = "24" then%>
				  <option value="24">24&nbsp;期</option>
				  <%end if%>
				  </select>
				  </div>
  		      </td>
            </tr>
          </table></td>
        </tr>
      </table>
    </div>
    <div class="linepay">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" height="40" valign="top"><span class="fontbuy3">网上银行：</span><span class="STYLE2">需要开通网银與网上支付功能</span></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="3%" height="63"><input type="radio" name="payway" id="payway" value="3" onclick="setPaywayValue(this);" /></td>
              <td width="28%"><img src="/images/BuyPage/pay_03.jpg" /></td>
              <td width="5%" align="center"><input type="radio" name="payway" id="payway" value="4" onclick="setPaywayValue(this);" /></td>
              <td width="30%"><img src="/images/BuyPage/pay_07.jpg" /></td>
              <td width="4%"><input type="radio" name="payway" id="payway" value="6" onclick="setPaywayValue(this);" /></td>
              <td width="30%"><img src="/images/BuyPage/pay_12.jpg" /></td>
            </tr>
            <tr>
              <td><input type="radio" name="payway" id="payway" value="5" onclick="setPaywayValue(this);" /></td>
              <td><img src="/images/BuyPage/pay_13.jpg" /></td>
              <td align="center"><%if session("client_sn") = "227238" then%><input type="radio" name="payway" id="payway" value="8" onclick="setPaywayValue(this);" /><%end if%></td>
              <td><%if session("client_sn") = "227238" then%><img src="/images/BuyPage/pay_14.jpg" /><%end if%></td>
			  <td align="center"></td>
              <td height="25"></td>
            </tr>
			
          </table></td>
        </tr>
      </table>
    </div>
    <div class="linepay">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2%">&nbsp;</td>
          <td width="98%" height="40" valign="top"><span class="fontbuy3">支付平台</span></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="3%" height="57"><input type="radio" name="payway" id="payway" value="9" onclick="setPaywayValue(this);" /></td>
                <td width="28%"><img src="/images/BuyPage/pay_18.jpg" /></td>
                <td width="5%" align="center"><input type="radio" name="payway" id="payway" value="10" onclick="setPaywayValue(this);" /></td>
                <td width="30%"><img src="/images/BuyPage/pay_19.jpg" /></td>
                <td width="4%"><input type="radio" name="payway" id="payway" value="11" onclick="setPaywayValue(this);" /></td>
                <td width="30%"><img src="/images/BuyPage/pay_20.jpg" /></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>大额支付</td>
                <td align="center">&nbsp;</td>
                <td height="25">国际卡（<span class="STYLE1">请于交易前完成3D验证</span>）</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
			<%if session("client_sn") = "227238" then%>
              <tr>
                <td width="3%" height="57"><input type="radio" name="payway" id="payway" value="12" onclick="setPaywayValue(this);" /></td>
                <td width="28%"><img src="/images/BuyPage/pay_03_test.jpg" /></td>
                <td width="5%" align="center"><input type="radio" name="payway" id="payway" value="13" onclick="setPaywayValue(this);" /></td>
                <td width="30%"><img src="/images/BuyPage/pay_20_test.jpg" /></td>
                <td width="4%"></td>
                <td width="30%"></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>招商银行改版测试</td>
                <td>&nbsp;</td>
                <td>支付宝改版测试</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
			<%end if%>
          </table></td>
        </tr>
      </table>
    </div>
    <div id="bankinfo" class="hid">八家银行支持单笔4万:工行、中行、建行、广发、兴业、光大、邮储、中信银行。</div>
	
    <div class="btn" id="submitbutton1">
      <input name="button" type="button" class="btnclass" id="button" value="+ 下一步" onclick="gotopay();" />
    </div>
    <div class="btn_submit" id="submitbutton2" style="display:none">
    </div>
			
</div>
</td>
</tr>
</table>
	<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<font color="FFFFFF"><%response.write Session("direct_contract_sn") & "__" & str_valid_month%></font>
<!--內容end-->

