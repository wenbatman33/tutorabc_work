<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<style type="text/css">
/*pay-------------------------------------------------------------*/
.pay{
	width:600px;
	left:50%;
	top:20%;
	font-size:12px;
	margin-left:50px;
	margin-top:50px;
	font-family:Verdana, Geneva, sans-serif, "新細明體";}
.pay ul li{
	width:350px;
	clear:both;}
.pay ul li.a{
	height:30px;
	+height:25px;
	line-height:30px;
	+line-height:25px;
	line-height:25px\0;}
.pay ul li.b{
	color:#999999;
	height:15px;
	line-height:15px;
	margin-bottom:10px;}
.pay ul li div{
	float:left;}
.pay ul div.pay_t{
	width:80px;
	padding-right:5px;
	text-align:right;}
.pay ul li div input.pay_input{
	width:50px;}	
.pay ul li div.pay_word{
	padding:0 10px;}
.pay div.cvs2{
	position: relative;
	color:#88AFAE;
	cursor:pointer;}
.cvs2:after {
	content: "";
	clear: both;
	display: block;}
.cvs2_pic{
	width:154px;
	display:none;
	position:absolute;
	left:100px;
	top:-65px;
	z-index:2;}
.cvs2:hover > .cvs2_pic{
	display:block;}
.pay div img{
	padding-left:10px;}
.pay_text{
	border-top:2px #333333 solid;
	padding-top:10px;
	margin-top:10px;}
.pay_button{
	height:25px;}
/*pay END-------------------------------------------------------------*/	
</style>
</head>
<%
Dim int_client_sn : int_client_sn = ""								'客戶編號
Dim int_contract_sn : int_contract_sn = session("ach_contract_sn")	'合約編號
Dim finalexpiredate : finalexpiredate = "0000"          		    '最後一期的時間(簡易格式)
Dim finalexpiredateword : finalexpiredateword = ""          		'最後一期的時間(正確格式)
Dim strPayName : strPayName = ""
Dim strPath : strPath = ""
Dim strClientAddress : strClientAddress = ""
if ( isEmptyOrNull(int_client_sn) ) then
   int_client_sn = session("client_sn")
end if
Dim exe : exe = request("exe")

if(isEmptyOrNull(int_contract_sn)) then
    int_contract_sn = getRequest("int_contract_sn",CONST_DECODE_NO)
end if
if(isEmptyOrNull(int_client_sn)) then
    int_client_sn = getRequest("int_client_sn",CONST_DECODE_NO)
end if

'判斷是否參數有值，沒有導回首頁
if ( isEmptyOrNull(int_client_sn) OR isEmptyOrNull(int_contract_sn) ) then
	str_tmp_go_page =  "member_login.asp"
	'call sub alertGo
	Call alertGo( getWord("contract_agree_1") , str_tmp_go_page, CONST_NOT_ALERT_THEN_REDIRECT)
	Response.End
end if

'找尋
Set objVIPABC_PAYMENT_RECORD = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
strSql = "SELECT top 1 paydate, payman FROM client_payment_record WITH (NOLOCK) WHERE contract_sn = @contract_sn AND Valid = 1 order by sn desc"
intSqlWriteResult = objVIPABC_PAYMENT_RECORD.excuteSqlStatementEach(strSql,  Array(int_contract_sn) , CONST_TUTORABC_RW_CONN)
if(not objVIPABC_PAYMENT_RECORD.eof) then
    finalexpiredate = right(year(cdate(objVIPABC_PAYMENT_RECORD("paydate"))), 2) & right("0" & month(cdate(objVIPABC_PAYMENT_RECORD("paydate"))), 2)
	finalexpiredateword = year(cdate(objVIPABC_PAYMENT_RECORD("paydate"))) & "年" & right("0" & month(cdate(objVIPABC_PAYMENT_RECORD("paydate"))), 2) & "月"
    strPayName = objVIPABC_PAYMENT_RECORD("payman")
end if
objVIPABC_PAYMENT_RECORD.close

strSql = "SELECT caddr,c_province,c_area,c_region,ccode FROM lead_basic WITH (NOLOCK) WHERE client_sn = @ClientSn "
intSqlWriteResult = objVIPABC_PAYMENT_RECORD.excuteSqlStatementEach(strSql,  Array(int_client_sn) , CONST_TUTORABC_RW_CONN)
if(not objVIPABC_PAYMENT_RECORD.eof) then
    strPath = Server.Mappath("/xml/ChinaProvinceAndArea.xml")
    strValueName = findXmlValueName(strPath, objVIPABC_PAYMENT_RECORD("c_province"), objVIPABC_PAYMENT_RECORD("c_area"), objVIPABC_PAYMENT_RECORD("c_region"))
	strClientAddress = trim(strValueName & objVIPABC_PAYMENT_RECORD("caddr"))
    strClientAddress = "(" & objVIPABC_PAYMENT_RECORD("ccode") & ")" & strClientAddress
end if
objVIPABC_PAYMENT_RECORD.close
Set objVIPABC_PAYMENT_RECORD = nothing

if(exe = "yes") then
    strPayName = getRequest("PayName",CONST_DECODE_NO)		'客戶編號
    Dim intClientSn : intClientSn = getRequest("int_client_sn",CONST_DECODE_NO)		'客戶編號
    Dim intContractSn : intContractSn = getRequest("int_contract_sn",CONST_DECODE_NO)	'合約編號
    Dim exmonth : exmonth = getRequest("exmonth",CONST_DECODE_NO)	'有效期年
    Dim exyear : exyear = getRequest("exyear",CONST_DECODE_NO)	'有效期月
    Dim cvs : cvs = getRequest("cvs",CONST_DECODE_NO)	'CVS2
    Dim strCardNumber : strCardNumber = getRequest("card_number",CONST_DECODE_NO)   '卡號
    Dim strAccountName : strAccountName = getRequest("account_name",CONST_DECODE_NO)    '戶名
    Dim strIdCardNumber : strIdCardNumber = getRequest("id_card_number",CONST_DECODE_NO) '身份證字號
    Dim strCellPhone : strCellPhone = getRequest("cell_phone",CONST_DECODE_NO)  '電話
    Dim strEmail : strEmail = getRequest("email",CONST_DECODE_NO)   'email
    Dim strBankSn : strBankSn = getRequest("vipabc_ach_bank",CONST_DECODE_NO)   '銀行代號
    Dim strPayMethod : strPayMethod = getRequest("installment_type_of_payment",CONST_DECODE_NO)
    if ( isEmptyOrNull(intClientSn) ) then
       intClientSn = g_var_client_sn
    end if

    '找是否有新增過
    Set objVIPABC_ACH_CARD_INFO_FIRST_PART = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
    strSql = "SELECT top 1 CardInfoFirstPartId FROM VIPABC_ACH_CARD_INFO_FIRST_PART WITH (NOLOCK) WHERE ContractSn = @ContractSn AND ClientSn = @ClientSn AND Valid = 1 order by CardInfoFirstPartId desc"
    intSqlWriteResult = objVIPABC_ACH_CARD_INFO_FIRST_PART.excuteSqlStatementEach(strSql,  Array(intContractSn, intClientSn) , CONST_TUTORABC_RW_CONN)
    if(not objVIPABC_ACH_CARD_INFO_FIRST_PART.eof) then
        intCardInfoFirstPartId = objVIPABC_ACH_CARD_INFO_FIRST_PART("CardInfoFirstPartId")
    end if
    objVIPABC_ACH_CARD_INFO_FIRST_PART.close
    Set objVIPABC_ACH_CARD_INFO_FIRST_PART = nothing

    if(isEmptyOrNull(intCardInfoFirstPartId) or 0 = intCardInfoFirstPartId) then  '新增
        arrParameters = Array(intClientSn, intContractSn, strAccountName, cvs, "0", strBankSn, strEmail)
        strSql = " INSERT INTO [VIPABC_ACH_CARD_INFO_FIRST_PART]([ClientSn],[ContractSn],[CardBankAcctName],[Cwthorizationcode],[CreateDete],[Creator],[CardBankSn],[Email]) "
        strSql = strSql & " VALUES(@ClientSn,@ContractSn,@PayName,@CVVCode,getdate(),@StaffSn,@CardBankSn,@Email) "
        intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
        if (intSqlWriteResult <> CONST_FUNC_EXE_SUCCESS ) then
			Response.write "insert FIRST_PART failure <br/>"
            Response.write "g_str_run_time_err_msg:" & g_str_run_time_err_msg &"<br>"
		end if

        Set objVIPABC_ACH_CARD_INFO_FIRST_PART = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
        strSql = "SELECT top 1 CardInfoFirstPartId FROM VIPABC_ACH_CARD_INFO_FIRST_PART WITH (NOLOCK) WHERE ContractSn = @ContractSn AND ClientSn = @ClientSn order by CardInfoFirstPartId desc"
        intSqlWriteResult = objVIPABC_ACH_CARD_INFO_FIRST_PART.excuteSqlStatementEach(strSql,  Array(intContractSn, intClientSn) , CONST_TUTORABC_RW_CONN)
        if(not objVIPABC_ACH_CARD_INFO_FIRST_PART.eof) then
            intCardInfoFirstPartId = objVIPABC_ACH_CARD_INFO_FIRST_PART("CardInfoFirstPartId")
        end if
        objVIPABC_ACH_CARD_INFO_FIRST_PART.close
        Set objVIPABC_ACH_CARD_INFO_FIRST_PART = nothing
        
        if(not isEmptyOrNull(intCardInfoFirstPartId)) then
            arrParameters = Array(intCardInfoFirstPartId, strCardNumber, exmonth & exyear, "0")
            strSql = " INSERT INTO [CashFlow].[dbo].[VIPABC_ACH_CARD_INFO_SECOND_PART]([CardInfoFirstPartId],[CardNo],[CardExpireDate],[Valid],[CreateDate],[Creator]) "
            strSql = strSql & " VALUES(@CardInfoFirstPartId, dbo.f_RSAEncry(@CardNo,63,47,59), @CardExpireDate, 1, getdate(), @StaffSn) "
            intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
            if (intSqlWriteResult <> CONST_FUNC_EXE_SUCCESS ) then
			    Response.write "insert SECOND_PART failure <br/>"
                Response.write "g_str_run_time_err_msg:" & g_str_run_time_err_msg &"<br>"
		    end if
        end if
    else    'update
        arrParameters = Array(strAccountName, cvs, "0", strBankSn, strEmail, intClientSn, intContractSn)
        strSql = " UPDATE VIPABC_ACH_CARD_INFO_FIRST_PART "
        strSql = strSql & " SET CardBankAcctName = @PayName, Cwthorizationcode = @CVVCode, Creator = @StaffSn, CreateDete = getdate(), CardBankSn=@CardBankSn, Email=@Email "
        strSql = strSql & " WHERE ClientSn = @ClientSn AND ContractSn = @ContractSn "
        intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
        if (intSqlWriteResult <> CONST_FUNC_EXE_SUCCESS ) then
			Response.write "update FIRST_PART failure <br/>"
            Response.write "g_str_run_time_err_msg:" & g_str_run_time_err_msg &"<br>"
		end if

        arrParameters = Array(strCardNumber, exmonth & exyear, "0", intCardInfoFirstPartId)
        strSql = " UPDATE [CashFlow].[dbo].[VIPABC_ACH_CARD_INFO_SECOND_PART] "
        strSql = strSql & " SET   CardNo = dbo.f_RSAEncry(@CardNo,63,47,59), CardExpireDate = @CardExpireDate, Creator = @StaffSn, CreateDate=GETDATE() "
        strSql = strSql & " WHERE CardInfoFirstPartId = @CardInfoFirstPartId "
        intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
        if (intSqlWriteResult <> CONST_FUNC_EXE_SUCCESS ) then
			Response.write "update SECOND_PART failure <br/>"
            Response.write "g_str_run_time_err_msg:" & g_str_run_time_err_msg &"<br>"
		end if
    end if

    if(not isEmptyOrNull(intCardInfoFirstPartId)) then
        'update client_payment_record.payman
        arrParameters = Array(strPayName, intContractSn, intContractSn)
        strSql = " UPDATE  client_payment_record "
        strSql = strSql & " SET     payman = @PayName "
        strSql = strSql & " WHERE   contract_sn = @ContractSn AND valid = 1 "
        strSql = strSql & " AND sn IN ( "
        strSql = strSql & "     SELECT  client_payment_record.sn "
        strSql = strSql & "     FROM    dbo.client_payment_record WITH (NOLOCK) "
        strSql = strSql & "             INNER JOIN dbo.client_payment_detail ON client_payment_record.sn = dbo.client_payment_detail.payrecord_sn "
        strSql = strSql & "     WHERE   client_payment_record.contract_sn = @ContractSn2 "
        strSql = strSql & "             AND dbo.client_payment_detail.payment_mode IN ( 48, 50 ) ) "
        intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)

        'update client_payment_detail.payment_mode
        arrParameters = Array(strPayMethod, intContractSn)
        strSql = " UPDATE  dbo.client_payment_detail "
        strSql = strSql & " SET     payment_mode = @PayMode "
        strSql = strSql & " WHERE   sn IN ( "
        strSql = strSql & "     SELECT  client_payment_detail.sn "
        strSql = strSql & "     FROM    dbo.client_payment_record WITH ( NOLOCK ) "
        strSql = strSql & "             INNER JOIN dbo.client_payment_detail WITH ( NOLOCK ) ON client_payment_record.sn = dbo.client_payment_detail.payrecord_sn "
        strSql = strSql & "     WHERE   contract_sn = @ContractSn AND client_payment_detail.payment_mode IN ( 48, 50 )"
        strSql = strSql & " ) "
        intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
        if (intSqlWriteResult  <> CONST_FUNC_EXE_SUCCESS ) then
		    'Response.write "update detail failure: " & strSql & " <br/>"
            'Response.write "g_str_run_time_err_msg:" & g_str_run_time_err_msg &"<br>"
	    end if
        'update client_temporal_contract.pmethod
        arrParameters = Array(strPayMethod, intContractSn, intClientSn)
        strSql = " UPDATE dbo.client_temporal_contract SET pmethod = @PayMode WHERE sn = @ContractSn AND client_sn = @ClientSn "
        intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
        if (intSqlWriteResult <> CONST_FUNC_EXE_SUCCESS ) then
		    Response.write "update contract failure <br/>"
            Response.write "g_str_run_time_err_msg:" & g_str_run_time_err_msg &"<br>"
	    end if
        'update client_basic 資料
        if(not isEmptyOrNull(strIdCardNumber) and not isEmptyOrNull(strCellPhone) and Len(strCellPhone) > 4) then
            arrParameters = Array(strIdCardNumber, Left(strCellPhone, 4), Right(strCellPhone, Len(strCellPhone)-4), intClientSn)
            strSql = " UPDATE client_basic SET idno=@IDCardNumber, cphone_code=@CPhoneCode, cphone=@CPhone WHERE sn = @ClientSn "
            intSqlWriteResult = excuteSqlStatementWrite(strSql, arrParameters, CONST_TUTORABC_RW_CONN)
            if (intSqlWriteResult <> CONST_FUNC_EXE_SUCCESS ) then
		        'Response.write "update client failure <br/>"
                'Response.write "g_str_run_time_err_msg:" & g_str_run_time_err_msg &"<br>"
	        end if
        end if
    end if
end if

%>
<!--內容start-->
<script language="javascript">
function go() {
    var finalexpiredate = "<%=finalexpiredate%>";
    var PayName = document.getElementById("PayName");
    var exmonth = document.getElementById("exmonth");
    var exyear = document.getElementById("exyear");
    var cvs = document.getElementById("cvs");
    var account_name = document.getElementById("account_name");
    var card_number = document.getElementById("card_number");
	var installment_type_of_payment = document.getElementsByName("installment_type_of_payment");
	var vipabc_ach_bank = document.getElementById("vipabc_ach_bank");
	var id_card_number = document.getElementById("id_card_number");
	var cell_phone = document.getElementById("cell_phone");
	var email = document.getElementById("email");
	
	if ("" == PayName.value)
	{
	    alert("请输入付款人");
	    PayName.focus();
	    return false;
    }
	else if (installment_type_of_payment[0].checked == false && installment_type_of_payment[1].checked == false)
	{
	    alert("请选择信用卡代扣/借记卡代扣");
	    installment_type_of_payment[0].focus();
	    return false;
	}
	else if (vipabc_ach_bank.value == "" || vipabc_ach_bank.value == "0")
	{
	    alert("请选择开户行/机构");
	    vipabc_ach_bank.focus();
	    return false;
	}
	else if (account_name.value == "")
	{
	    alert("请输入户名");
	    account_name.focus();
	    return false;
	}
	else if (card_number.value == "")
	{
	    alert("请输入有账号/卡号");
	    card_number.focus();
	    return false;
	}
	else if (!checknum(card_number.value))
	{
	    alert("账号/卡号须为数字");
	    card_number.focus();
	    return false;
	}
	else if ((exmonth.value == "" || exyear.value == "") && installment_type_of_payment[0].checked == true)
	{
	    alert("请输入有效期");
	    exmonth.focus();
	    return false;
	}
	else if (cvs.value == "" && installment_type_of_payment[0].checked == true)
	{
	    alert("请输入CVS2");
	    cvs.focus();
	    return false;
	}
	else if ((!checknum(exmonth.value) || !checknum(exyear.value) || !checknum(cvs.value)) && installment_type_of_payment[0].checked == true)
	{
	    alert("有效期与CVS2须为数字");
	    exmonth.focus();
	    return false;
	}
	else if ((exmonth.value.length != 2 || exyear.value.length != 2 || parseInt(exmonth.value) > 12) && installment_type_of_payment[0].checked == true)
	{
	    alert("格式错误，有效期须为2码且為正確年月格式");
	    exmonth.focus();
	    return false;
	}
	else if (cvs.value.length != 3 && installment_type_of_payment[0].checked == true)
	{
	    alert("格式错误，CVS2须为3码");
	    cvs.focus();
	    return false;
	}
	else if ((exyear.value + exmonth.value < finalexpiredate) && installment_type_of_payment[0].checked == true)
	{
	    alert("有效期须大于最后一期时间，最后一期日期为<%=finalexpiredateword%>");
	    exmonth.focus();
	    return false;
	}
	else if (id_card_number.value == "")
	{
	    alert("请输入身份证号码");
	    id_card_number.focus();
	    return false;
	}
	else if (g_obj_verify.isNumberOrEnglish(id_card_number.value) != g_obj_error_handle.success_code || id_card_number.value.length < 10)
	{
	    alert("请输入正确的身份证号码");
	    id_card_number.focus();
	    return false;
	}
	else if (cell_phone.value == "")
	{
	    alert("请输入手机号码");
	    cell_phone.focus();
	    return false;
	}
	else if (!checknum(cell_phone.value) || cell_phone.value.length < 10)
	{
	    alert("格式错误，请填写正确的手机号码");
	    cell_phone.focus();
	    return false;
	}
	else
	{
        if ("" != email.value)
	    {
	        if(g_obj_verify.isEmailAddressValid(email.value)!= g_obj_error_handle.success_code)
            {
                alert("格式错误，请填写正确的电子邮箱");
	            email.focus();
                return false;
            }
        }
        $("#exe").val("yes");
	    return true;
	}
	return false;
}

 
function checknum(num)
{
	var checknumber = "0123456789"
	for (i=0; i< num.length; i++)
	{
		for (j=0; j< checknumber.length; j++)
		{
			if (checknumber.charAt(j)==num.charAt(i))
			{
			   break;
			}
		}
		if (checknumber.length==j)
		{
			return false;
		}
	}
    return true;
}

function changeVIPABCPayMode(p_PayMode)
{
    var strPayMode = p_PayMode
    var datToday = new Date();

    //借記卡沒有有效期跟CVS2
    if ("50" == strPayMode)
    {
        $("#tr_exdate").hide();
        $("#tr_cvs2").hide();
    }
    else
    {
        $("#tr_exdate").show();
        $("#tr_cvs2").show();
    }

    //銀行下拉選單
    $.ajax({
        type: "POST",
        cache: false,
        url: "GetProductData.asp",
        data: "Type=9&SelectID=vipabc_ach_bank&PayMode=" + strPayMode,
        success: function (strProductOpt)
        {
            $("#span_vipabc_ach_bank").html(strProductOpt);
        },
        error: function () { alert("Type=9 Access Fail !"); }
    });
}

function print()
{
    $("#exe").val("");
    if (true == $.browser.msie && eval(parseInt($.browser.version, 10)) <= 8)
    {
        $("#div_ach_content").printElement({ printMode: 'popup' });
    }
    else
    {
        $("#div_ach_content").printElement();
    }
}
</script>
<div class="vip_pro">
    <div class="pay">
        <form name="fr" id="fr" method="post" action="https://<%=CONST_VIPABC_WEBHOST_NAME%>/program/member/CardExpireDateCheck.asp" lang="zh-cn"> <!--https-->
        <input type="hidden" name="int_contract_sn" value="<%=int_contract_sn%>" />
        <input type="hidden" name="int_client_sn" value="<%=int_client_sn%>" />
        <input type="hidden" name="exe" id="exe" value="<%=exe%>" />
        <div id="div_ach_content">
        <table width="100%" style="font-size: 14.0px;">
            <tr>
                <td align="center"><span style="font-size: 18.0px; font-family: 新細明體; font-weight: bolder;">个人授权协议书</span></td>
            </tr>
            <tr>
                <td align="right">编号：<%=int_contract_sn%></td>
            </tr>
            <tr>
                <td>甲方（付款人）：
                    <%if(exe = "yes") then%>
                    <%=strPayName%>
                    <%else%>
                    <input type="text" id="PayName" name="PayName" size="20" maxlength="10" value="<%=strPayName%>" />
                    <%end if%>
                </td>
            </tr>
            <tr>
                <td>乙方（收款商户）： <b><u>北京创意麦奇教育信息咨询有限公司</u></b></td>
            </tr>
            <tr>
                <td>丙方： <b><u>快钱支付清算信息有限公司</u></b></td>
            </tr>
            <tr>
                <td>
                    <br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;根据国家有关法律法规，就乙方通过第三方快钱支付清算信息有限公司(下称快钱公司)提供的委托代收服务向甲方收取咨询服务费 事宜（以下简称约定业务），甲乙双方经协商一致，订立本协议。<br/>
                    <br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;经甲方依本协议书的约定，授权由快钱公司根据乙方电子指令，代理乙方扣划甲方指定账户内的资金，并划入乙方账户。为执行上述事项，双方签订本协议，具体如下：<br/>
                </td>
            </tr>
            <tr>
                <td>
                    <div style="margin-left: 25px;"><br/>
                        <table>
                            <tr>
                                <td>甲方付款账户信息（以下简称指定账户）如下：</td>
                            </tr>
                            <tr>
                                <td>
                                    <%
                                    if(exe = "yes") then
                                        if("50" = strPayMethod) then
                                            response.write "借记卡代扣"
                                        else
                                            response.write "信用卡代扣"
                                        end if
                                    else
                                        Dim strTempPayModeSn : strTempPayModeSn = ""
                                        Dim strPayModeName : strPayModeName = ""
                                        Set objCfgPayMode = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
                                        strSql = " SELECT * FROM cfg_pay_mode WITH (NOLOCK) WHERE sn in (48,50) AND web_site = 'C' "
                                        intSqlWriteResult = objCfgPayMode.excuteSqlStatementEach(strSql, "", CONST_VIPABC_RW_CONN)
                                        while(not objCfgPayMode.eof)
                                            strChecked = ""
                                            strTempPayModeSn = objCfgPayMode("sn")
                                            if(sCstr(strTempPayModeSn) = sCstr(pmethod)) then
                                                strChecked = "checked"
                                            end if
                                            if("47" = strTempPayModeSn) then
                                                strPayModeName = "客戶每期登入官網刷卡"
                                            elseif("48" = strTempPayModeSn) then
                                                strPayModeName = "信用卡代扣"
                                            elseif("50" = strTempPayModeSn) then
                                                strPayModeName = "借记卡代扣"
                                            else
                                                strPayModeName = objCfgPayMode("data1")
                                            end if
                                    %>
                                    <input type="radio" name="installment_type_of_payment" value="<%=strTempPayModeSn%>" onclick="changeVIPABCPayMode(this.value);" <%=strChecked%>/><%=strPayModeName%>
                                    <%      objCfgPayMode.movenext
                                        wend
                                    end if
                                    %>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    开户行/机构：
                                    <%
                                    if(exe = "yes") then
                                        Set objCfgBankInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
                                        strSql = " SELECT id, BankCName FROM cfg_bank_info WITH (NOLOCK) WHERE id = @sn "
                                        intSqlWriteResult = objCfgBankInfo.excuteSqlStatementEach(strSql, Array(strBankSn), CONST_VIPABC_RW_CONN)
                                        if(not objCfgBankInfo.eof) then
                                            response.write objCfgBankInfo("BankCName")
                                        end if
                                    else
                                    %>
                                    <span id="span_vipabc_ach_bank">
                                        <select name="vipabc_ach_bank" id="vipabc_ach_bank">
                                            <option value="0">請選擇</option>
                                        </select>
                                    </span>
                                    <%end if%>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    户       名：
                                    <%if(exe = "yes") then%>
                                        <%=strAccountName%>
                                    <%else%>
                                    <input type="text" maxlength="10" id="account_name" name="account_name" />
                                    <%end if%>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    账号/卡号：
                                    <%if(exe = "yes") then%>
                                        <%=strCardNumber%>
                                    <%else%>
                                    <input type="text" size="26" maxlength="25" id="card_number" name="card_number" />
                                    <%end if%>
                                </td>
                            </tr>
                            <%if(not isEmptyOrNull(exmonth&exyear) or exe <> "yes") then%>
                            <tr id="tr_exdate">
                                <td>
                                    有效期：
                                    <%if(exe = "yes") then%>
                                        <%=exmonth & " 月" & exyear & " 年"%>
                                    <%else%>
                                    <input class="pay_input" type="text" name="exmonth" id="exmonth" value="" maxlength="2" size="5" /> 月
                                    <input class="pay_input" type="text" name="exyear" id="exyear" value="" maxlength="2" size="5" /> 年
                                    <div style="color: #999999; line-height: 15px;">请输入信用卡正面的有效期，如：09/13</div>
                                    <%end if%>
                                </td>
                            </tr>
                            <%end if%>
                            <%if(not isEmptyOrNull(cvs) or exe <> "yes") then%>
                            <tr id="tr_cvs2">
                                <td>
                                    <div class="pay_t">CVS2：
                                    <%if(exe = "yes") then%>
                                        <%=cvs%>
                                    <%else%>
                                        <input class="pay_input" type="text" name="cvs" id="cvs" value="" maxlength="3" size="5" />
                                        <%if(exe <> "yes") then%>
                                        <div class="cvs2">什么是CVS2?<div class="cvs2_pic"><img src="/images/card.jpg" /></div></div>
                                        <%end if%>  
                                    <%end if%>
                                    </div>
                                </td>
                            </tr> 
                            <%end if%>
                            <tr>
                                <td>
                                    身份证号码：
                                    <%if(exe = "yes") then%>
                                        <%=strIdCardNumber%>
                                    <%else%>
                                    <input type="text" size="21" maxlength="20" id="id_card_number" name="id_card_number" />
                                    <%end if%>  
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    付款人联系信息如下：
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    手机号码：
                                    <%if(exe = "yes") then%>
                                        <%=strCellPhone%>
                                    <%else%>
                                    <input type="text" size="12" maxlength="11" id="cell_phone" name="cell_phone" />
                                    <%end if%>  
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    电子邮箱(选填)：
                                    <%if(exe = "yes") then%>
                                        <%=strEmail%>
                                    <%else%>
                                    <input type="text" size="40" maxlength="40" id="email" name="email" />
                                    <%end if%>  
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <p>一、 甲方承诺：该付款账户仅限于甲方缴存应付给乙方的约定业务款项；该账户内资金均为甲方合法自有资金，不从事非法经营，不涉及洗钱，若因甲方违反相关法律法规给乙方、快钱公司造成损失的，应当承担损害赔偿责任。</p>
                    <p>甲方承诺甲方上述信息是真实、准确和有效的，甲方在上述信息发生变更时应及时通知乙方，否则因此而导致的所有风险和责任均由甲方自行承担。</p>
                    <p>二、甲方授权：以下均为有效的、不可撤销的授权。</p>
                    <p style="text-indent: -14.0pt; margin-left: 14.0pt;">
                    1. 甲方授权乙方根据业务往来需要将指定账户的资金扣收至乙方账户，快钱公司根据本协议代理乙方的扣划行为。
                    </p>
                    <p style="text-indent: -14.0pt; margin-left: 14.0pt;">
                    2. 如因账户状态异常（如账户挂失、冻结等）、账户余额不足、支付金额超过限额等原因导致交易无法完成的，甲方将及时办理换卡、挂失、重新开立账户/银行账户等手续，并及时通知乙方。因甲方原因导致交易失败或出现差错的，相应后果由甲方自行承担。
                    </p>
                    <p>三、本协议自签订之日起生效，一式二份，双方各执一份，具有同等法律效力。</p>
                    <p>四、甲方可通过以下方式解除本协议。</p>
                    <p style="text-indent: -14.0pt; margin-left: 14.0pt;">1. 甲方提前10个工作日以书面方式通知乙方（格式以附件《终止授权通知书》为准）。甲方上述书面通知到达乙方且各方办妥相关手续后，本协议失效。</p>
                    <p>五、本协议书适用中华人民共和国法律，有关本协议书的一切争议由乙方所在地人民法院管辖。</p>
                    <p>六、乙方已提请甲方注意对本协议各条款作全面、准确的理解，并应甲方的要求做了相应的条款说明。签约双方对本协议的含义认识一致。甲方愿意承担因其授权所产生的一切责任及风险。</p>
                    <p>七、本协议未尽事宜，由甲及乙方另行协商解决。</p>
                </td>
            </tr>
            <tr>
                <td><br/></td>
            </tr>
            <tr>
                <td>甲方:（中文亲签）</td>
            </tr>
            <tr>
                <td><br/></td>
            </tr>
            <tr>
                <td>乙方:<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;北京创意麦奇教育信息咨询有限公司</td>
            </tr>
            <tr>
                <td><br/></td>
            </tr>
            <tr>
                <td>签约日期：<%=getFormatDateTime(now(),5)%></td>
            </tr>
            <tr>
                <td>签约地点：<%=strClientAddress%></td>
            </tr>
            <tr>
                <td><br/></td>
            </tr>
            <%if(exe <> "yes") then%>
            <tr>
                <td>
                    <p>本会员同意授权VIPABC依照本授权协议书内容就本会员提供之银行账户/信用卡帐户进行按月扣款作业。本会员一旦按下「确认送出」之按键，即视为本会员已经同意完全遵守本个人授权协议书之所有约定。若会员未寄回个人授权协议书纸本，此约定仍视为有效。</p>
                </td>
            </tr>
            <%end if%>
            <%if(exe = "yes") then%>
            <tr>
                <td><p>本授权协议书业经会员确认无误，<b>敬请会员将本协议书列印并于确认栏位签字后寄回本公司(地址为:<u>上海市黄浦区西藏中路168号1802室</u>，收件人:<u>财务部ACH小组</u>)</b>，以利相关作业之进行。</p></td>
            </tr>
            <%end if%>
            <tr>
                <td align="center">
                <%if(exe = "yes") then%>
                <div>
                    <input type="button" value="打印" onclick="print();" class="pay_button" />&nbsp;
                    <input type="button" value="结束打印" onclick="window.location.href = 'http://<%=CONST_VIPABC_WEBHOST_NAME%>';" class="pay_button" />
                </div>
                <%else%>
                <div><input type="submit" value="确认送出" onclick="return go();" class="pay_button" /></div>
                <%end if%>
                </td>
            </tr>
        </table>
        </div>
        </form>
        <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
    </div>
</div>
<!--內容end-->