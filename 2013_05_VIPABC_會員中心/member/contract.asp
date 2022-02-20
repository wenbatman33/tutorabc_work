<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
    '若已點過電子合約則導回首頁
'    if ( isEmptyOrNull(Session("contract_sn"))) then         
'        Call alertGo(getWord("contract_agree_3"), "/index.asp", CONST_ALERT_THEN_REDIRECT)
'        Response.End
'    end if
Dim intContractType : intContractType = getRequest("ContractType", CONST_DECODE_NO)	'1:一般短期合約 2: Learning Card
Dim intClientSn : intClientSn = getRequest("ClientSn", CONST_DECODE_NO)
Dim strRandomKey : strRandomKey = getRequest("RandomKey", CONST_DECODE_NO)
Dim strContractSn : strContractSn = getRequest("contract_sn", CONST_DECODE_NO)
%>
<script type="text/javascript">
//resizeTo(850,700);
var strContractType = "<%=intContractType%>";
var strClientSn = "<%=intClientSn%>";
var strRandomKey = "<%=strRandomKey%>";
var strContractSn = "<%=strContractSn%>";
$(document).ready(function(){
	//載入ajax ajax_contract.asp 將頁面載入
	if("1" == strContractType || "2" == strContractType)
	{
		$("#div_contact_info").load("ajax_short_contract.asp?ContractType=" + strContractType + "&ClientSn=" + strClientSn + "&RandomKey=" + strRandomKey);
	}
	else
	{
		$("#div_contact_info").load("ajax_contract.asp?contract_sn=" + strContractSn);
	}
});
</script>


 
<div class="temp_contant">
	<!--內容start-->
	<div class='page_title_5'><h2 class='page_title_h2'>课程购买合约</h2></div>
    <div class="box_shadow" style="background-position:-80px 0px;"></div> 
    <div id="div_contact_info"></div>
	<!--內容end-->
<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<font color="white"><%=trim(Request.ServerVariables("REMOTE_ADDR"))%></font>
<font color="white"><%=session("contract_sn")%></font>

