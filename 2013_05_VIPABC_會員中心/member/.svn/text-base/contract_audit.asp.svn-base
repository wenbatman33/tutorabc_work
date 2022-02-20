<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type="text/javascript">
<!--
/// <summary>
/// 20100709 阿捨新增 連結
/// 接收Flash選單值 改變連結
/// </summary>
/// <param name="text">按鈕值</param>
function GotoURL() {
    window.location="contract.asp"
}
//-->
</script>
<%
'若已點過電子合約則導回首頁
Dim str_BrushValue : str_BrushValue = getClientBrushCardSn(Session("contract_sn"),"contract")
if ( str_BrushValue <> 0 ) then 
	Session("contract_sn") =  str_BrushValue
end if 

if ( isEmptyOrNull(Session("contract_sn")) ) then
	 Call alertGo(getWord("contract_agree_3"), "/index.asp", CONST_ALERT_THEN_REDIRECT)
     Response.End
end if 

%>
<div class="temp_contant">
	<!--內容start-->
    <div class="main_membox">
    <form method="POST" action="contract_audit.asp" webbot-action="--WEBBOT-SELF--">
    <div class="pho" align="center" style="height:350px">
    	<!--明細start-->
        <!--
        <div id="url_button" style="top:150px;position:relative;">請使用flash player 10 以上版本</div>-->
        <div id="url_button" style="top:150px;position:relative;"><a href="contract.asp"><img src="/images/contract.png" /></a></div>
            <script type="text/javascript">
               /*
                var fo = new FlashObject("/flash/url_button_contract_audit.swf", "url_button", "315", "93", "7", "#000000");
                fo.addParam("wmode", "transparent");
                fo.addParam("allowFullscreen", "true");
                fo.addParam("bgcolor", "#ffffff");
                fo.addParam("quality", "high");
                fo.write("url_button");
               */            
            </script>
        <!--明細end-->            
    </form>
    </div>
	<!--內容end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>   