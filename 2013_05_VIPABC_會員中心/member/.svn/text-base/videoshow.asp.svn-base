<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<% dim bolDebugMode : bolDebugMode = false %>
<!--#include virtual="/program/member/reservation_class/include/getComboProductData.asp"-->
<%
'會有六個session 五個變數
'bolComboProduct, bolComboHaving1On1, bolComboHaving1On3, bolComboHavingLobby, bolComboHavingRecord
'20121112 阿捨新增 套餐產品改版 沒有大會堂上限 導回課程選擇頁面
if ( true = bolComboProduct ) then
    if ( true = bolComboHavingRecord ) then
    else
        Call alertGo("提醒您：您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务","" ,CONST_ALERT_THEN_GO_BACK)
        Response.End
    end if
end if
%>
<!-- Menber Content Start -->

<div class="main_con"  align="center">
	<% if (not isEmptyOrNull(session("client_sn"))) then%>
		<iframe src ="videoshow_authenticate.asp" width="650" height="1000" frameborder=0></iframe>
	<%
	else
		Response.Redirect "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/member/member_login.asp"
	end if
	%>
	 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!-- Main Content End -->

