<!--#include virtual="/lib/include/html/template/template_change_jr.asp"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<link type="text/css" href="/lib/javascript/JQuery/css/redmond/jquery-ui-1.8.2.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
    $(document).ready(function () {
        //設定dialog大小
        $("#dialog").dialog("destroy");
        $("#dialog-form-helper").dialog({
            autoOpen: false,
            height: 350,
            width: 550,
            modal: true,
            close: function () {
                $("#dialog-form-helper").html("");
            }
        });
    });

    /// <summary>
    /// 20100902 阿捨新增 小幫手頁面
    /// 呼叫ajax載入iframe (為了符合css)
    /// </summary>
    function ShowDialog() {
        //dialog載入
        $("#dialog-form-helper").html("<iframe src='/program/include/realtime_message_JR.asp?client_sn=<%=g_var_client_sn%>' width='100%' height='100%' scrolling='no' frameborder='0'></iframe>");

        //開啟dialog
        $("#dialog-form-helper").dialog("open");
        $("#dialog-form-helper").dialog("option", "resizable", false);
    }
</script>

<div id="dialog-form-helper" title="问题小帮手" /></div>

<% 
dim bolDebugMode : bolDebugMode = false
dim bolTestCase : bolTestCase = false 

dim intClientSn : intClientSn = int_member_client_sn
dim strClientEmail : strClientEmail = session("client_email")
dim intUrlId : intUrlId = sCint(getRequest("url_id", CONST_DECODE_NO))
dim bolIframe : bolIframe = true

select case (intUrlId)
    case 0
        Response.Redirect "/index.asp"
        Response.End
    case 2 '帳戶資訊
        'function setNewMemberUsedPoints(ByVal p_intContractSn, ByVal p_intClientSn, ByVal p_intDebugMode, ByVal p_intConnectType)
        '/****************************************************************************************
        '描述：設定實際使用堂數 for member.contract
        '傳入參數：
        '[p_intContractSn:Integer] 合約編號
        '[p_intClientSn:Integer] 客戶編號
        '[p_intDebugMode] 除錯模式 0: 關閉 1:開啟
        '[p_intConnectType:Integer] 連線的主機型態
        '回傳值 ： ture or false
        '牽涉數據表：member.Contract
        '備註 ：
        '歷程(作者/日期/原因) ：[ArTHurWu] 2013/06/04 Created
        '*****************************************************************************************/
        Call setNewMemberUsedPoints(str_now_use_contract_sn, intClientSn, 0, CONST_VIPABC_RW_CONN)
    case 7
        strAspxUrl = "/aspx/ShortSession/Home/Index"
        bolIframe = false
        if ( true = bolTestCase ) then
           'strAspxUrl = "http://192.168.23.124:9001/aspx/ShortSession/Home/Index"
            strAspxUrl = "http://192.168.23.166/aspx/ShortSession/"
        end if
end select

if ( true = bolDebugMode ) then
    response.Write "intUrlId: " & intUrlId & "<br/>"
    response.Write "strAspxUrl: " & strAspxUrl & "<br/>"
    response.end
end if

if ( true = bolIframe ) then
%>
<!-- Menber Content Start -->
<div class="main_con"  align="left">
	<% if ( not isEmptyOrNull(intClientSn) ) then %>
        <iframe src ="/program/aspx/iframeLogin.asp?url_id=<%=intUrlId%>" width="720" height="1800" frameborder="0"></iframe>
	        <%
	else
		Response.Redirect "/program/member/member_login.asp"
	end if
	%>
	 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!-- Main Content End -->
<% else %>
<html>
<table width="100%" align="left"><img id="img_ajax_loader" src="/lib/javascript/JQuery/images/ajax-loader.gif" border="0" />
<form name="asp_jr" id="asp_jr" action="<%=strAspxUrl%>" method="POST">
<input type="hidden" name="ClientSN" value="<%=intClientSn%>" />
<input type="hidden" name="client_sn" value="<%=intClientSn%>" />
<input type="hidden" name="email" value="<%=strClientEmail%>" />
</form>
<script type="text/javascript">asp_jr.submit();</script>
</table>
</html>
<% end if %>