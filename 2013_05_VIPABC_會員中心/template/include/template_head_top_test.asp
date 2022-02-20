<!--#include virtual="/text/Alt_cfg.asp"-->
<head>
<link rel="canonical" href="http://www.vipabc.com<%=Request.ServerVariables("PATH_INFO")%>?<%=Request.ServerVariables("QUERY_STRING")%>" />
<meta http-equiv="Content-Type" content="text/html; charset=<%=CONST_WEBSITE_CHARSET%>" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta http-equiv="X-UA-Compatible" content="IE=8">
<%
dim siteurl : siteurl = Request.ServerVariables("PATH_INFO")	
dim defaultpagetitle : defaultpagetitle = "VIPABC_全球最大的真人在线英语外教网|商务英语培训|英语口语培训|英语听力培训"
dim defaultpagedescription : defaultpagedescription = "VIPABC真人在线英语外教网为全球最大在线语言培训业者,自行研发DCGS动态课程学习系统,以学习者为中心,和全球外教同步视频实时互动、提供真人在线, 外教一对一VIP级量身订作的英语培训平台,在24小时不间断的全英语学习环境中,透过1~3人小班制及精准配对的个性化课程提升商务英语口语能力,随时随地学英语,全球老外到你家。4006-30-30-30"
dim defaultpagekeywords : defaultpagekeywords = "全球外教,英语培训,英语培训机构,商务英语,英语口语,外教一对一,在线英语,在线英语培训,成人英语,英语外教,英语学习,英语网校,英语会话,英语口语培训,在线学英语,网上学英语,网络学英语,外教口语,学英语,英语听力,网络英语,学英文,英语教育,电话英语,英语家教,托福雅思英语,VIPABC"
dim strFrameUrl : strFrameUrl = Request.ServerVariables("script_name")
%>
<!--lewisli 2013/03/27 start-->
<%
if ( siteurl = "/program/news/news_detail.asp") then
	dim intNewsId : intNewsId = request("NewsId")	
%>
<!--#include virtual="/backoffice/inc/conn.asp"-->
<%
dim objNewTitle : objNewTitle = null
dim objTitlesql : objTitlesql = ""
set objNewTitle = server.CreateObject("adodb.recordset")
strSql = " SELECT NewsId, NewsTitle, NewsContent, NewsCategory FROM News WHERE (NewsId = "& intNewsId &") AND (NewsStatus = 1) "
objNewTitle.open strSql,conn,1,1
defaultpagetitle = objNewTitle("NewsTitle") & "| VIPABC真人在线_实时互动_同步视频"
	objNewTitle.close
    set objNewTitle=nothing
end if
%>
<!--lewisli 2013/03/27 end-->
<title><%=defaultpagetitle%></title>
<%
if ( siteurl = "/program/news/news_detail.asp") then
else
%>
<meta name="description" content="<%=defaultpagedescription%>" />
<meta name="keywords" content="<%=defaultpagekeywords%>" />
<%
end if
'定義變數
Dim str_homepage_template
Dim str_flash_index_main_banner
Dim str_flash_content_flash
Dim bol_homepage_index

if (Request.ServerVariables("script_name") = "/index.asp" or Request.ServerVariables("script_name") = "/index_test.asp") then '判斷是否為首頁
	bol_homepage_index = true '確認為首頁參數		
else
	bol_homepage_index = false
end if
'20120320 阿捨新增 VIPABCJR客戶判斷 bolVIPABCJR 
%>
<!--#include virtual="/program/member/getJRClient.asp"-->
<meta http-equiv="content-language" content="zh-cn" />
<meta name="Contact" content="电话:4006-303030" />
<meta name="Copyright" content="麦奇教育集团VIPABC Inc." />
<meta name="Author" content="VIPABC英语" />
<meta name="Generator" content="VIPABC英语" />
<script type="text/javascript" src="/lib/javascript/swfobject.js"></script>
<script type="text/javascript" src="/lib/javascript/AC_RunActiveContent.js"></script>
<script type="text/javascript" src="/lib/javascript/amplify_prototype.js"></script>
<script type="text/javascript" src="/lib/javascript/verify.js"></script>
<script type="text/javascript" src="/lib/javascript/error_handle.js"></script>
<% if ( true = bolVIPABCJR AND false = bol_homepage_index ) then %>
<script type="text/javascript" src="/lib/javascript/jquery-1.6.2.min.js"></script>
<% else %>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery.js"></script>
<% end if %>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/plugin/jquery.textarea-expander.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/plugin/jquery.tablesorter.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/plugin/jquery.watermarkinput.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/plugin/ui.datepicker.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/plugin/jquery.blockUI.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/ui/ui.core.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/ui/ui.tabs.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/ui/jquery.ui.position.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery.printElement.js"></script>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/plugin/jquery.cookie.js"></script>
<link rel="stylesheet" href="/lib/javascript/JQuery/css/ui.datepicker.css" type="text/css" media="screen" title="no title" />
<link rel="shortcut icon" type="image/ico" href="/favicon.ico">
<!-- VM12061107 Hunter：安裝百度代碼 -->
<script type="text/javascript" src="/lib/javascript/baidu_analytics.js"></script>
<!--<script type="text/javascript" src="/lib/javascript/google_analytics.js"></script>-->
<script type="text/javascript">
<!--
var _gaq = _gaq || [];
_gaq.push(['_setAccount', 'UA-26071606-2']);
_gaq.push(['_trackPageview']);
(function () {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
})();
//-->
</script>
<!-- Gridsum tracking code begin. -->
<script type='text/javascript'>
<!--
(function () {
    var s = document.createElement('script');
    s.type = 'text/javascript';
    s.async = true;
    s.src = (location.protocol == 'https:' ? 'https://ssl.' : 'http://static.') + 'gridsumdissector.com/js/Clients/GWD-000718-73B5CF/gs.js';
    var firstScript = document.getElementsByTagName('script')[0];
    firstScript.parentNode.insertBefore(s, firstScript);
})();
//-->
</script>
<!--Gridsum tracking code end. -->
<%
'輸出Javascript Code For 設定子選單的Css Class Name用
'Call responseJsCodeForSubMenuCssClass()
'Call BlackIPDelayFrame()
if ( Session("client_sn") = "263739" OR Request.Cookies("stupid_spy") = "faking" ) then
    'dim strNowUrl : strNowUrl = Request.ServerVariables("script_name")
    dim strStop : strStop = getRequest("stop", CONST_DECODE_ESCAPE)
    if ( isEmptyOrNull(strStop) ) then
	    call delayFrame(90, "/program/fix_incident.asp")
        response.end
    end if
end if
%>
<% if (Not isEmptyOrNull(getOnlineServicesGuid())) then %>
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>
<% end if %>
<!--<script type="text/javascript" src="/lib/javascript/menu.js"></script>-->
<script type="text/javascript">
<!--
$(document).ready(function () {
    $("#<%=CONST_HTML_MAIN_CONTINENT_DIV_ID%>").append($("#<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>").parent());
});

function selectHomepageTemplate(p_str_template) {
    //JQuery AJAX
    var str_postdata = "";
    str_postdata = "str_template=" + p_str_template;
    $.ajax({
        type: "POST",
        cache: false,
        url: "/lib/include/html/ajax_select_homepage_template.asp",
        data: str_postdata,
        success: function (msg) {
            window.top.location.reload();
        },
        error: function () {
            //alert("Ajax Error!");
        }
    });
}
//-->
</script>
<% if ( true = bol_homepage_index ) then %>
<link href="/css/index_test.css" rel="stylesheet" type="text/css" />
<% else %>
<link media="screen" type="text/css" href="/css<%=strVIPABCJR%>/css_cn_test.css" rel="stylesheet" />
<% end if %>
<!--右側AD START-->
<!--#include virtual="/program/include/floating_banner.asp" -->
<!--include virtual="/program/include/floating_banner_left.asp" -->
<!--右側AD END-->