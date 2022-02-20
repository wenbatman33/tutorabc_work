<!--#include virtual="/lib/include/html/template/include/template_html_top.asp"-->
<!--#include virtual="/lib/include/html/template/include/template_head_top_test.asp"-->
<!--#include virtual="/lib/include/html/template/include/template_head_bottom.asp"-->
<% 
'20121108 阿捨新增 合約頁面無法複製
dim strThisPageURL : strThisPageURL = Request.ServerVariables("script_name")
dim bolContractPage : bolContractPage = false
if ( "/program/member/contract.asp" = strThisPageURL ) then
    bolContractPage = true
end if
%>
<% if ( true = bolContractPage ) then %>
<body ONDRAGSTART="window.event.returnValue=false" ONCONTEXTMENU="window.event.returnValue=false" onSelectStart="event.returnValue=false">
<% else %>
<body>
<% end if %>
<div id="all">
<!--#include virtual="/lib/include/html/template/include/template_body_embed_top.asp"-->
<!--include virtual="/lib/include/html/template/include/include_head_menu.asp"-->
<!--#include virtual="/lib/include/html/menu.asp" -->
<!--include virtual="/lib/include/html/template/include/include_head_flash.asp"-->
<div class="clear"></div>
<div id="main">
<div class="all_width" >
    <!-- Menber Content Start -->
    <div class="main">
        <div id="<%=CONST_HTML_MAIN_CONTINENT_DIV_ID%>" class="main_con">
           <!--#include virtual="/lib/include/html/template/include/include_path.asp"-->
           <!--contents start-->       
           <!--contents end-->
        </div>
        <!--#include virtual="/lib/include/html/template/include/include_sidebar.asp"-->
    <div class="clear"></div>
</div>
<!-- Main Content End -->
<!-- include virtual="/lib/include/html/template/include/include_navigation.asp"-->
</div>
</div>
<!--#include virtual="/lib/include/html/template/include/include_footer.asp"-->
<!--#include virtual="/lib/include/html/template/include/template_body_embed_bottom.asp"-->
</div>
</body>
<!--#include virtual="/lib/include/html/template/include/template_html_bottom.asp"-->