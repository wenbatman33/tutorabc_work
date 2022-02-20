<%
    '被 /include/html/template/template_innerpage_full.asp include
    '判斷 付費會員 和 一般會員 和 非會員的語法 在/lib/include/html/template/include/template_head_top.asp
%>
<!--header start-->
<div class="header">
<div class="logo" >
<%
dim str_logoFlash : str_logoFlash = ""
if (bol_homepage_index) then
	str_logoFlash = "logo.swf"
else
	str_logoFlash = "logo2.swf"
end if
%>
<div id="flash_logo" onclick="javascript:location.href='/index.asp';" style="cursor:pointer;padding-left:10px;">
	<img width="300" height="52" src="/images/logo_sub.png">
</div>
</div>
<div class="word1">
		<%if (isRegisterMember()) then %>
        Welcome, <%if (Not isEmptyOrNull(Session("client_fname"))) then Response.Write Session("client_fname") else Response.Write Session("client_cname") end if%> , <a href="/program/member/member.asp"><%=getWord("head_menu_1")%></a>
        <%if (checkReservationClassQualification_test(CONST_CLASS_TYPE_ONE_ON_ONE & "," & CONST_CLASS_TYPE_ONE_ON_THREE, CONST_VIPABC_RW_CONN)) then%>
        ‧<a href="/program/member/reservation_class/normal_and_vip.asp"><%=getWord("SHCEDULE_CLASS")%></a>
        <%end if%>
        ‧<a href="/program/member/member_logout_exe.asp"><%=getWord("head_menu_2")%></a>
        <%else%>
        <%=getWord("head_menu_3")%><a href="/program/member/member_login.asp"><%=getWord("head_menu_4")%></a>，<%=getWord("head_menu_5")%><a href="/program/regist/regist_add.asp"><%=getWord("REG_USER")%></a>
        <%end if%>
        <%if (Request.ServerVariables("PATH_INFO") <> "/index.asp") then %>‧<a href="/index.asp"><%=getWord("RETURN_INDEX")%></a><%end if %><!--回首页-->
		<a href="/program/faq/faq.asp"><img src="/images/images_cn/index/faq.gif" width="38" height="14" border="0"  title="<%=getWord("head_menu_6")%>"></a></div>

<!--menu start-->
<div class="menu" id="div_all_main_menu">
<div class="right"></div>
<%'if (Not isEmptyOrNull(getOnlineServicesGuid())) then%>
<!--<div style="z-index:2;">-->
    <!--<div class="btn7"><a id="a_menu_online_service" href="/program/service/open_onlive_service.asp"  target="_Blank" title="<%'=getWord("navigation_10")%>"></a></div>-->
    <!--<div id="div_sub_menu_online_service" class="menu2" style="width:80px; display:none;"></div>-->
<!--</div>-->
<%'end if%>

<!--搜索 START -->
<div style="z-index:2;">
    <div id="div_main_menu_search" class="btn6"><a id="a_menu_search" href="/program/search/index.asp" title="<%=getWord("SEARCH")%>"></a></div>
    <div id="div_sub_menu_search" class="menu2" style="width:60px; display:block; top:26px;"></div>
</div>
<!--搜索 END   -->

<!--新闻资讯 START -->
<div style="z-index:2;">
    <div id="div_main_menu_news" class="btn9"><a id="a_menu_news" href="/program/news/news.asp" title="新闻资讯"></a></div>
    <div id="div_sub_menu_news" class="menu2" style="width:80px; display:block; top:26px;"></div>
</div>
<!--新闻资讯 END   -->

<!--特惠活動 START -->
<div style="z-index:2;">
    <div id="div_main_menu_event" class="btn5"><a id="a_menu_event" href="/program/member/reservation_class/expo_reservation.asp" title="<%=getWord("head_menu_7")%>"></a></div>
    <div id="div_sub_menu_event" class="menu2" style="width:80px; display:block; top:26px;"></div>
</div>
<!--特惠活動 END   -->

<!--成果分享 START -->
<div style="z-index:2;">
    <div id="div_main_menu_achievements" class="btn4"><a id="a_menu_achievements" href="/program/mgm/index.asp" title="<%=getWord("head_menu_8")%>"></a></div>
    <div id="div_sub_menu_achievements" class="menu2" style="width:80px; display:block; top:26px;"></div>
</div>
<!--成果分享 END   -->

<%if (isRegisterMember()) then %>
<!--會員專區 START -->
<div style="z-index:2;">
    <div id="div_main_menu_member" class="btn3"><a id="a_menu_member" href="/program/member/member.asp" title="<%=getWord("CLASS_VOCABULARY_4")%>"></a></div>
    <div id="div_sub_menu_member" class="menu2" style="width:100px; display:block; top:26px;"></div>
</div>
<!--會員專區 END   -->
<%end if%>

<!--訂做專屬課程 START -->
<div style="z-index:2;">
    <div id="div_main_menu_classroom" class="btn2" ><a id="a_menu_classroom" href="/program/classroom/demo_order.asp" title="<%=getWord("head_menu_9")%>"></a></div>
    <div id="div_sub_menu_classroom" class="menu2" style="width:140px; display:block; top:26px;"></div>
</div>
<!--訂做專屬課程 END -->

<!--專業顧問團隊 START -->
<div style="z-index:2;">
    <div id="div_main_menu_consultants" class="btn1"  ><a id="a_menu_consultants" href="/program/consultant/consultant_list.asp" title="<%=getWord("head_menu_10")%>"></a></div>
    <div id="div_sub_menu_consultants" class="menu2" style="width:100px; display:block; top:26px;"></div>
</div>
<!--專業顧問團隊 END   -->

<!--产品特色 END   -->
<div style="z-index:2;">
    <div id="div_main_menu_product_feature" class="btn8"><a id="a_menu_product_feature" href="/program/product/product.asp" title="<%=getWord("MENU_PRODUCT_FEATURE")%>"></a></div>
    <div id="div_sub_menu_product_feature" class="menu2" style="width:80px; display:block; top:26px;"></div>
</div>
<!--产品特色 END   -->

<div id="div_main_start_picture" class="left"></div>
<div class="clear"></div>
</div>
<!--menu end-->
</div>

<!--header end-->
<!--header 2010.01.05-->