<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<% dim bolDebugMode : bolDebugMode = false %>
<!--#include virtual="/program/member/reservation_class/include/getComboProductData.asp"-->
<%
'會有六個session 五個變數 for 套餐判斷
'bolComboProduct, bolComboHaving1On1, bolComboHaving1On3, bolComboHavingLobby, bolComboHavingRecord
'20121206 阿捨新增 一對四判斷 bolComboHaving1On4
if ( true = bolDebugMode ) then
    response.Write "bolComboProduct: " & bolComboProduct & "<br />"
    response.Write "bolComboHaving1On1: " & bolComboHaving1On1 & "<br />"
    response.Write "bolComboHaving1On3: " & bolComboHaving1On3 & "<br />"
    response.Write "bolComboHaving1On4: " & bolComboHaving1On4 & "<br />"
    response.Write "bolComboHavingLobby: " & bolComboHavingLobby & "<br />"
    response.Write "bolComboHavingRecord: " & bolComboHavingRecord & "<br />"
end if
%>
<!--#include virtual="/program/member/reservation_class/include/getNewProductData.asp"-->
<% 
'會有2個session 2個變數 for 非套餐正常產品
if ( true = bolDebugMode ) then
    response.Write "bol1On3Session: " & bol1On3Session & "<br />"
    response.Write "bol1On4Session: " & bol1On4Session & "<br />"
end if
%>
<!--內容start-->
<style>
    body
    {
        margin: 0;
    }
    .wxts
    {
        width: 600px;
        height: 165px;
        background-image: url(/images/wxts.jpg);
        background-repeat: no-repeat;
    }
    .title
    {
        width: 95px;
        margin-left: 125px;
        margin-top: 30px;
        font-size: 12px;
        font-weight: bold;
        color: #686868;
        float: left;
    }
    .img
    {
        width: 205px;
        height: 79px;
        float: right;
        margin-right: 28px;
        margin-top: 40px;
    }
</style>
<% 
'20120316 阿捨新增 無限卡產品判斷	
dim bolUnlimited : bolUnlimited = false
if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
    bolUnlimited = true 
end if

'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493
'20121112 阿捨新增 套餐產品改版
'改讀/program/member/reservation_class/getComboProductData.asp

 '20120606 VM12053006 學習儲值卡官網+IMS功能 Penny - 學習卡無法訂一對一的課，跳alert訊息
 dim bolLearningCard : bolLearningCard = false
if ( Not isEmptyOrNull(str_now_use_product_sn) ) then
    if ( Instr(CONST_VIPABC_NO_ONE_ON_ONE_SESSION_PRODUCT_SN, ("," & str_now_use_product_sn & ",")) > 0 ) then
        bolLearningCard = true 
    end if
end if

'20121026 阿捨新增 員工免費課程 不可訂1對1
dim bolStaffFreeClass : bolStaffFreeClass = false
if ( Instr(CONST_VIPABC_STAFF_FREE_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
    bolStaffFreeClass = true 
end if

'20121026 阿捨新增 公關產品 不可訂1對1
dim bolPRFreeClass : bolPRFreeClass = false
if ( Instr(CONST_VIPABC_PR_FREE_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
    bolPRFreeClass = true 
end if

'20121112 阿捨新增 1對1按鈕3種隱藏方式判斷
dim bol1On1HideBtn : bol1On1HideBtn = false '整個隱藏
dim bol1On1GrayBtn : bol1On1GrayBtn = false '反灰顯示
dim bol1On1AlertBtn : bol1On1AlertBtn = false '整個顯示, 但點擊後直接alert
if ( true = bolUnlimited OR true = bolStaffFreeClass ) then '無限卡, 員工堂數 =>整個隱藏
    bol1On1HideBtn = true
end if

if ( (false = bolComboHaving1On1 AND true = bolComboProduct) OR true = bolPRFreeClass ) then '套餐產品1對1上限為0, 公關堂數, => 反灰顯示
    bol1On1GrayBtn = true
end if
if ( true = bolLearningCard ) then '學習儲值卡=>整個顯示, 但點擊後直接alert
    bol1On1AlertBtn = true
end if

if ( true = bolDebugMode ) then
    response.Write "bolPRFreeClass: " & bolPRFreeClass & "<br />"
    response.Write "bolComboHaving1On1: " & bolComboHaving1On1 & "<br />"
    response.Write "bol1On1GrayBtn: " & bol1On1GrayBtn & "<br />"
    response.Write "bolStaffFreeClass: " & bolStaffFreeClass & "<br />"
end if

if ( 578 = sCInt(str_now_use_product_sn) ) then
	bol1On1HideBtn = true
end if

if ( true = Session("test_staff") ) then
    bolShortSession = true
else
    bolShortSession = false
end if

'20121112 阿捨新增 套餐產品 所有課程類型判斷
'todo 沒有灰色圖所有先不用灰色按鈕判斷
'dim bol1On3GrayBtn : bol1On3GrayBtn = false '反灰顯示
'dim bolLobbyGrayBtn : bolLobbyGrayBtn = false '反灰顯示
'if ( true = bolComboHaving1On3 ) then '套餐產品1對3上限為0 => 反灰顯示
'else
'    bol1On3GrayBtn = true
'end if
'if ( true = bolComboHavingLobby ) then '套餐產品Lobby上限為0 => 反灰顯示
'else
'    bolLobbyGrayBtn = true
'end if

dim bol1On3HideBtn : bol1On3HideBtn = false '整個隱藏
dim bolLobbyHideBtn : bolLobbyHideBtn = false '整個隱藏
if ( true = bolComboProduct ) then
    if ( true = bolComboHaving1On3 ) then
    else
        bol1On3HideBtn = true
    end if
    if ( true = bolComboHavingLobby ) then
    else
        bolLobbyHideBtn = true
    end if
end if

'20121206 阿捨新增 一對三 or  一對四 判斷
'讀取/program/member/reservation_class/getNewProductData.asp
'表面寫小班制 實際為1對3 與 1對4
%>
<div class="main_membox">
    <div class="choose">
        <div class="w">请选择课程类型</div>
        <% 
            dim str1On1BtnUrl : str1On1BtnUrl = "" '1對1按鈕連結
            dim str1On1BtnHtml : str1On1BtnHtml = "" '1對1按鈕Html
            dim str1On3BtnUrl : str1On3BtnUrl = "" '1對3按鈕連結
            dim str1On3BtnHtml : str1On3BtnHtml = "" '1對3按鈕Html
            dim strLobbyBtnUrl : strLobbyBtnUrl = "" 'Lobby按鈕連結
            dim strLobbyBtnHtml : strLobbyBtnHtml = "" 'Lobby按鈕Html

            dim str1On4BtnUrl : str1On4BtnUrl = "" '1對4按鈕連結
            dim str1On4BtnHtml : str1On4BtnHtml = "" '1對4按鈕Html

            dim strShortLobbyBtnUrl : strShortLobbyBtnUrl = "" 'Lobby按鈕連結
            dim strShortLobbyBtnHtml : strShortLobbyBtnHtml = "" 'Lobby按鈕Html

            'todo 此地方不能讀資料庫 是因為非所有產品都有設定課程類型 目前只有套餐有
            str1On1BtnUrl = "<a href='reservation_class.asp?clst=" & CONST_CLASS_TYPE_ONE_ON_ONE & "' title='预订1对1课程' alt='预订1对1课程'></a>"
            'str1On3BtnUrl = "<a href='reservation_class.asp?clst=" & CONST_CLASS_TYPE_ONE_ON_THREE & "' title='预订1对3课程' alt='预订1对3课程'></a>"
            '20121206 阿捨新增 將所有1對n 改為 "小班制" 選項
            str1On3BtnUrl = "<a href='reservation_class.asp?clst=" & CONST_CLASS_TYPE_ONE_ON_THREE & "' title='预订小班制课程' alt='预订小班制课程'></a>"
            strLobbyBtnUrl = "<a href='lobby_near_order.asp' title='预订大会堂课程' alt='预订大会堂课程'></a>"
            
            '20121206 阿捨新增 一對三 or  一對四 判斷
            str1On4BtnUrl = "<a href='reservation_class.asp?clst=" & CONST_CLASS_TYPE_ONE_ON_FOUR & "' title='预订小班制课程' alt='预订小班制课程'></a>"
            
            strShortLobbyBtnUrl = "<a href='/program/aspx/AspxLogin.asp?url_id=7' target='_blank' title='预订10mins大会堂课程' alt='预订10mins大会堂课程'></a>"

            if ( true = bol1On3Session ) then
            else
                str1On3BtnUrl = ""
            end if

            if ( true = bol1On4Session ) then
            else
                str1On4BtnUrl = ""
            end if

            'bolShortSession = true
             if ( true = bolShortSession ) then
            else
                strShortLobbyBtnUrl = ""
            end if

            '20120316 阿捨新增 無限卡產品判斷 不能上 ONE ON ONE
            '20121112 阿捨新增 1對1按鈕兩種隱藏方式判斷
            if ( true = bol1On1HideBtn ) then
                str1On1BtnUrl = ""
            elseif ( true = bol1On1GrayBtn ) then
                str1On1BtnUrl = "<a href=""javascript:alert('您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务。');""><img src='/images/images_cn/order_1to1_grey.gif' /></a>"
            elseif ( true = bol1On1AlertBtn ) then
                str1On1BtnUrl = "<a href=""javascript:alert('您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务。');""></a>"
            end if 
            
            'todo 沒有灰色圖所有先不用灰色按鈕判斷
            'if ( true = bol1On3GrayBtn ) then
            '    str1On3BtnUrl = "<a href=""javascript:alert('您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务。');""><img src='/images/images_cn/order_1to1_grey.gif' /></a>"
            'end if
            if ( true = bol1On3HideBtn ) then
                str1On3BtnUrl = ""
            end if

            'todo 沒有灰色圖所有先不用灰色按鈕判斷
            'if ( true = bolLobbyGrayBtn ) then
            '    strLobbyBtnUrl = "<a href=""javascript:alert('您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务。');""><img src='/images/images_cn/order_1to1_grey.gif' /></a>"
            'end if
            if ( true = bolLobbyHideBtn ) then
                strLobbyBtnUrl = ""
            end if

            '當Url存在Html才包起它呈現
            if ( Not isEmptyOrNull(str1On1BtnUrl) ) then
                str1On1BtnHtml = "<div class='btn1'>" & str1On1BtnUrl & "</div>"
            else
                str1On1BtnHtml = ""
            end if

            if ( Not isEmptyOrNull(str1On3BtnUrl) ) then
                str1On3BtnHtml = "<div class='btn2'>" & str1On3BtnUrl & "</div>"
            else
                str1On3BtnHtml = ""
            end if

            '兩者不能並存 否則會出現兩個小班制
            if ( Not isEmptyOrNull(str1On4BtnUrl) AND isEmptyOrNull(str1On3BtnUrl) ) then
                str1On4BtnHtml = "<div class='btn2'>" & str1On4BtnUrl & "</div>"
            else
                str1On4BtnHtml = ""
            end if

            if ( Not isEmptyOrNull(strLobbyBtnUrl) ) then
                strLobbyBtnHtml = "<div class='btn3'>" & strLobbyBtnUrl & "</div>"
            else
                strLobbyBtnHtml = ""
            end if

            if ( Not isEmptyOrNull(strShortLobbyBtnUrl) ) then
                strShortLobbyBtnHtml = "<div class='btn4'>" & strShortLobbyBtnUrl & "</div>"
            else
                strShortLobbyBtnHtml = ""
            end if
        %>
        <%=str1On1BtnHtml%>
        <%=str1On3BtnHtml%>
        <%=str1On4BtnHtml%>
		<br class="clear" />
        <%=strLobbyBtnHtml%>
        <%=strShortLobbyBtnHtml%>
        <br class="clear" />
        <br />
        <br />
        <br />
        <br />
        <!--20130422 Beson 活動Banner下架-->
		<!--
		<div align="center">
	        <a href="http://www.vipabc.com/count.asp?code=JJ7a2k1MEs"><img src="/images/banner2_698-140.jpg" target="_blank" border="0"></a>
        </div>
		-->
		<!--
        <div align="center">
	        <a href="/program/mgm/LinkingPage/mgm_edm/index.asp"><img src="/images/mgm_edm_banner.jpg" border="0"></a>
        </div>
        <div align="center">
            <div class="wxts">
                <div class="title">你一定知道：</div>
                <div class="img"><a href="/program/mgm/LinkingPage/mgm_edm/index.asp"><img src="/images/img.jpg" border="0" /></a></div>
            </div>
            <div align="center"><a href="/program/member/reservation_class/lobby_near_order.asp"><img src="/images/BANNER_20120815.jpg" border="0" width="600" height="130" /></a></div>
        </div>
         -->
    </div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->