<!--#include virtual="/lib/include/global.inc" -->
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

Dim intClientSn : intClientSn = getRequest("client", CONST_DECODE_NO)
Dim intClassType : intClassType = getRequest("class_type", CONST_DECODE_NO)
'產品類型
Dim strProductType : strProductType = getRequest("product_type", CONST_DECODE_NO)
'cfg_product.point_type
if ( true = bolDebugMode ) then
    response.Write "strProductType: " & strProductType & "<br />"
end if

Dim strProductPointType : strProductPointType = getRequest("product_point_type", CONST_DECODE_NO)
'服務開始日
Dim strServiceStart : strServiceStart = getRequest("service_start", CONST_DECODE_NO)

Dim arrSessionPoints, intAvailableSession, intAvailablePoints, strClassDesc, intClassTypeTmp, flt_remain_week_session

Dim strClassLink : strClassLink = ""
	
dim strProductStatus : strProductStatus = ""

Dim strPurchaseAccountSn : strPurchaseAccountSn = str_purchase_account_sn '離目前時間最近的一筆 客戶於PURCHASE的序號
Dim strNowUseProductSn : strNowUseProductSn = str_now_use_product_sn '離目前時間最近的一筆 客戶合約的產品序號
'CS12041502-VIPABC訂課規則修改 20120515 tree--開始--
Dim intProductNewOrOld : intProductNewOrOld = 0 '取得產品是否為舊(1)還是新(2) ，VIPABC 舊產品客戶訂課時間為8小時前(1對1扣2堂) 新產品客戶訂課時間為24小時前(1對1扣3堂)
	
'20120723 阿捨新增 套餐產品堂數判斷
dim intMaxSession : intMaxSession = 0
dim intHavingSession : intHavingSession = 0
dim intComboGiftSession : intComboGiftSession = 0
'CS12041502-VIPABC訂課規則修改 20120515 tree--結束--

'20120316 阿捨新增 無限卡產品判斷
dim bolUnlimited : bolUnlimited = false
if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, ("," & strNowUseProductSn & ",")) > 0 ) then
    bolUnlimited = true 
end if

'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493
'20121112 阿捨新增 套餐產品改版
'改讀/program/member/reservation_class/getComboProductData.asp

'20121026 阿捨新增 員工免費課程 不可訂1對1
Dim bolStaffFreeClass : bolStaffFreeClass = false
if ( Instr(CONST_VIPABC_STAFF_FREE_PRODUCT_SN, (strNowUseProductSn & ",")) > 0 ) then
    bolStaffFreeClass = true 
end if

'20121026 阿捨新增 公關產品 不可訂1對1
Dim bolPRFreeClass : bolPRFreeClass = false
if ( Instr(CONST_VIPABC_PR_FREE_PRODUCT_SN, (strNowUseProductSn & ",")) > 0 ) then
    bolPRFreeClass = true 
end if

'CS12041502-VIPABC訂課規則修改 20120515 tree--開始--
intProductNewOrOld = getProductNewOrOld(strNowUseProductSn) '取得產品是否為舊(1)還是新(2) 產品編號小於等於433=舊
dim bolNewProduct : bolNewProduct = false
if ( true = bolDebugMode ) then
    response.Write "intProductNewOrOld: " & intProductNewOrOld & "<br />"
end if
if ( 2 = sCInt(intProductNewOrOld) ) then
    bolNewProduct = true
end if
'CS12041502-VIPABC訂課規則修改 20120515 tree--結束--

'20121112 阿捨新增 訂1對1權限 預設是可訂的
dim bolOneOnOne : bolOneOnOne = true
if ( true = bolUnlimited OR true = bolStaffFreeClass OR true = bolPRFreeClass ) then
    bolOneOnOne = false
end if

'課程類型描述
'20121109 阿捨新增 改讀課程類型資料庫
'Set objProductTypeData = getProductTypeData("", "C", 0, CONST_VIPABC_RW_CONN) '課程類型空值表讀出全部
'while not objProductTypeData.eof
'    intDataClassType = objProductTypeData("sst_number")
'    strDataClassName = objProductTypeData("sst_description_cn")
'    if ( sCInt(intDataClassType) = sCInt(intClassType) ) then
'        strClassDesc = strDataClassName & "课程"
'    else
'        strClassLink = strClassLink & " <a href='/program/member/reservation_class/reservation_class.asp?clst="& intDataClassType & "' >"&strDataClassName &"课程<a/> <br />"
'    end if
'    objProductTypeData.movenext
'wend

'todo 但錄影檔跟大會堂 與1對1, 1對3頁面沒有互通, 殘念不能讀db, 為了UI寫死以下設定
if ( not isEmptyOrNull(intClassType) ) then
    if ( sCInt(CONST_CLASS_TYPE_NORMAL) = sCInt(intClassType) ) then
        'strClassDesc = "1对6课程"
         strClassDesc = "小班制课程"
    elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_ONE) = sCInt(intClassType) ) then
        strClassDesc = "1对1课程"
        'strClassLink = " <a href='/program/member/reservation_class/reservation_class.asp?clst="& CONST_CLASS_TYPE_ONE_ON_THREE &"' >1对3课程<a/> "
        if ( true = bol1On4Session ) then
            strClassLink = " <a href='/program/member/reservation_class/reservation_class.asp?clst="& CONST_CLASS_TYPE_ONE_ON_FOUR &"' >小班制课程<a/> "
        else
            strClassLink = " <a href='/program/member/reservation_class/reservation_class.asp?clst="& CONST_CLASS_TYPE_ONE_ON_THREE &"' >小班制课程<a/> "
        end if
        '20121108 阿捨新增 套餐產品堂數判斷
        if ( true = bolComboProduct ) then
            'intMaxSession = getComboProductMaxSession(strNowUseProductSn, CONST_CLASS_TYPE_ONE_ON_THREE, 0, CONST_VIPABC_RW_CONN)
            if ( true = bolComboHaving1On3 or true = bolComboHaving1On4 ) then
            else
                strClassLink = ""
            end if
        else 
        end if
    elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_THREE) = sCInt(intClassType) ) then
        'strClassDesc = "1对3课程"
        strClassDesc = "小班制课程"
        strClassLink = " <a href='/program/member/reservation_class/reservation_class.asp?clst="& CONST_CLASS_TYPE_ONE_ON_ONE &"' >1对1课程<a/> "
        '20120723 阿捨新增 套餐產品堂數判斷
        if ( true = bolComboProduct ) then
            'intMaxSession = getComboProductMaxSession(strNowUseProductSn, CONST_CLASS_TYPE_ONE_ON_ONE, 0, CONST_VIPABC_RW_CONN)
            if ( true = bolComboHaving1On1 ) then
            else
                strClassLink = ""
            end if 
        elseif ( true = bolOneOnOne ) then
        else
            '20121112 阿捨新增 無限卡, 員工免費課程, 公關產品 不可訂1對1
            strClassLink = ""
        end if
    '20121206 阿捨新增 一對四
    elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_FOUR) = sCInt(intClassType) ) then
        strClassDesc = "小班制课程"
        strClassLink = " <a href='/program/member/reservation_class/reservation_class.asp?clst="& CONST_CLASS_TYPE_ONE_ON_ONE &"' >1对1课程<a/> "
         '20120723 阿捨新增 套餐產品堂數判斷
        if ( true = bolComboProduct ) then
            'intMaxSession = getComboProductMaxSession(strNowUseProductSn, CONST_CLASS_TYPE_ONE_ON_ONE, 0, CONST_VIPABC_RW_CONN)
            if ( true = bolComboHaving1On1 ) then
            else
                strClassLink = ""
            end if 
        elseif ( true = bolOneOnOne ) then
        else
            '20121112 阿捨新增 無限卡, 員工免費課程, 公關產品 不可訂1對1
            strClassLink = ""
        end if
    elseif ( sCInt(CONST_CLASS_TYPE_LOBBY) = intClassType ) then
        strClassDesc = "大会堂课程"
    end if
end if

if ( true = bolVIPABCJR ) then
    arrSessionPoints = getCustomerContractSessionPointInfoJR(intClientSn, "", "", CONST_VIPABC_RW_CONN)
else	
    arrSessionPoints =  getCustomerContractSessionPointInfo("", "", strPurchaseAccountSn, CONST_VIPABC_RW_CONN)
end if
intAvailableSession = 0
intAvailablePoints = 0
flt_remain_week_session = 0
if ( IsArray(arrSessionPoints) ) then
    if ( not isEmptyOrNull(arrSessionPoints(20)) ) then
        '可用堂數
        intAvailableSession = sCDbl(arrSessionPoints(20))
        '可用點數
        if (intAvailableSession > 0 ) then
            intAvailablePoints = intAvailableSession * CONST_SESSION_TRANSFORM_POINT
        end if
    end if

    '當週剩餘堂數
    if (Not isEmptyOrNull(arrSessionPoints(21))) then
        flt_remain_week_session = arrSessionPoints(21)
    end if
end if
%>
<!--說明start-->
<div>
<h1>
<%=getWord("WELCOME_JOIN_VIPABC") & "! " & getWord("YOU_CAN_RESERVATION_NOW") & "<font color='red'>" & Replace(strServiceStart, ".", "/") & "</font>" & getWord("AFTER_CONTRACT_START")%><!--欢迎您加入VIPABC行列! 您现在可预订YYYY/MM/DD(合约开始)之後的课程。-->
<br />
<%
'定時定額產品
if ( sCInt(CONST_PRODUCT_QUOTA) = sCInt(strProductType) ) then
%>
<!--当周期应用堂数<font color="red">&nbsp;<%=flt_remain_week_session%></font>-->
<%
else '非定時定額產品，顯示堂數資訊
    if ( true = bolUnlimited ) then

    else '非無限卡產品，顯示堂數資訊
        '20120723 阿捨新增 套餐產品堂數判斷
        if ( true = bolComboProduct ) then
            intHavingSession = getComboProductSession(intClientSn, str_now_use_contract_sn, intClassType, 1, 0, CONST_VIPABC_RW_CONN)
            intMaxSession = getComboProductMaxSession(strNowUseProductSn, intClassType, 0, CONST_VIPABC_RW_CONN)
            intComboGiftSession = getComboProductGift(str_now_use_contract_sn, "", intClassType, "", 0, CONST_TUTORABC_RW_CONN)
%>
可用<%=strClassDesc%>总堂数<font color="red">&nbsp;<%=sCDbl(intMaxSession) + sCDbl(intComboGiftSession)%></font>
<br />
剩余<%=strClassDesc%>堂数<font color="red">&nbsp;<%=sCDbl(intHavingSession) + sCDbl(intComboGiftSession)%></font>
<br />
<%
       else
%>
可用总堂数<font color="red">&nbsp;<%=sCDbl(intAvailableSession) + sCDbl(flt_remain_week_session)%></font>
<br />
剩余堂数<font color="red">&nbsp;<%=sCDbl(intAvailableSession)%></font>
<br />
<%
        end if
    end if 
end if 
%>
</h1>
<h2>
<img src="/images/images_cn/b_ps.gif" width="15" height="16" />您目前预订的是<span class="ps"><%=strClassDesc%></span>&nbsp;<%=strClassLink%>
<br />
</h2>
<br /><br />
<h1>
</h1>
<br class="clear">
<%'CS12041502-VIPABC訂課規則修改 20120515 tree--開始--
 if ( true = bolDebugMode ) then
    response.Write "bolNewProduct: " & bolNewProduct & "<br />"
end if
%>
<% if ( true = bolNewProduct ) then %><%'以下文字只有新產品看的見 
        if ( true = bolDebugMode ) then
            response.Write "CONST_PRODUCT_QUOTA: " & CONST_PRODUCT_QUOTA & "<br />"
            response.Write "strProductType: " & strProductType & "<br />"
        end if
%>
    <% if ( sCInt(CONST_PRODUCT_QUOTA) = sCInt(strProductType) ) then %><%'以下文字只有定期定額看的見 %>
        <% if ( sCInt(CONST_CLASS_TYPE_ONE_ON_ONE) = sCInt(intClassType) ) then '7、一對一 
                if ( true = bolDebugMode ) then
                    response.Write "bolOneOnOne: " & bolOneOnOne & "<br />"
                end if
            %>
            <% if ( true = bolOneOnOne ) then '20121112 阿捨新增 無限卡, 員工免費課程, 公關產品 不可訂1對1 %>
            <img src="/images/images_cn/ps_100719.gif" width="45" height="15" />&nbsp;&nbsp;预约/取消课程时请注意：<br />       
            1.新增预约课程:课程开始的24小时前完成，收取3课时数<br />
            2.取消预约课程:需在课程开始的4小时前完成<br />
        <%  end if 'if ( true = bolOneOnOne ) then
            end if 'if ( sCInt(CONST_CLASS_TYPE_ONE_ON_ONE) = sCInt(intClassType) ) then
        %>
        <% if ( sCInt(CONST_CLASS_TYPE_ONE_ON_THREE) = sCInt(intClassType) ) then '3、一對三%>
            <img src="/images/images_cn/ps_100719.gif" width="45" height="15" />&nbsp;&nbsp;预约/取消课程时请注意：<br />
            1.新增预约课程：课程开始的24小时前完成，扣除1课时数。<br />
            2.取消预约课程：动作需在课程开始的4小时前完成<br />
            3.临时预约24小时之内课程(区块内呈现橘色，名额有限)<br />
            &nbsp;&nbsp;(1)开课前24小时~开课前12小时预约：扣除1.25课时数。<br />
            &nbsp;&nbsp;(2)开课前12小时~开课前4小时预约：扣除1.5课时数。<br />
        <% end if %>
    <% elseif ( 1 = sCInt(strProductType) OR 3 = sCInt(strProductType) OR 4 = sCInt(strProductType) OR 5 = sCInt(strProductType) ) then %><%'以下文字只有一般產品(1)OR無限卡(3)OR終身無限卡(4)OR套餐產品(5)看的見 %>
        <% if ( sCInt(CONST_CLASS_TYPE_ONE_ON_THREE) = sCInt(intClassType) OR sCInt(CONST_CLASS_TYPE_ONE_ON_FOUR) = sCInt(intClassType) ) then '3、一對三%>
            <img src="/images/images_cn/ps_100719.gif" width="45" height="15" />&nbsp;&nbsp;预约/取消课程时请注意：<br />       
            1.新增订位:需在课程开始的24小时前完成新增订位作业。<br />
            2.取消订位:需在课程开始的4小时前完成取消订位作业。<br />
        <% end if %>
    <% end if %>
<% end if %>
<%'CS12041502-VIPABC訂課規則修改 20120515 tree--結束--%>
</div>
<!--說明end-->