<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<% 
dim bolDebugMode : bolDebugMode = false 
dim bolTestCase : bolTestCase = true
%>
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

'20120316 阿捨新增 無限卡產品判斷	
Dim bolUnlimited : bolUnlimited = false
if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
    bolUnlimited = true 
end if

'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493
'20121112 阿捨新增 套餐產品改版
'改讀/program/member/reservation_class/getComboProductData.asp

'20121026 阿捨新增 員工免費課程 不可訂1對1
Dim bolStaffFreeClass : bolStaffFreeClass = false
if ( Instr(CONST_VIPABC_STAFF_FREE_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
    bolStaffFreeClass = true 
end if

'20121026 阿捨新增 公關產品 不可訂1對1
Dim bolPRFreeClass : bolPRFreeClass = false
if ( Instr(CONST_VIPABC_PR_FREE_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
    bolPRFreeClass = true 
end if

 '20120606 VM12053006 學習儲值卡官網+IMS功能 Penny - 學習卡無法訂一對一的課，跳alert訊息
 Dim bolLearningCard : bolLearningCard = false
if ( Not isEmptyOrNull(str_now_use_product_sn) ) then
    if ( Instr(CONST_VIPABC_NO_ONE_ON_ONE_SESSION_PRODUCT_SN, ("," & str_now_use_product_sn & ",")) > 0 ) then
        bolLearningCard = true 
    end if
end if

'20121112 阿捨新增 訂1對1權限 預設是可訂的
dim bolOneOnOne : bolOneOnOne = true
if ( true = bolUnlimited OR true = bolStaffFreeClass OR true = bolPRFreeClass ) then
    bolOneOnOne = false
end if

'課程類型
dim intClassType : intClassType = getRequest("clst", CONST_DECODE_NO)
    
if ( isEmptyOrNull(intClassType) ) then
    Response.Redirect "class_choose.asp"
    Response.End
else
    if ( true = bolDebugMode ) then
        response.Write "str_now_use_product_sn:" & str_now_use_product_sn & "<br />"
        response.Write "intClassType:" & intClassType & "<br />"
        response.Write "bolLearningCard:" & bolLearningCard & "<br />"
        response.Write "CONST_VIPABC_NO_ONE_ON_ONE_SESSION_PRODUCT_SN:" & CONST_VIPABC_NO_ONE_ON_ONE_SESSION_PRODUCT_SN & "<br />"
        response.End
    end if
    if ( sCInt(CONST_CLASS_TYPE_ONE_ON_THREE) = sCInt(intClassType) ) then
        if ( true = bolComboProduct ) then
            if ( true = bolComboHaving1On3 ) then
            else
                Call alertGo("提醒您：您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务","" ,CONST_ALERT_THEN_GO_BACK)
                Response.End
	        end if
        end if
        if ( true = bol1On3Session ) then
        else
            Response.Redirect "class_choose.asp"
            Response.End
        end if
    elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_ONE) = sCInt(intClassType) ) then
        '20120606 VM12053006 學習儲值卡官網+IMS功能 Penny - 學習卡無法訂一對一的課，跳alert訊息
	    if ( true = bolLearningCard ) then
            Call alertGo("提醒您：您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务","" ,CONST_ALERT_THEN_GO_BACK)
            Response.End
	    end if
		
	    '20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493
        if ( true = bolComboProduct ) then
	        if ( true = bolComboHaving1On1 ) then
            else
                Call alertGo("提醒您：您欲了解本项服务，请来电客服专线" & CONST_SERVICE_PHONE2 & "，由专人为您服务","" ,CONST_ALERT_THEN_GO_BACK)
                Response.End
	        end if
        end if
         '20120316 阿捨新增 無限卡產品判斷 不能上 ONE ON ONE
         '20121112 阿捨新增 無限卡, 員工免費課程, 公關產品 不可訂1對1
         if ( true = bolOneOnOne ) then
         else
            Response.Redirect "class_choose.asp"
            Response.End
         end if
    elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_FOUR) = sCInt(intClassType) ) then
        if ( true = bol1On4Session ) then
        else
            Response.Redirect "class_choose.asp"
            Response.End
        end if
    '大會堂轉址到 大會堂訂課頁面
    elseif ( sCInt(CONST_CLASS_TYPE_LOBBY) = sCint(intClassType) ) then
        Response.Redirect "lobby_near_order.asp"
        Response.End
    '不是class_choose.asp頁面轉址過來的
    elseif ( sCInt(CONST_CLASS_TYPE_ONE_ON_ONE) <> sCInt(intClassType) AND sCInt(CONST_CLASS_TYPE_ONE_ON_THREE) <> sCInt(intClassType) AND sCInt(CONST_CLASS_TYPE_ONE_ON_FOUR) <> sCInt(intClassType) ) then
        Response.Redirect "class_choose.asp"
        Response.End
    end if
end if
%>
<script type="text/javascript" src ="javascript/class.js"></script>
<script type="text/javascript">
<!--
var g_str_client_sn = "<%=getSession("client_sn", CONST_DECODE_NO)%>";
var g_str_now_use_product_sn = "<%=str_now_use_product_sn%>";

// 1 = 一般產品 2 = 定時定額產品
var g_str_product_type = "<%=int_client_purchase_product_type%>";

// product point_type
var g_str_product_point_type = "<%=int_product_point_type%>";

// class
var g_strClassType = "<%=getRequest("clst", CONST_DECODE_NO)%>";

var g_strServiceStart = "<%=Replace(str_service_start_date, "/", ".")%>";
//-->
</script>
<div class="main_membox">
    <!--說明 START-->

        <div class='page_title_1'><h2 class='page_title_h2'>预订课程</h2></div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 
    <div id="div_class_description" class="top_class6">
    </div>
    <!--說明 END-->
    <div class="book">
        <form id="frm_reservation" method="post" action="reservation_class_exe<%=strVIPABCJR%>.asp" onsubmit="return g_obj_class.submitForm();">
        <%'20130318 Create 加入傳入profileId和contractId 給 reservation_class_exe_VIPJR.asp %>
        <input type="hidden" name="profileId" value="<%=request("profileId")%>" />
        <input type="hidden" name="contractId" value="<%=request("contractId")%>" />
        <input type="hidden" name="clst" value="<%=strClassType%>" />
        <!--切換下個星期 START -->
        <div id="div_change_week" class="btop">
        </div>
        <!--切換下個星期 END -->
        <!--選擇訂課時間 START -->
        <div id="div_chose_time">
        </div>
        <!--選擇訂課時間 END -->
        <!--確定按鈕 START -->
        <div class="t-center m-top15">
            <%
            '檢查是否要顯示 如果沒有全時段就顯示 
                Dim int_run_result
                Dim obj_sql_opt_read : Set obj_sql_opt_read = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
                'client email
                Dim strClientEmail : strClientEmail = getSession("client_email", CONST_DECODE_NO)
                int_run_result = obj_sql_opt_read.excuteSqlStatementEach(" SELECT count(*) AS sum FROM lobby_dialoguebox WHERE valid=1 AND client_email='" & strClientEmail & "'",  "", CONST_TUTORABC_R_CONN)
                If ( int_run_result <> CONST_FUNC_EXE_SUCCESS ) then
                        'Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
                        'Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
                end if 
                if ( obj_sql_opt_read("sum") > 0 ) then
            %>
            <span style="color: #2a6bd1; font-size: 12px;">
                <input type="checkbox" id="open_remind" />&nbsp;我想了解同时段还有哪些大会堂课程</span>
            <%
                end if
            %>
            <input type="submit" id="btn_submit" value="确定订课" class="btn_1" /></div>
        <!--確定按鈕 END -->
        </form>
    </div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>