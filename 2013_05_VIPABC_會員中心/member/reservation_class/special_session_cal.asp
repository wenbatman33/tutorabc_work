<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<% 
dim bolDebugMode : bolDebugMode = false '正式改為false
dim bolTestCase : bolTestCase = false

Dim str_search_topic : str_search_topic = trim(unEscape(Request("lobby_topic"))) '搜尋條件: 主題
Dim str_search_con : str_search_con = trim(Request("con_sn")) '搜尋條件: 顧問編號
Dim str_search_lev : str_search_lev = trim(Request("lobby_suit_lev")) '搜尋條件: 適合等級
dim strCheckDivOpen : strCheckDivOpen = "" '預設值
dim strReservedCalOrList : strReservedCalOrList = "cal" '確認訂課後 判斷載入頁面

'傳cal分頁
dim intNowPage : intNowPage = getRequest("page", CONST_DECODE_NO)
dim strNowDate : strNowDate = ""
dim strMaxNowDate : strMaxNowDate = ""

if ( isEmptyOrNull(intNowPage) ) then '一開始分頁為0
    intNowPage = 0
end if

strNowDate = DateAdd("d", intNowPage*3, date()) '計算頁數顯示之日期Min
strMaxNowDate = DateAdd("d", 3, strNowDate) '計算頁數顯示之日期Max

if ( true = bolDebugMode )  then
    response.write "intNowPage : " & intNowPage &"<br/>"
    response.write "strNowDate : " & strNowDate &"<br/>"
    response.write "strMaxNowDate : " & strMaxNowDate &"<br/>"
    response.write "strCheckDivOpen : " & strCheckDivOpen &"<br/>"
end if
%>
<!--內容開始-->
<div class="oneColFixCtr">
    <div id="container">
        <div id="mainContent">
            <div id="session_booking">
                <!-- end .session_booking_head -->
                <div class="session_booking_body">
                    <!-- end .head_bar -->
                    <div class="session_booking_cal">
                        <div class="session_booking_cal_section">
                            <!--時段區塊-->
                            <!--早 展開 傳入參數：開始時間-1；結束時間；今日日期；今日日期+3；載入區塊；區間開關；開合圖示；分頁-->
                            <div class="session_booking_cal_section_h">
                                <a href="javascript:;" onclick="showTable('4','11','<%=strNowDate%>', '<%=strMaxNowDate%>','morning','<%=intNowPage%>', '<%=str_search_topic%>','<%=str_search_con%>','<%=str_search_lev%>');alertTime('morning');">
                                    <img class="section_ico open" id="morning1" src="/images/SpecialSession/icon_spacer.gif" />
                                    <span id="morning5">上午时段 05:30~12:00</span></a><!--img class中的open和close控制區塊展開關閉的樣式-->
                                <div class="session_booking_cal_section_b" id="morning2">
                                </div>
                                <!--開啟的div位置-->
                                <input type="hidden" id="morning3" name="morning3" value="close" /><!--區間開關變數-->
                                <input type="hidden" id="morning4" name="morning4" value="" /><!--判斷查詢資料是否存在-->
                            </div>
                            <!--end .session_booking_cal_section_h -->
                            <!--下 展開 傳入參數：開始時間-1；結束時間；今日日期；今日日期+3；載入區塊；區間開關；開合圖示；分頁-->
                            <div class="session_booking_cal_section_h">
                                <a href="javascript:;" onclick="showTable('11','17','<%=strNowDate%>', '<%=strMaxNowDate%>','noon','<%=intNowPage%>', '<%=str_search_topic%>','<%=str_search_con%>','<%=str_search_lev%>');alertTime('noon');">
                                    <img class="section_ico close" id="noon1" src="/images/SpecialSession/icon_spacer.gif" />
                                    <span id="noon5">下午时段 12:30~18:00</span></a><!--img class中的open和close控制區塊展開關閉的樣式-->
                                <div class="session_booking_cal_section_b" id="noon2">
                                </div>
                                <!--開啟的div位置-->
                                <input type="hidden" id="noon3" name="noon3" value="close" /><!--區間開關變數-->
                                <input type="hidden" id="noon4" name="noon4" value="" /><!--判斷查詢資料是否存在-->
                            </div>
                            <!--晚 展開 傳入參數：開始時間-1；結束時間；今日日期；今日日期+3；載入區塊；區間開關；開合圖示；分頁-->
                            <div class="session_booking_cal_section_h">
                                <a href="javascript:;" onclick="showTable('17','23','<%=strNowDate%>', '<%=strMaxNowDate%>','night','<%=intNowPage%>', '<%=str_search_topic%>','<%=str_search_con%>','<%=str_search_lev%>');alertTime('night');">
                                    <img class="section_ico close" id="night1" src="/images/SpecialSession/icon_spacer.gif" />
                                    <span id="night5">晚间时段 18:30~23:30</span></a><!--img class中的open和close控制區塊展開關閉的樣式-->
                                <div class="session_booking_cal_section_b" id="night2">
                                </div>
                                <!--開啟的div位置-->
                                <input type="hidden" id="night3" name="night3" value="close" /><!--區間開關變數-->
                                <input type="hidden" id="night4" name="night4" value="" /><!--判斷查詢資料是否存在-->
                            </div>
                        </div>
                        <!--end .session_booking_cal_section-->
                    </div>
                    <!--end .session_booking_cal-->
                </div>
                <!-- end .session_booking_body -->
            </div>
            <!--end .session_booking-->
        </div>
        <!-- end .mainContent -->
    </div>
    <!-- end .container -->
</div>
<!--end .oneColFixCtr-->
<!--內容結束-->
<!--燈箱內容開始-->
<!--移至最外層-->
<!--燈箱內容結束-->
<input type="hidden" id="session_sn" value="" />
<input type="hidden" id="session_status" value="" />
<input type="hidden" id="session_sdate" value="" />
<input type="hidden" id="session_stime" value="" />
<input type="hidden" id="client_attend_sn" value="" />
<span id="check_box" style="display: none;"></span>
<script type="text/javascript" language="javascript">
<!--
var strTime_morning = ""; //紀錄時段
var strTime_noon = "";
var strTime_night = "";
function showTable(intMinTime, intMaxTime, strMinDate, strMaxDate, strCheckDivOpen, intNowPage, intLobbyTopic, intConsultantSn, intLobbySuitLevel ) 
{
    //document.body.style.cursor="progress";
    var strCheckValue = $("#" + strCheckDivOpen + "3").val(); //區間開關變數
    if ( strCheckValue == "open" ) //區間原本為開 click後觸發
    {
        $("#" + strCheckDivOpen + "3").attr("value", "close"); //區間關
        $("#" + strCheckDivOpen + "1").attr("class", "section_ico close"); //開合圖示變化，若關閉則圖示(+)(close)

        //TODO : 若session其中三個時段的其中兩個為false 則不能清空

        $("#" + strCheckDivOpen + "2").html(""); //div清空
        $.ajax({
            url: "/program/member/reservation_class/clear_session.asp",
            cache: false,
            async: false,
            data: {
                strCheckDivOpen : strCheckDivOpen
            },
            success: function (result) {
            },
            error: function (result) {
                //alert("error:" + result);
            }
        });
        //
        if ( "morning" == strCheckDivOpen )
        {
            strTime_morning = "";
        }
        if ( "noon" == strCheckDivOpen )
        {
            strTime_noon = "";
        }
        if ( "night" == strCheckDivOpen )
        {
            strTime_night = "";
        }
    }
    else if (strCheckValue == "close") //區間原本為關 click後觸發
    {
        $("#" + strCheckDivOpen + "3").attr("value", "open"); //區間開
        $("#" + strCheckDivOpen + "1").attr("class", "section_ico open"); //開合圖示變化，若打開則圖示(-)(open)
        $("#" + strCheckDivOpen + "2").html("<img id='img_ajax_loader' src='/lib/javascript/JQuery/images/ajax-loader.gif' border='0' />"); //loading圖示
        //打開內容
        $.ajax({
            url: "/program/member/reservation_class/special_session_cal_content<%=strVIPABCJR%>.asp",
            cache: false,
            async: false,
            data: {
                Min_Time : intMinTime,
                Max_Time : intMaxTime,
                Min_Date : strMinDate,
                Max_Date : strMaxDate,
                intNowPage: intNowPage,
                strCheckDivOpen: strCheckDivOpen,
                lobby_suit_lev : intLobbySuitLevel,
                lobby_topic : intLobbyTopic, 
                con_sn : intConsultantSn
            },
            success: function (result) {
                $("#" + strCheckDivOpen + "2").html(result);
            },
            error: function (result) {
                //alert("error:" + result);
            }
        });
        //查詢 查無條件且有查詢動作 紀錄變數
        var strHasData = $("#"+ strCheckDivOpen + "4").val();
        if ( strHasData == "False" && ( intLobbySuitLevel != "" || intLobbyTopic != "" || intConsultantSn !="" ) ) 
        {
            if ( "morning" == strCheckDivOpen )
            {
                strTime_morning = $("#" + strCheckDivOpen + "5").html() + "\n";
            }
            if ( "noon" == strCheckDivOpen )
            {
                strTime_noon = $("#" + strCheckDivOpen + "5").html() + "\n";
            }
            if ( "night" == strCheckDivOpen )
            {
                strTime_night = $("#" + strCheckDivOpen + "5").html() + "\n";
            }
        }
    }
    //document.body.style.cursor="default";
}

//查詢 警示視窗 
//關閉時段時就把時段清空 只顯示開的時段提醒客戶此時段查無課程
function alertTime(strCheckDivOpen) {
    var strTime = strTime_morning + strTime_noon + strTime_night;
    var strCheckValue = $("#" + strCheckDivOpen + "3").val(); //區間開關變數
    //若狀態打開 或是 一進來的預設打開 就顯示警示視窗
    if ( strCheckValue == "open" || strCheckDivOpen == "default" ) 
    {
        if ( strTime != "")
        {
            alert("以下时段查无大会堂课程：\n"+strTime+"\n请重新选择查询条件。");
        }   
    }
}

//燈箱 確認訂位 lobby_near_order.asp
function submitButtonDialog()
{
    var intSessionStatus = $("#session_status").val();
     $("#dialog-form").dialog("close"); //確認or取消訂課動作完畢再關閉燈箱
    if ( 1 == intSessionStatus )  //若目前狀態已訂課 則click button為取消訂課
    {
        var strSessionSDate = $("#session_sdate").val();
        var intSessionSTime = $("#session_stime").val();
        var intClientAttendSn = $("#client_attend_sn").val();
        var intSessionSn = $("#session_sn").val();
        var strSessionName = ("chk_lobby_class_" + intSessionSn) ;
        //var intPage = $("#now_page").val();
        //var intPage = '<%=intNowPage%>';
        
        //20121026 取消訂課 要移除暫存區(lobby_class.js) 再取消訂課 否則暫存區的課程資料還是會在
        str_order_class_datetime = strSessionSDate.replace(new RegExp("/", "gm"), "") + intSessionSTime.substring(0, 2); //  訂課日期時間 YYYYMMDDHHmm
        g_obj_reservation_class_ctrl.removeClassInfo(str_order_class_datetime, g_obj_reservation_class_ctrl.arr_class_type.lobby); //移除暫存區
        //取消課程  for calendar( intPageType : 2) intClientAttendSn有訂位的才有值
        g_obj_reservation_class_ctrl.cancelClass(intClientAttendSn, strSessionSDate, intSessionSTime, '<%=intNowPage%>', 2, strSessionName);
    }
    else //若目前狀態未訂課 則click button為確認訂課
    {
        var bolCheckOk = g_obj_reservation_class_ctrl.submitReservationClass();
        if ( true == bolCheckOk ) 
        {
            $('#form_reservation_class_module').submit();
        }
    }
    RefreshSession();  //重整 剩餘堂數
}

//預設第一次會跑 燈箱確認/取消button
$(document).ready(function () 
{
    //用session判斷開合 點前三天後三天的區域
    //clear_session
    var strReservedCalOrList = "cal";    // 合約結束日
    <%
    strSessionMorning = getSession("morning", CONST_DECODE_ESCAPE)
    if ( (isEmptyOrNull(strSessionMorning) AND isEmptyOrNull(strSessionNoon) AND isEmptyOrNull(strSessionNight)) or ("false" = strSessionMorning AND "false" = strSessionNoon AND "false"= strSessionNight) ) then
        strSessionMorning = "true"
        Session("morning") = strSessionMorning
    end if
    if  ( "true" = strSessionMorning ) then 
    %>
    showTable('4', '11', '<%=strNowDate%>', '<%=strMaxNowDate%>', 'morning', '<%=intNowPage%>', '<%=str_search_topic%>','<%=str_search_con%>','<%=str_search_lev%>');
    <% end if %>

    <% 
    strSessionNoon = getSession("noon", CONST_DECODE_ESCAPE)
    if  ( "true" = strSessionNoon ) then 
    %>
    showTable('11', '17', '<%=strNowDate%>', '<%=strMaxNowDate%>', 'noon', '<%=intNowPage%>', '<%=str_search_topic%>','<%=str_search_con%>','<%=str_search_lev%>');
    <% end if %>

    <%
    strSessionNight = getSession("night", CONST_DECODE_ESCAPE)
    if  ( "true" = strSessionNight ) then 
    %>
    showTable('17', '23', '<%=strNowDate%>', '<%=strMaxNowDate%>', 'night', '<%=intNowPage%>', '<%=str_search_topic%>','<%=str_search_con%>','<%=str_search_lev%>');
    <% end if %>

    //查詢 預設給default值
    alertTime("default"); 
});
//-->
</script>