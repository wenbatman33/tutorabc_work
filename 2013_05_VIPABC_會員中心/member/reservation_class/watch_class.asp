<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<% 
dim bolDebugMode : bolDebugMode = false '正式改為false
dim bolTestCase : bolTestCase = false
%>

<script type="text/javascript" src="javascript/reservation_class<%=strVIPABCJR%>.js"></script>
<script type="text/javascript" src="javascript/watch_class<%=strVIPABCJR%>.js"></script>
<!--內容start-->
<div class="main_membox">
    <div class='page_title_1'><h2 class='page_title_h2'>观看课表</h2></div>
    <div class="box_shadow" style="background-position:-80px 0px;"></div>
    <!--上版位課程start-->
    <div class="top_class2">
        <!--顯示年和切換月的按鈕 START-->
        <div id="div_calender_head"></div>
        <!--顯示年和切換月的按鈕 END-->
        <%if (checkReservationClassQualification(CONST_CLASS_TYPE_ONE_ON_ONE & "," & CONST_CLASS_TYPE_ONE_ON_THREE, CONST_VIPABC_RW_CONN)) then%>
        <div class="right"><%=str_msg_font_you_can_go%><!--您也可以前往--><a href="/program/member/reservation_class/normal_and_vip.asp"><%=str_msg_font_order_class%><!--预定课程--></a></div>
        <%end if%>
        <div class="rightDate">
            <select id="sel_change_calender_mode" class="class_select" onchange="g_obj_class_calender.onchangeSelectCalenderMode();">
                <option value="1" selected="selected"><%=str_msg_font_calender_mode%></option><!--月历模式-->
                <option value="2"><%=str_msg_font_week_mode%></option><!--列表模式-->
            </select>
        </div>
    </div>
    <!--上版位課程end-->
    <!--課程列表start-->
    <div id="div_calender_day"></div>
    <!--課程列表end-->
    <div class="floatL">
        <span class="ball_red">●</span><span class="color_red">小班制</span><!--一般课( 1-6人 )-->
        <span class="ball_green"> ●</span><span class="color_green">1对1</span><!--一般课( 1對1 )-->
        <span class="ball_blue">●</span><span class="color_blue"><%=str_msg_font_lobby_class%></span><!--大会堂-->
    </div>
    <div class="floatR">
    <span class="color_brown"><%=str_msg_font_calender_note%><!--备注：请将鼠标移至日期上即可观看详细内容--></span></div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<script type='text/javascript' src="/lib/javascript/framepage/js/watch_class.js"></script>
<!--內容end-->