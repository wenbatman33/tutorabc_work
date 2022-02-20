<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<%
'使用新的課程前導頁
'Response.Redirect "class_choose.asp"
'2012-08-28 修正導到語言程度分析 Johnny
Dim envirService 
set envirService = new EnvironmentTestService
Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")  
Dim haveTestOk : haveTestOk = envirService.checkEnvironmentHaveTestOk(objPCInfoInsert)
str_path_info = Request.ServerVariables("PATH_INFO")
if "1" = session("boleverclass") then
	Response.Redirect "class_choose.asp"
elseif (session("bolFirstTestQuestionaryhaveTestNOTOkOrder")="Y" and newMemberTestPass = 1 and not haveTestOk="N")  or ( newMemberTestPass = 1 and  haveTestOk="Y")   then  ' 參考變數 E:\web\www.vipabc.com\program\member\reservation_class\include\include_prepare_and_check_data.inc
	Response.Redirect "class_choose.asp"
else
	if "1" = bolFinishDCGS and ( not isEmptyOrNull(int_client_attend_level) and not CStr(int_client_attend_level) = "0" ) then
		'Call alertGo("亲爱的客户您好\n欢迎您加入VIPABC学习行列，为确保您的上课学习质量，上课前请您先至「新手上路」专区完成下列步骤：\nStep 1.测试上课电脑设备\nStep 2.DCGS学习偏好设定\nStep 3.语言程度分析\nStep 4.观看教学影片","/program/member/environment_test/guide_index.asp",CONST_ALERT_THEN_REDIRECT) 
		'Response.Redirect "/program/member/environment_test/guide_index.asp"
	else
		Response.Redirect "/program/member/learning_level.asp"
	end if
end if
set envirService = nothing
Response.End
%>
<script type="text/javascript">
<!--
    $(document).ready(function(){
        var str_now_page_url = '<%=Request.ServerVariables("PATH_INFO")%>';
        // Ajax 載入 推薦大會堂
		// Ajax 載入 请选择课程类型/时间
        $("#div_add_normal_vip_class").load("/program/member/reservation_class/include/ajax_reservation_normal_and_vip_class.asp", {ndate : new Date(), page_url: escape(str_now_page_url)});

        // Ajax 載入 推薦大會堂
        $("#div_commend_lobby_class_info").load("/program/member/reservation_class/include/ajax_commend_lobby_class.asp", {ndate : new Date()});
    });
//-->
</script>
<!--內容start-->
<div class="main_membox">
    <!--新增一般/VIP 課程 START-->
    <div id="div_add_normal_vip_class"></div>
    <!--新增一般/VIP 課程 END-->
    <!--2010.01.18新增:找不到或沒有資料的共用訊息框-->
    <div class="wowbox">
        <div class="wow">
        <p id="remind_contract_start_msg" class="top_class_qa"><%=getWord("WELCOME_JOIN_VIPABC") & "! " & getWord("YOU_CAN_RESERVATION_NOW") & "<font color='red'>" & str_service_start_date & "</font>" & getWord("AFTER_CONTRACT_START")%></p><!--欢迎您加入VIPABC行列! 您现在可预订YYYY/MM/DD(合约开始)之後的课程。-->
        <p class="top_class_qa"><%=str_msg_font_no_order_class%><!--您目前尚未预订咨询课程--><a href="/program/faq/faq.asp?linkHashPoint=member_kind_q14"><%=str_msg_font_how_to_order%><!--该如何预订？--></a></p>
        </div>
    </div>
    <!--2010.01.18新增:找不到或沒有資料的共用訊息框-->
    <!--推薦大會堂 start-->
    <div id="div_commend_lobby_class_info"></div>
    <!--推薦大會堂 end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->