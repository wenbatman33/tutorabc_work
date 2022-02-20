<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<link type="text/css" href="/lib/javascript/JQuery/css/redmond/jquery-ui-1.8.2.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
<!--
$(document).ready(function () {
	//設定dialog大小
	$("#dialog").dialog("destroy");
	$("#dialog-form").dialog({
		autoOpen: false,
		height: 350,
		width: 550,
		modal: true,
		close: function() {
			$("#dialog-form").html("");
		}
	}); 
});

/// <summary>
/// 20100902 阿捨新增 小幫手頁面
/// 呼叫ajax載入iframe (為了符合css)
/// </summary>
function ShowDialog(){
	//dialog載入
	$("#dialog-form").html("<iframe src='../include/realtime_message.asp?client_sn=<%=g_var_client_sn%>' width='100%' height='100%' scrolling='no' frameborder='0'></iframe>");
	
	//開啟dialog
	$("#dialog-form").dialog("open");
	$("#dialog-form").dialog("option", "resizable", false);
}

$(".AreaExpand").each(function ()
{
    var strTmpHeight = $(this).css("height");
    strTmpHeight = (strTmpHeight == null) ? "0" : strTmpHeight;
    strTmpHeight = strTmpHeight.replace("px", "");
    $(this).TextAreaExpander(strTmpHeight, 999);
});

$(function ()
{
    $("#btn_submit").click(function()
    {       
        if (check_form())
        {
            $("#frm1").submit(); 
        }
    });
    //檢查是否填寫        
});

function check_form()
{   
    //請輸入教學質量
    if ($("input[name='consultant']:checked").size() <= 0)
    {
        alert("请输入顾问教学质量");
        return false;
    }
    //請輸入讲话速度
    if ($("input[name='speed']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_TALK_SPEED")%>");
        return false;
    }
    //請輸入時間分配
    if ($("input[name='distribution']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_SET_TIME")%>");
        return false;
    }
    //請輸入教學技巧
    if ($("input[name='tskill']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_TEACH_SKILL")%>");
        return false;
    }
    //請輸入教學態度
    if ($("input[name='behavior']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_TEACH_ATTITUDE")%>");
        return false;
    }
    //請輸入教材品質
    if ($("input[name='materials']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_MATERIAL_QUALITY")%>");
        return false;
    }
    //請輸入難易度
    if ($("input[name='difficulty']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_DIFFUCALLTY")%>");
        return false;
    }
    //請輸入需求面
    if ($("input[name='requirement']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_NEED")%>");
        return false;
    }
    //請輸入通訊質量
    if ($("input[name='overall']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_CONECTION")%>");
        return false;
    }
    //請輸入通訊狀態
    if ($("input[name='connection']:checked").size() <= 0)
    {
        alert("<%=getWord("PLZ_INPUT_CONECTION_STATUS")%>");
        return false;
    }
	$("#btn_submit").attr("disabled",true);
    return true;
}
//-->       
</script>
<%
Dim str_session_sn : str_session_sn = request("Session_sn")
if ( str_Session_sn="" ) then 
    response.Redirect "learning_record.asp"
elseif ( g_var_client_sn = "" ) then
    '如果session 掉了 導回首頁
    response.Redirect "http://"&CONST_VIPABC_WEBHOST_NAME&"/index.asp"
end if

Dim mtl '小幫手提供的conn開法所需變數
Dim str_sql '小幫手提供的conn開法所需變數
Dim var_arr '小幫手提供的conn開法所需變數
Dim str_tmp_sql '小幫手提供的conn開法所需變數
Dim arr_result '小幫手提供的conn開法所需變數

Dim bol_feedback_disable : bol_feedback_disable = false '預設都可以填

Dim str_rs_consultant_points
Dim str_rs_materials_points
Dim str_rs_overall_points
Dim str_rs_comments
Dim str_rs_compliment
Dim str_rs_suggestions
Dim str_rs_behavior
Dim str_rs_distribution
Dim str_rs_spoke
Dim str_rs_speed
Dim str_rs_interaction
Dim str_rs_requirement
Dim str_rs_difficulty
Dim str_rs_repetition
Dim str_rs_connection
Dim str_rs_tskill
Dim str_rs_t_other
Dim str_rs_c_other
Dim str_rs_m_other

Dim str_session_year : str_session_year = left(str_session_sn,4)
Dim str_session_month : str_session_month = mid(str_session_sn,5,2)
Dim str_session_day : str_session_day = mid(str_session_sn,7,2)
Dim str_session_hour : str_session_hour	= mid(str_session_sn,9,2)
Dim str_session_date : str_session_date	= DateValue(str_session_year&"/"&str_session_month&"/"&str_session_day)

Dim str_frml_action	: str_frml_action = ""
Dim str_submit_value : str_submit_value	= ""

Dim str_room_type 

'因session_sn是明碼 故判斷是否沒有該堂課轉址
var_arr = Array(g_var_client_sn, str_session_sn)
str_tmp_sql = "SELECT sn FROM client_attend_list(nolock) WHERE (client_sn = @client_sn ) AND (valid = 1) AND (session_sn = @session_sn )"
arr_result = excuteSqlStatementRead(str_tmp_sql, var_arr, CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then  '有資料
	if (Ubound(arr_result) < 0) then
		'g_bol_feedback_disable = true
		response.Write "<script langeage='javascript'>alert('"&getWord("NO_DATA")&"');location.href='learning_record.asp';</script>"
	end if
end if


'判斷是否有訂該堂課 沒上該堂課
var_arr = Array(g_var_client_sn, str_session_sn)
str_tmp_sql = "SELECT session_state FROM client_attend_list(nolock) WHERE (client_sn = @client_sn ) AND (valid = 1) AND (session_state = '0') AND (session_sn = @session_sn )"
arr_result = excuteSqlStatementRead(str_tmp_sql, var_arr, CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then  '有資料
	if (Ubound(arr_result) > 0) then
		bol_feedback_disable = true
	end if
end if

'判斷可以不是填寫表單的時間
'客戶反應2/15那天無法填寫評鑑，所以特地開放讓客戶填
if (("2011021521404" = str_session_sn) and ("244276" = g_var_client_sn)) then
	bol_feedback_disable = false
elseif ( (datediff("d", str_session_date, now()) > 1 ) or (datediff("d", str_session_date, now())=1 and hour(now())>9) ) then 
	bol_feedback_disable = true
end if

'撈會員對此堂課的feedback
var_arr = Array(str_session_sn, g_var_client_sn)
str_tmp_sql="select * from client_session_evaluation(nolock) where session_sn=@g_str_session_sn and client_sn=@g_var_client_sn "
arr_result = excuteSqlStatementRead(str_tmp_sql, var_arr, CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then  '有資料
	if (Ubound(arr_result) > 0) then
        str_rs_consultant_points = trim(arr_result(6,0)) 
        str_rs_materials_points = trim(arr_result(7,0)) 
        str_rs_overall_points = trim(arr_result(8,0))
        str_rs_comments = trim(arr_result(9,0))
        str_rs_compliment = trim(arr_result(10,0))
        str_rs_suggestions = trim(arr_result(11,0))
	
        str_rs_behavior = trim("0, "&arr_result(15,0))
        str_rs_distribution = trim("0, "&arr_result(16,0))
        str_rs_spoke = trim("0, "&arr_result(17,0))
        str_rs_speed = trim("0, "&arr_result(18,0))
        str_rs_interaction = trim("0, "&arr_result(19,0))
        str_rs_requirement = trim("0, "&arr_result(21,0))

        str_rs_difficulty = trim("0, "&arr_result(22,0))
        str_rs_repetition = trim("0, "&arr_result(23,0))
        str_rs_connection = trim("0, "&arr_result(24,0))
        str_rs_tskill = trim("0, "&arr_result(20,0))
	
        str_rs_t_other = trim(arr_result(25,0))
        str_rs_c_other = trim(arr_result(26,0))
        str_rs_m_other = trim(arr_result(27,0))
	
        Dim str_input_disabled : str_input_disabled = ""
        bol_feedback_disable = true  '填過了
        if ( true = bol_feedback_disable ) then
            str_input_disabled = "disabled"
        end if
	end if
end if
%>
<div id="temp_contant">
    <!--內容start-->
    <div id="issue_membox">
<%
'判斷是不是大會堂 要不要客戶互評
var_arr = Array(g_var_client_sn, str_session_sn)
str_tmp_sql = "SELECT special_sn,class_yesno FROM client_attend_list(nolock) WHERE (client_sn = @client_sn ) AND (session_sn = @session_sn ) AND (valid = 1)"
arr_result = excuteSqlStatementRead(str_tmp_sql, var_arr, CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then  '有資料
	if (Ubound(arr_result) > 0) then
		if ( isnull(arr_result(0,0)) or arr_result(0,0)=""  ) then  '一般課程
			str_frml_action = "session_feedback_classmate.asp"
			str_submit_value= getWord("continue_feedback_other")
			str_room_type = "regular"
		else '大會堂
			str_frml_action = "session_feedback_exe.asp"
			str_submit_value= getWord("BUY_1_40")
			str_room_type = "special"
		end if
	end if
end if
%>
<script type="text/javascript">
$(function ()
{
    //將客戶填過得consultant答案顯示出來
    $("input[name='consultant']").each(function ()
    {
        if ($(this).val() == "<%=str_rs_consultant_points%>")
        {
            $(this).attr("checked", true);
        }
    });
    //將客戶填過得overall答案顯示出來
    $("input[name='overall']").each(function ()
    {
        if ($(this).val() == "<%=str_rs_overall_points%>")
        {
            $(this).attr("checked", true);
        }
    });
    //將客戶填過得materials答案顯示出來
    $("input[name='materials']").each(function ()
    {
        if ($(this).val() == "<%=str_rs_materials_points%>")
        {
            $(this).attr("checked", true);
        }
    });
    //分散
    checkAnswer("<%=str_rs_speed%>", "speed");
    checkAnswer("<%=str_rs_distribution%>", "distribution");
    checkAnswer("<%=str_rs_tskill%>", "tskill");
    checkAnswer("<%=str_rs_behavior%>", "behavior");
    checkAnswer("<%=str_rs_difficulty%>", "difficulty");
    checkAnswer("<%=str_rs_requirement%>", "requirement");
    checkAnswer("<%=str_rs_connection%>", "connection");
    //將所有的input 如果已經填過 變成disabled
    $("input").attr("disabled", "<%=str_input_disabled%>");
});

function checkAnswer(strAnswer, itemName)
{
    //將客戶填過得distribution答案顯示出來
    $("input[name='" + itemName + "']").each(function ()
    {
        var inputValue = $(this).val();
        var strRegularExp = new RegExp(inputValue, "g")
        //如果有在其中
        if (strRegularExp.test(strAnswer))
        {
                $(this).attr("checked", true);
        }
    });
    //將所有的input 如果已經填過 變成disabled
}

function CDDL()
{
	if (document.all("HDDL").style.visibility=="visible")
	{
		document.all("HDDL").style.visibility="hidden";
	}
	else
	{
		document.all("HDDL").style.visibility="visible";
	}
}
</script>
<!--问题小帮手 start-->
<div id="dialog-form" title="问题小帮手" /></div>
<!--问题小帮手 end-->
        <form action="<%=str_frml_action%>" name="frm1" id="frm1" method="post" onSubmit="return check_form();">
        <div><table align="center" width="100%"><tr><td align="center"><img border="0" src="/images/images_cn/homemeet-helper.jpg" style="cursor:hand" onclick="ShowDialog();" /></td></tr></table></div>
        <input name="session_sn" id="session_sn" type="hidden" value="<%=str_Session_sn%>" />
        <input name="session_feedback_disable" id="session_feedback_disable" type="hidden"
            value="<%=bol_feedback_disable%>" />
        <!--問卷資料 start-->
        <div class="main_mylist">
            <ul>
                <li><span class="m-right5">
                    <%=getWord("session_feedback")%></span><span class="black"><span class="red">10</span><%=getWord("session_feedback_2")%>，<span
                        class="red">1</span><%=getWord("session_feedback_3")%>。<%=getWord("mouse_over")%></span></li>
            </ul>
        </div>
        <!--問卷資料 end-->
        <!--教學質量標頭 start-->
        <div class="con_title">
            <div class="con_header_left">
                <img src="/images/images_cn/arrow_path.gif" alt="" width="4" height="12" hspace="2" />顾问教学质量
            </div>
            <div class="con_header_right ">
                <input type="radio" name="consultant" value="10" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_1")%>">10</span>
                <input type="radio" name="consultant" value="9" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_2")%>">9</span>
                <input type="radio" name="consultant" value="8" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_3")%>">8</span>
                <input type="radio" name="consultant" value="7" id="test" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_4")%>">7</span>
                <input type="radio" name="consultant" value="6" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_5")%>">6</span>
                <input type="radio" name="consultant" value="5" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_6")%>">5</span>
                <input type="radio" name="consultant" value="4" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_7")%>">4</span>
                <input type="radio" name="consultant" value="3" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_8")%>">3</span>
                <input type="radio" name="consultant" value="2" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_9")%>">2</span>
                <input type="radio" name="consultant" value="1" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_consultant_10")%>">1</span>
            </div>
        </div>
        <div class="clear">
        </div>
        <!--教學質量 end-->
        <!--問卷一 start-->
        <div class="iss_main_memfiled_all">
            <div class="iss_main_memfiled_left_2">
                <%=getWord("talk_speed")%></div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="radio" name="speed" value="2" /><%=getWord("fine")%></li>
                    <li>
                        <input type="radio" name="speed" value="3" /><%=getWord("slow")%></li>
                    <li>
                        <input type="radio" name="speed" value="1" /><%=getWord("fast")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--問卷一 end-->
        <div class="iss_main_memfiled_line">
        </div>
        <!--問卷二 start-->
        <div class="iss_main_memfiled_all">
            <div class="iss_main_memfiled_left_2">
                <%=getWord("time_allot")%>
                (<%=getWord("DCGS_10")%>)</div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="distribution" value="3" /><%=getWord("fine")%></li>
                    <li>
                        <input type="checkbox" name="distribution" value="1" /><%=getWord("talk_time_allot_not_enough")%></li>
                    <li>
                        <input type="checkbox" name="distribution" value="2" /><%=getWord("talk_chance_not_enough")%></li>
                    <li>
                        <input type="checkbox" name="distribution" value="4" /><%=getWord("mtl_not_finish")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--問卷二 end-->
        <div class="iss_main_memfiled_line">
        </div>
        <!--問卷三 start-->
        <div class="iss_main_memfiled_all">
            <div class="iss_main_memfiled_left_2">
                <%=getWord("teaching_skill")%>
                (<%=getWord("DCGS_10")%>)</div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="tskill" value="3" /><%=getWord("fine")%></li>
                    <li>
                        <input type="checkbox" name="tskill" value="5" /><%=getWord("not_point")%></li>
                    <li>
                        <input type="checkbox" name="tskill" value="1" /><%=getWord("only_read_mtl")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="iss_main_memfiled_left_2">
            </div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="tskill" value="4" /><%=getWord("not_guidance")%></li>
                    <li>
                        <input type="checkbox" name="tskill" value="6" /><%=getWord("not_program")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--問卷三 end-->
        <div class="iss_main_memfiled_line">
        </div>
        <!--問卷四 start-->
        <div class="iss_main_memfiled_all">
            <div class="iss_main_memfiled_left_2">
                <%=getWord("teaching_attitude")%>(<%=getWord("DCGS_10")%>)</div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="behavior" value="5" /><%=getWord("conscientious")%></li>
                    <li>
                        <input type="checkbox" name="behavior" value="4" /><%=getWord("kindly")%></li>
                    <li>
                        <input type="checkbox" name="behavior" value="3" /><%=getWord("patience")%></li>
                    <li>
                        <input type="checkbox" name="behavior" value="2" /><%=getWord("loose")%></li>
                    <li>
                        <input type="checkbox" name="behavior" value="1" /><%=getWord("bad")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--問卷四 end-->
        <!--教學品質標頭 start-->
        <div class="con_title">
            <div class="con_header_left">
                <img src="/images/images_cn/arrow_path.gif" alt="" width="4" height="12" hspace="2" /><%=getWord("mtl_quality")%>
            </div>
            <div class="con_header_right ">
                <input type="radio" name="materials" value="10" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_1")%>">10</span>
                <input type="radio" name="materials" value="9" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_2")%>">9</span>
                <input type="radio" name="materials" value="8" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_3")%>">8</span>
                <input type="radio" name="materials" value="7" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_4")%>">7</span>
                <input type="radio" name="materials" value="6" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_5")%>">6</span>
                <input type="radio" name="materials" value="5" />
                <span class="color_1 bold m-right5" title="<%=getWord("FEEDBACK_CONSULTANT_6")%>">5</span>
                <input type="radio" name="materials" value="4" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_8")%>">4</span>
                <input type="radio" name="materials" value="3" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_9")%>">3</span>
                <input type="radio" name="materials" value="2" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_10")%>">2</span>
                <input type="radio" name="materials" value="1" />
                <span class="color_1 bold m-right5" title="<%=getWord("feedback_mtl_0")%>">1</span>
            </div>
        </div>
        <div class="clear">
        </div>
        <!--教學品質 end-->
        <!--問卷一 start-->
        <div class="iss_main_memfiled_all">
            <div class="iss_main_memfiled_left_2">
                <%=getWord("difficulties")%></div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="radio" name="difficulty" value="1" /><%=getWord("exactly")%></li>
                    <li>
                        <input type="radio" name="difficulty" value="2" /><%=getWord("hard")%></li>
                    <li>
                        <input type="radio" name="difficulty" value="3" /><%=getWord("easy")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--問卷一 end-->
        <div class="iss_main_memfiled_line">
        </div>
        <!--問卷二 start-->
        <div class="iss_main_memfiled_all">
            <div class="iss_main_memfiled_left_2">
                <%=getWord("requirement")%></div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="requirement" value="10" /><%=getWord("requirement_1")%></li>
                    <li>
                        <input type="checkbox" name="requirement" value="6" /><%=getWord("requirement_2")%></li>
                    <li>
                        <input type="checkbox" name="requirement" value="4" /><%=getWord("requirement_3")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="iss_main_memfiled_left_2">
            </div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="requirement" value="5" /><%=getWord("requirement_4")%></li>
                    <li>
                        <input type="checkbox" name="requirement" value="8" /><%=getWord("requirement_5")%></li>
                    <li>
                        <input type="checkbox" name="requirement" value="7" /><%=getWord("requirement_6")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="iss_main_memfiled_left_2">
            </div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="requirement" value="9" /><%=getWord("requirement_7")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--問卷二 end-->
        <!--通讯质量標頭 start-->
        <div class="con_title">
            <div class="con_header_left">
                <img src="/images/images_cn/arrow_path.gif" alt="" width="4" height="12" hspace="2" /><%=getWord("communication_quality")%>
            </div>
            <div class="con_header_right ">
                <input type="radio" name="overall" value="10" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_1")%>">10</span>
                <input type="radio" name="overall" value="9" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_2")%>">9</span>
                <input type="radio" name="overall" value="8" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_3")%>">8</span>
                <input type="radio" name="overall" value="7" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_4")%>">7</span>
                <input type="radio" name="overall" value="6" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_5")%>">6</span>
                <input type="radio" name="overall" value="5" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_6")%>">5</span>
                <input type="radio" name="overall" value="4" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_7")%>">4</span>
                <input type="radio" name="overall" value="3" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_8")%>">3</span>
                <input type="radio" name="overall" value="2" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_9")%>">2</span>
                <input type="radio" name="overall" value="1" />
                <span class="color_1 bold m-right5" title="<%=getWord("communication_10")%>">1</span>
            </div>
        </div>
        <div class="clear">
        </div>
        <!--通讯质量 end-->
        <!--問卷一 start-->
        <div class="iss_main_memfiled_all">
            <div class="iss_main_memfiled_left_2">
                <%=getWord("communication_status")%>(<%=getWord("DCGS_10")%>)</div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="connection" value="7" /><%=getWord("fine")%></li>
                    <li>
                        <input type="checkbox" name="connection" value="2" /><%=getWord("communication_status_1")%></li>
                    <li>
                        <input type="checkbox" name="connection" value="1" /><%=getWord("communication_status_2")%></li>
                    <li>
                        <input type="checkbox" name="connection" value="4" /><%=getWord("communication_status_3")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="iss_main_memfiled_left_2">
            </div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="connection" value="5" /><%=getWord("communication_status_4")%></li>
                    <li>
                        <input type="checkbox" name="connection" value="3" /><%=getWord("communication_status_5")%></li>
                    <li>
                        <input type="checkbox" name="connection" value="8" /><%=getWord("communication_status_6")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="iss_main_memfiled_left_2">
            </div>
            <div class="iss_main_memfiled_right_3">
                <ul>
                    <li>
                        <input type="checkbox" name="connection" value="9" /><%=getWord("REQUIREMENT_7")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--問卷二 end-->
        <!--通讯质量標頭 start-->
        <div class="con_title">
            <div class="con_header_left">
                <img src="/images/images_cn/arrow_path.gif" alt="" width="4" height="12" hspace="2" />
                <%=getWord("even_its_broken_English")%>, even it's broken English!</div>
            <div class="con_header_right ">
            </div>
        </div>
        <div class="clear">
        </div>
        <!--通讯质量 end-->
        <div class="advice">
            <!--09/11/19:修改start-->
            <%=getWord("propose")%><br>
            <!--09/11/19:修改end-->
            <textarea id="suggestions" name="suggestions" class="login_box5 width17 height3 AreaExpand"><%if str_rs_suggestions<>"" then response.write str_rs_suggestions end if%></textarea>
            <!--09/11/19:新增2個br-->
            <br />
            <br />
            <!--09/11/19:修改start-->
            <%=getWord("praise")%><br>
            <!--09/11/19:修改end-->
            <textarea id="compliment" name="compliment" class="login_box5 width17 height3 AreaExpand"><%if str_rs_compliment<>"" then response.write str_rs_compliment end if%></textarea>
            <div class="clear">
            </div>
			<%if isEmptyOrNull(str_rs_consultant_points) then%>
				<div>
				<input type="checkbox" name="C1" value="ON" id="C1" onClick="CDDL();"><%="请客服人员以电话与我联络(我们将于24小时内尽速与您联系。)"%><label id="HDDL" style="visibility:hidden;"> <%="联络时间请于"%>
			<%
				mystring = split("不拘|上午|下午|晚上", "|", -1, 1)
			%>
				<select size="1" name="ctime" style="font-family: Arial; font-size: 8pt">
				<option value="0"><%=mystring(0)%></option>
				<option value="1"><%=mystring(1)%></option>
				<option value="2"><%=mystring(2)%></option>
				<option value="3"><%=mystring(3)%></option>
				</select>
				</label>
				</div>
			<%end if%>
        </div>
        <!--確認鈕 start-->
        <div class="t-right m-top15 ">
            <!--09/11/19m-top25改m-top15-->
            <input type="button" name="btn_submit" id="btn_submit" value="+ <%=str_submit_value%>"
                class="btn_3" <% if ( str_room_type = "special" and  bol_feedback_disable ) then %> disabled="disabled"
                <% end if %>/>
        </div>
        <!--確認鈕 end-->
        </form>
    </div>
    <!--內容 end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<% 
'20101012 阿捨新增 寫入出席狀態
dim arrParam : arrParam = null
dim strSql : strSql = ""
dim arrSqlResult : arrSqlResult = ""

arrParam = Array(9, str_session_sn, g_var_client_sn)
strSql = "UPDATE client_attend_list SET " 
strSql = strSql & "class_yesno = @class_yesno "
strSql = strSql & "WHERE (client_attend_list.session_sn = @session_sn) AND (client_attend_list.client_sn = @client_sn)"
arrSqlResult = excuteSqlStatementWrite(strSql, arrParam, CONST_VIPABC_R_CONN)
if (isSelectQuerySuccess(arrSqlResult)) then
	if (arrSqlResult < 0) then 
		'寫入失敗
	else
		'寫入成功
	end if
end if
%>
<iframe src="refresh.html" frameborder="0" width="10" height="10"></iframe>