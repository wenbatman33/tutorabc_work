<!--#include virtual="/lib/include/global.inc" -->
<%
	'變數設定區
	Dim SessionSn : SessionSn = Request("Session_sn") '取得課程編號
	Dim intClientSn : intClientSn = Request("client_sn") '取得客戶編號
	Dim qlNumberOfEvaluation : qlNumberOfEvaluation = 25 '問卷編號, 課後評比
	Dim qlNumberOfClassmate : qlNumberOfClassmate = 29 '問卷編號, 對其他客戶評鑑的問卷
%>
<!--#include virtual="/program/include/language/cn.asp"-->
<link type="text/css" href="/lib/javascript/JQuery/css/redmond/jquery-ui-1.8.2.custom.css" rel="stylesheet" />
<link type="text/css" href="/css/new_feedback.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery.js"></script><!--因為dialog的關係無法改版本-->
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript">
<!--
/*
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
*/


//頁面調整 開始---
$(function () {
	var Level_1 = $(".new-feedback").find("#level_1");
	var Level_2 = $(".new-feedback").find("#level_2");
	var tmp_level_2;
    Level_2.hide(); //把所有Level_2的都先hide起來
    Level_1.click(function () {
		tmp_level_2 = $(this).find("#level_2");
        if ($($(this).children()[2]).find(":checked").length == 1) {
			tmp_level_2.hide();
			if(tmp_level_2.length == 2) {
				$(tmp_level_2[1]).show();
				$(tmp_level_2[0]).find(":checked").attr("checked",false);
			}
			else {
				tmp_level_2.show();
			}
        }
        else {
            tmp_level_2.hide();
			if ($($(this).children()[0]).find(":checked").length == 1) {
				if(tmp_level_2.length == 2) {
					$(tmp_level_2[0]).show();
					$(tmp_level_2[1]).find(":checked").attr("checked",false);
				}
				else
				{
					$(tmp_level_2).hide();
					$(tmp_level_2).find(":checked").attr("checked",false);
				}
			}
			else
			{
				tmp_level_2.find(":checked").attr("checked",false);
			}
        }
    });
	
	//教學態度調整位置,但是如果教學態度不是第5題時, 就要改變了
	var tmp_level_1_4 = $($(".new-feedback").find("#level_1")[4]);
	var level_1_tmp = tmp_level_1_4.children("span");
	var level_2_tmp = tmp_level_1_4.find("#level_2");
	tmp_level_1_4.html(level_1_tmp);
	tmp_level_1_4.append(level_2_tmp);
	
	//需求面調整位置,但是如果需求面不是第8題時, 就要改變了
	var tmp_level_1_7 = $($(".new-feedback").find("#level_1")[7]);
	tmp_level_1_7.find("span").each(function(index) {
		if (index != 0) {
			$(this).before("<br/>");
		}
	});
	tmp_level_1_7.find(":radio").click(function() {
		if ($(this).val() == "9") {
			tmp_level_1_7.find('input[name="rn"]').attr("disabled","");
		}
		else {
			tmp_level_1_7.find('input[name="rn"]').val("");
			tmp_level_1_7.find('input[name="rn"]').attr("disabled","false");
		}
	});

	//教學技巧Level2調整位置
	$(Level_1[3]).find("#level_2").find("span").each(function(index) {
		if (index != 0) {
			$(this).before("<br/>");
		}
	});

	//通訊狀態Level2調整位置
	$(Level_1[9]).find("#level_2").find("span").each(function(index) {
		if (index != 0) {
			$(this).before("<br/>");
		}
	});

	//其他同學上課評鑑調整,第2題選項調整
	if (Level_1.length > 10) {
		for (var i=11;i<=Level_1.length;i=i+2) {
			$(Level_1[i]).find("#level_2").find("span").each(function(index) {
				if (index != 0) {
					$(this).before("<br/>");
				}
			});
		}
	}
});
//頁面調整 結束---

//滑鼠移動到燈箱上的事件 開始--
$(function(){      
	$("#teacherscore").hover(function() {   
			$("#ConsultantRating").css("left",$(this).offset().left+30);
			$("#ConsultantRating").css("top",$(this).offset().top+40);  
			$("#ConsultantRating").show();
		},
		function() {
			$("#ConsultantRating").hide();
	});

	$("#mertirialscore").hover(function(){   
			$("#MaterialRating").css("left",$(this).offset().left+30);
			$("#MaterialRating").css("top",$(this).offset().top+40);  
			$("#MaterialRating").show();
		},
		function() {
			$("#MaterialRating").hide();
    });
	
	$("#netscore").hover(function(){   
            $("#NetRating").css("left",$(this).offset().left+30);
			$("#NetRating").css("top",$(this).offset().top+40);  
            $("#NetRating").show();
		},
		function() {
		    $("#NetRating").hide();
    });

}); 
//滑鼠移動到燈箱上的事件 結束--

	function chkdata() {
		var Level_1 = $(".new-feedback").find("#level_1");
		var strMsg = "";
		//顧問教學品質
		if ($(Level_1[0]).find(":radio:checked").length == 0) {
			strMsg = strMsg + "[<%=strLangSessionFeedback00024%>]";
		}
		//教材品質
		if ($(Level_1[5]).find(":radio:checked").length == 0) {
			strMsg = strMsg + "[<%=strLangSessionFeedback00003%>]";
		}
		//通訊品質
		if ($(Level_1[8]).find(":radio:checked").length == 0) {
			strMsg = strMsg + "[<%=strLangSessionFeedback00004%>]";
		}
		
		//時間分配
		if ($($(Level_1[2]).children()[2]).find(":checked").length == 1) {
			if($(Level_1[2]).find("#level_2").find(":checkbox:checked").length == 0) {
				$($(Level_1[2]).find("#level_2").find(":checkbox")[0]).focus();
				alert("[<%=strLangSessionFeedback00005&"] "&strLangSessionFeedback00009%>");
				return false;
			}
		}

		//教學技巧
		if ($($(Level_1[3]).children()[2]).find(":checked").length == 1) {
			if($(Level_1[3]).find("#level_2").find(":checkbox:checked").length == 0) {
				$($(Level_1[3]).find("#level_2").find(":checkbox")[0]).focus();
				alert("[<%=strLangSessionFeedback00006&"] "&strLangSessionFeedback00009%>");
				return false;
			}
		}

		//教學態度
		if ($($(Level_1[4]).children()[2]).find(":checked").length == 1) {
			if($($(Level_1[4]).find("#level_2")[1]).find(":checkbox:checked").length == 0) {
				$($($(Level_1[4]).find("#level_2")[1]).find(":checkbox")[0]).focus();
				alert("[<%=strLangSessionFeedback00007&"] "&strLangSessionFeedback00009%>");
				return false;
			}
		}
		if ($($(Level_1[4]).children()[0]).find(":checked").length == 1) {
			if($($(Level_1[4]).find("#level_2")[0]).find(":checkbox:checked").length == 0) {
				$($($(Level_1[4]).find("#level_2")[0]).find(":checkbox")[0]).focus();
				alert("[<%=strLangSessionFeedback00007&"] "&strLangSessionFeedback00009%>");
				return false;
			}
		}

		//通訊狀態
		if ($($(Level_1[9]).children()[2]).find(":checked").length == 1) {
			if($(Level_1[9]).find("#level_2").find(":checkbox:checked").length == 0) {
				$($(Level_1[9]).find("#level_2").find(":checkbox")[0]).focus();
				alert("[<%=strLangSessionFeedback00008&"] "&strLangSessionFeedback00009%>");
				return false;
			}
		}
		
		//其他同學上課狀況
		var classmateStatus = true;
		$(".classmate").find("#level_1").each(function() {
			if ($($(this).children()[2]).find(":checked").length == 1) { //判斷是否有選取不佳選項
				if ($(this).find("#level_2").find(":checkbox:checked").length == 0) { //判斷至少勾選一個check box
					$($(this).find("#level_2").find(":checkbox")[0]).focus();
					classmateStatus = false;
				}
				else { //如果已有勾選一個或以上的check box
					if ($(this).find("#level_3").length != 0) { //判斷第三層的radio button是否有選取
						$(this).find("#level_2").find(":checkbox:checked").each(function() {
							if ($(this).parent().next().find(":radio:checked").length == 0) {
								$($(this).parent().next().find(":radio")[0]).focus();
								classmateStatus = false;
							}
						});
						$(this).find("#level_2").find(":checkbox:not(:checked)").each(function() {
							$(this).parent().next().find(":radio").attr("checked",false);
						});
					}
				}
			}
		});
		if (!classmateStatus){
			alert("<%=strLangSessionFeedback00012%>");
			return false;
		}
		
		if (strMsg != "") {
			alert("<%=strLangSessionFeedback00010%> " + strMsg);
			return false;
		}
		else
		{
			//資料儲存中，請確定是否送出
             var strTempMsg = "<%=strLangSessionFeedback00023%>";
            if(confirm(strTempMsg))
            {
                return true;
            }
            else
            {
                return false;
            }
		}
		
	}

$(function() {
	$("#C1").click(function() {
		if ($(this).attr("checked"))
		{
			document.all("HDDL").style.visibility="visible";
		}
		else
		{
			document.all("HDDL").style.visibility="hidden";
		}
	});
});
//-->       
</script>
<%
Dim str_session_sn : str_session_sn = request("Session_sn")
'if ( str_Session_sn="" ) then 
'    response.Redirect "learning_record.asp"
'elseif ( g_var_client_sn = "" ) then
'    '如果session 掉了 導回首頁
'    response.Redirect "http://"&CONST_VIPABC_WEBHOST_NAME&"/index.asp"
'end if

Dim mtl '小幫手提供的conn開法所需變數
Dim str_sql '小幫手提供的conn開法所需變數
Dim var_arr '小幫手提供的conn開法所需變數
Dim str_tmp_sql '小幫手提供的conn開法所需變數
Dim arr_result '小幫手提供的conn開法所需變數

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

'因session_sn是明碼 故判斷是否沒有該堂課轉址
'var_arr = Array(g_var_client_sn, str_session_sn)
'str_tmp_sql = "SELECT sn FROM client_attend_list(nolock) WHERE (client_sn = @client_sn ) AND (valid = 1) AND (session_sn = @session_sn )"
'arr_result = excuteSqlStatementRead(str_tmp_sql, var_arr, CONST_VIPABC_RW_CONN)
'if (isSelectQuerySuccess(arr_result)) then  '有資料
'	if (Ubound(arr_result) < 0) then
'		'g_bol_feedback_disable = true
'		response.Write "<script langeage='javascript'>alert('"&getWord("NO_DATA")&"');location.href='learning_record.asp';</script>"
'	end if
'end if

'判斷是否由咨詢室進入

if Request("status")="save" then '不是從咨詢室進入
	'**** 判斷客戶有無上課 start ****
	str_sql = " SELECT  ClassStatusId "
	str_sql = str_sql & " FROM    member.ClassRecord(NOLOCK) "
	str_sql = str_sql & " WHERE   ( client_sn = @intClientSn ) "
	str_sql = str_sql & "         AND ( ClassRecordValid = 1 ) "
	str_sql = str_sql & "         AND ( CONVERT(VARCHAR(8), member.ClassRecord.ClassStartDate, 112) "
	str_sql = str_sql & "               + RIGHT('00'+ ISNULL(CONVERT(VARCHAR(2), member.ClassRecord.ClassStartTime),''), 2) "
	str_sql = str_sql & "               + RIGHT('000'+ ISNULL(CONVERT(VARCHAR(3), member.ClassRecord.ClassRoomNumber),''), 3) = @SessionSn ) "
	var_arr = Array(intClientSn,SessionSn)
	arr_result = excuteSqlStatementRead(str_sql , var_arr, CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) > 0) then
			if (arr_result(0) = 3 ) then 'status = 3 出席
				bolClassAttend = true
			else
				bolClassAttend = false
			end if
		end if
			bolClassAttend = false
	else
		bolClassAttend = false
	end if 			
	'**** 判斷客戶有無上課 end ****
else '是從咨詢室進入
	bolClassAttend = false
end if
'response.write bolClassAttend

%>
<div id="temp_contant">
    <!--內容start-->
    <div id="issue_membox">
<!--问题小帮手 start-->
<!--<div id="dialog-form" title="问题小帮手" /></div>-->
<!--问题小帮手 end-->
        <form action="session_feedback_jr_exe.asp" name="frm1" id="frm1" method="post" onSubmit="return chkdata();">
        <div><table align="center" width="100%"><tr><td align="center"><img border="0" src="/images/images_cn/homemeet-helper.jpg" style="cursor:hand" onclick="parent.ShowDialog();" /></td></tr></table></div>
        <input name="session_sn" id="session_sn" type="hidden" value="<%=str_Session_sn%>" /> 
		<input name="client_sn" id="client_sn" type="hidden" value="<%=intClientSn%>" />
		<input type="hidden" name="qlNumberOfEvaluation" value="<%=qlNumberOfEvaluation%>">
		<input type="hidden" name="qlNumberOfClassmate" value="<%=qlNumberOfClassmate%>">
<!--Beson 20111213 新版課後評鑑 Start-->
<%
function extendTree(ByVal qlNumber, ByVal parent, ByVal divLevel, ByVal typeOfInput, ByVal valueOfClassmate, ByVal valueOfQuestion, ByVal actionType, ByVal recordOfQuestion)
'/****************************************************************************************
'描述      ：產生題目的答案選項[以遞迴的方法]
'傳入參數  ：[qlNumber:Integer] 問卷編號 - questionary_list.ql_number
'			 [parent:Integer] 題目編號 - questionary_question_list.qql_number
'            [divLevel:Integer] 選項遞迴的階層
'            [typeOfInput:Array] 每一選項階層顯示的型別 - 例:Array("0","radio","checkbox","radio")
'			 [valueOfClassmate:Integer] 學生的編號 - 從1開始
'			 [valueOfQuestion:Integer] 題目的編號 - 從1開始
'			 [actionType:String] 儲存或是顯示 - "show" or "save"
'			 [recordOfQuestion:String] 儲存的結果字串, 用來做比對是否有選取
'回傳值    ：回傳 1.html tag 或 2.儲存的字串
'牽涉數據表： 
'歷程(作者/日期/原因) ：[Beson] 2011/12/14 Created
'*****************************************************************************************/
  Dim resData : resData = ""
  '找出這問題所有選項
	sql = "SELECT qqil.qqi_number,qqil.qqi_description,qqil.qqi_answer "
	sql = sql & "FROM questionary_list ql "
	sql = sql & "LEFT JOIN questionary_question_list  qql ON qql.ql_number = ql.ql_number "
	sql = sql & "LEFT JOIN questionary_question_item_list qqil ON qqil.qql_number = qql.qql_number "
	sql = sql & "WHERE ql.ql_number="&qlNumber&" AND qqil.qqi_description IS NOT NULL "
	sql = sql & "AND qqil.qqi_parent_number = '"& parent &"' ORDER BY qqil.qqi_sort_sn"
  rsm = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)

  if (Ubound(rsm) > 0) then
	if actionType = "show" then resData =resData&"<div id=""level_"&divLevel&""" style=""padding:0 0 0 "&(divLevel-1)*10&"px;"">"
	for valueOfInput=0 to Ubound(rsm,2)
	  if actionType = "show" then
		if instrrev(recordOfQuestion, ","&rsm(2,valueOfInput)&",", -1 , 1)>0 then isChecked=" checked" else isChecked=""
		if isnup then isDisable = " disabled=true"
		resData = resData & "<span style=""padding:0 1px 0 1px;""><input type="""&typeOfInput(divLevel)&""" name=""level_"&valueOfClassmate&"_"&valueOfQuestion&"_"&divLevel&"_"&parent&""" id=""level_"&valueOfClassmate&"_"&valueOfQuestion&"_"&divLevel&"_"&parent&""" value="""&rsm(2,valueOfInput) &""""&isChecked&isDisable&"  />&nbsp;" & rsm(1,valueOfInput)&"&nbsp;&nbsp;</span>"
	  elseif actionType = "save" then
		if valueOfInput =0 then resData = resData & request("level_"&valueOfClassmate&"_"&valueOfQuestion&"_"&divLevel&"_"&parent)&"," '因input tag都是同名字, 所以取一個就好了.
	  end if
	  resData = resData & extendTree(qlNumber,rsm(0,valueOfInput),divLevel+1,typeOfInput,valueOfClassmate,valueOfQuestion,actionType,recordOfQuestion)
	next
	if actionType = "show" then resData = resData & "</div>"

  end if

  extendTree = resData
end function

'判斷是否為RegularSession '非大會堂 & 1:1
Set objInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
sql = " SELECT  ClassStatusId "
sql = sql & " FROM    member.ClassRecord(NOLOCK) "
sql = sql & " WHERE ClassRecordValid = 1 "
sql = sql & "	AND ( CONVERT(VARCHAR(8), member.ClassRecord.ClassStartDate, 112) "
sql = sql & "	      + RIGHT('00'+ ISNULL(CONVERT(VARCHAR(2), member.ClassRecord.ClassStartTime),''), 2) "
sql = sql & "	      + RIGHT('000'+ ISNULL(CONVERT(VARCHAR(3), member.ClassRecord.ClassRoomNumber),''), 3) =  '"&SessionSn&"' ) "
sql = sql & "	AND LobbySessionId = 0 "
'sql = sql & "	AND ClassTypeId NOT IN (1,7) "
intSqlWriteResult = objInfo.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
if not objInfo.eof then
	isRegularSession = true
end if
objInfo.Close
Set objInfo = nothing

%>
<div class="new-feedback">
    <!--session desc START--><!--課程簡介-->
<%
		'此節課的課程資訊(教材名稱、上課時間、顧問名稱、課程等級、教材簡述)
		Set objInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
		sql = " SELECT  con_basic.basic_fname + ' ' + con_basic.basic_lname AS consultant , "
		sql = sql & "         CONVERT(NVARCHAR, member.ClassRecord.ClassStartDate, 111) AS Attend_Date , "
		sql = sql & "         member.ClassRecord.ClassStartTime AS Attend_Sestime , "
		'sql = sql & "         material.LTitle , "
		'sql = sql & "         material.LD , "
        sql = sql & "         material.Title AS LTitle , "
		sql = sql & "         abcjr.dbo.lesson_description_temp.description AS LD , "
		sql = sql & "         member.ClassRecord.ProfileLearnLevel AS Attend_Level , "
		sql = sql & "         material.Course , "
		sql = sql & "         con_basic.Con_Sn "
		sql = sql & " FROM    member.ClassRecord WITH ( NOLOCK ) "
		'sql = sql & "         LEFT OUTER JOIN material WITH ( NOLOCK ) ON member.ClassRecord.ClassMaterialNumber = material.Course "
        sql = sql & "         LEFT OUTER JOIN abcjr.dbo.lesson_plan_Temp AS material WITH ( NOLOCK ) ON member.ClassRecord.ClassMaterialNumber = material.Course "
        sql = sql & "         LEFT OUTER JOIN abcjr.dbo.lesson_description_temp WITH ( NOLOCK ) ON material.Course = abcjr.dbo.lesson_description_temp.course "
		sql = sql & "         LEFT OUTER JOIN con_basic WITH ( NOLOCK ) ON member.ClassRecord.ClassConsultantId = con_basic.Con_Sn "
		sql = sql & "         LEFT JOIN client_basic WITH ( NOLOCK ) ON client_basic.sn = member.ClassRecord.client_sn "
		sql = sql & " WHERE   ( client_sn = "&intClientSn&" ) "
		sql = sql & "         AND ( CONVERT(VARCHAR(8), member.ClassRecord.ClassStartDate, 112)"
		sql = sql & "               + RIGHT('00'+ ISNULL(CONVERT(VARCHAR(2), member.ClassRecord.ClassStartTime),''), 2) "
		sql = sql & "               + RIGHT('000'+ ISNULL(CONVERT(VARCHAR(3), member.ClassRecord.ClassRoomNumber),''), 3) = '"&SessionSn&"' ) "
		'response.write sql

		intSqlWriteResult = objInfo.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
		if not objInfo.eof then
			consultant = objInfo("consultant")
			attendDateTime = objInfo("Attend_Date") &" "& objInfo("Attend_Sestime")& ":30"
			LTitle = objInfo("LTitle")
			LD = objInfo("LD")
			Attend_Level = objInfo("Attend_Level")
			Course = objInfo("Course")
			con_sn = objInfo("con_sn")
		end if
		objInfo.Close
		Set objInfo = nothing
		if isEmptyOrNull(Attend_Level) or Attend_Level > 12 then Attend_Level=1
%>
<%
		isnup = false		'記錄是否沒有上課   'false 表示可填寫問卷
		if false = bolClassAttend then
			isnup = false
		end if
		'記錄只要日期不是今天的話，就不可以填寫
		if attendDateTime<>"" then
			if(datediff("d",attendDateTime ,date)>1 or (datediff("d",attendDateTime ,date)=1 and hour(now())>9) )  then isnup = true
		else
			isnup = false
		end if

		'測試用
		if "1" = Request("TestMode") then
			isnup = false
		end if
%>
    <div class="session-desc">
    <div class="top1"><%=strLangSessionFeedback00061%><!--課程簡介--></div>
    <div class="box1">
    <table cellpadding="0" cellspacing="0" class="se-desc">
    <tr><td class="left"><%=strLangSessionFeedback00062%><!--教材名稱：--></td>
    <td class="right"><%=LTitle%></td>
    </tr>
    <tr>
      <td class="left"><%=strLangSessionFeedback00063%><!--上課時間：--></td>
      <td class="right"><%=attendDateTime%></td>
    </tr>
    <tr>
      <td class="left"><%=strLangSessionFeedback00064%><!--顧問名稱：--></td>
      <td class="right"><%=consultant%></td>
    </tr>
    <tr>
      <td class="left"><%=strLangSessionFeedback00065%><!--課程等級：--></td>
      <td class="right"><img src="/images/sessionFeedback/feedback_LV<%=Attend_Level%>.gif"></td>
    </tr>
    <tr>
      <td class="left"><%=strLangSessionFeedback00066%><!--教材簡述：--></td>
      <td class="right"><%=LD%></td>
    </tr>
    </table>
    <div align="right">
		<!--單字預習-->
      <input type="button" name="button" id="button" value="<%=strLangSessionFeedback00067%>" class="btn-feedback" onClick="window.open('class_vocabulary.asp?mtl=<%=Course%>')">
    </div>
    </div>
    </div>
    <!--session desc END-->


    <!--Consultant Comment (Only for 1tox) START--><!--顧問課後評鑑-->
<%
		'判斷1對1,1對多的顧問課後評鑑, 大會堂是沒有顧問課後評鑑的
		Dim bolHaveConsultantEvaluation : bolHaveConsultantEvaluation = false
		Set objInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
		'20100203 pipi 工單號碼：GM09092303 ----補填評鑑---- start -----------
		sql = " "
		sql = sql & " SELECT  ISNULL(participation, 0) participation , "
		sql = sql & "         ISNULL(pronunciation, 0) pronunciation , "
		sql = sql & "         ISNULL(comprehension, 0) comprehension , "
		sql = sql & "         ISNULL(creativity, 0) creativity , "
		sql = sql & "         ISNULL(facility, 0) facility , "
		sql = sql & "         ISNULL(c_comments, '') AS c_comments "
		sql = sql & " FROM    ( SELECT    session_sn , "
		sql = sql & "                     client_sn , "
		sql = sql & "                     participation , "
		sql = sql & "                     pronunciation , "
		sql = sql & "                     comprehension , "
		sql = sql & "                     creativity , "
		sql = sql & "                     facility , "
		sql = sql & "                     c_comments "
		sql = sql & "           FROM      Client_Progress_Report "
		sql = sql & "           WHERE     session_sn = '"&SessionSn&"' "
		sql = sql & "           AND client_sn = '"&intClientSn&"' "
		sql = sql & "           UNION ALL "
		sql = sql & "           SELECT    session_sn , "
		sql = sql & "                     client_sn , "
		sql = sql & "                     participation , "
		sql = sql & "                     pronunciation , "
		sql = sql & "                     comprehension , "
		sql = sql & "                     creativity , "
		sql = sql & "                     facility , "
		sql = sql & "                     c_comments "
		sql = sql & "           FROM      Client_Progress_Report_history "
		sql = sql & "           WHERE     session_sn = '"&SessionSn&"' "
		sql = sql & "           AND client_sn = '"&intClientSn&"' "
		sql = sql & "         ) AS ClientProgressReport "
		sql = sql & " LEFT JOIN member.Profile ON ClientProgressReport.client_sn = member.Profile.client_sn "
		'response.write "<br><br>" & sql & "<br><br>"
		intSqlWriteResult = objInfo.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
		if not objInfo.eof then
			participation = objInfo("participation")
			pronunciation = objInfo("pronunciation")
			comprehension = objInfo("comprehension")
			creativity = objInfo("creativity")
			facility = objInfo("facility")
			c_comments = replace(objInfo("c_comments"),chr(13),"<br>")
			bolHaveConsultantEvaluation = true
		end if
		objInfo.Close
		Set objInfo = nothing
		ConsultantEvaluation = Array("","Needs to participate more","Good","Great","Excellent","Outstanding")

		if (bolHaveConsultantEvaluation) then
%>
    <div class="session-desc">
    <div class="top1"><%=strLangSessionFeedback00068%><!--顧問課後評鑑--></div>
    <div class="box1">
    	<table cellpadding="0" cellspacing="0" class="con-comment">
        <tr>
        <td class="left" style="width:150px">
        <p><%=strLangSessionFeedback00069%><!--參與度：--><span><%=ConsultantEvaluation(participation)%></span></p>
        <p><%=strLangSessionFeedback00070%><!--發　音：--><span><%=ConsultantEvaluation(pronunciation)%></span></p>
        <p><%=strLangSessionFeedback00071%><!--理解力：--><span><%=ConsultantEvaluation(comprehension)%></span></p>
        <p><%=strLangSessionFeedback00072%><!--創造力：--><span><%=ConsultantEvaluation(creativity)%></span></p>
        <p><%=strLangSessionFeedback00073%><!--流暢度：--><span><%=ConsultantEvaluation(facility)%></span></p>
        </td>
        <td class="right" valign="top">
			<p><%=strLangSessionFeedback00074%><!--顧問評語：--></p>
			<%
			leftCommentSpace = 0
			rightCommentSpace = 0
			leftCommentSpace = instr(c_comments," ")
			if (instrrev(c_comments," ",-1,1)) > 0 then
				rightCommentSpace = len(c_comments) - instrrev(c_comments," ",-1,1)
			end if
			if leftCommentSpace > 0 then
				if leftCommentSpace > 50 then
					divStyle = "style=""width:350px;position:absolute;word-wrap:break-word;"""
				elseif (rightCommentSpace > 50) then
					divStyle = "style=""width:350px;position:absolute;word-wrap:break-word;"""
				else
					divStyle = ""
				end if
			else
				divStyle = "style=""width:350px;position:absolute;word-wrap:break-word;"""
			end if
			%>
			<div <%=divStyle%>><%=c_comments%></div>
		</td>
        </tr>
        </table>
    </div>
    </div>
<%
		end if
%>
    <!--Consultant Comment (Only for 1tox) END-->
    
    
    <!--After Session START-->
    <div class="session-desc" id="ConsultantEvaluationCollection">
    <div class="top1"><%=strLangSessionFeedback00075%><!--課後評比-->
		<div class="clearall" style="display:none;top:5px">
			<img style="cursor:hand;" id="ClearConsultantEvaluation" src="/images/sessionFeedback/icon-clear.gif"/>
		</div>
	</div>
	<table cellpadding="0" cellspacing="0" class="rating">
    <tr>
<%
		'顧問平均評比
		Set objInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
		sql = " SELECT CONVERT(DECIMAL(3, 1), AVG(consultant_points + 0.0)) AS consultant "
		sql = sql & " FROM   ( SELECT TOP ( 10 ) consultant_points FROM    member.ClassRecord "
		sql = sql & "                 INNER JOIN dbo.client_session_evaluation ON member.ClassRecord.client_sn = dbo.client_session_evaluation.client_sn "
		sql = sql & "                 AND ( CONVERT(VARCHAR(8), member.ClassRecord.ClassStartDate, 112) "
		sql = sql & "                 + RIGHT('00' + ISNULL(CONVERT(VARCHAR(2), member.ClassRecord.ClassStartTime), ''), 2) "
		sql = sql & "                 + RIGHT('000' + ISNULL(CONVERT(VARCHAR(3), member.ClassRecord.ClassRoomNumber),''), 3) = dbo.client_session_evaluation.session_sn ) "
		sql = sql & " WHERE   ( member.ClassRecord.ClassConsultantId = '"&con_sn&"' ) "
		sql = sql & "         AND  client_session_evaluation.session_sn <= '"&SessionSn&"' "
		sql = sql & " ) AS recentConsultantPointOfThisCustomer "
		'response.write "<br><br>" & sql & "<br><br>"
		intSqlWriteResult = objInfo.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
		consultant = 0
		if not objInfo.eof then
			if not isEmptyOrNull(objInfo("consultant")) then
				consultant = objInfo("consultant")
			end if
		end if
		objInfo.Close
		Set objInfo = nothing
		if CDbl(consultant) >= 8 then
			consultant=CDbl(consultant)
		else
			consultant=8
		end if
		if formatnumber(consultant,0) = 10 then scoreImg = 9 else scoreImg = formatnumber(consultant,0)
%>
    <td class="left"><%=strLangSessionFeedback00083%><!--顧問平均評比--></td>
    <td class="right">
    <div class="tutor-score<%=scoreImg%>"><%=formatnumber(consultant,1)%></div><!--分數8開頭的css就套tutor-score8...以此類推-->
    </td>
    </tr>
<%
		Set objClassmate = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
		sql="select Consultant_Points,speed,distribution,tskill,behavior,Materials_Points,difficulty,requirement,overall_Points,connection,suggestions,compliment from client_session_evaluation with (nolock) where session_sn='"&SessionSn&"' and client_sn=" & intClientSn
		intSqlWriteResult = objClassmate.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
		comments = ""
		requirement = ""
		if not objClassmate.eof then
			comments = comments&"#,"&objClassmate("Consultant_Points")&","
			comments = comments&"#,"&objClassmate("speed")&","
			comments = comments&"#,"&objClassmate("distribution")&","
			comments = comments&"#,"&objClassmate("tskill")&","
			comments = comments&"#,"&objClassmate("behavior")&","
			comments = comments&"#,"&objClassmate("Materials_Points")&","
			comments = comments&"#,"&objClassmate("difficulty")&","
			comments = comments&"#,"&objClassmate("requirement")&","
			comments = comments&"#,"&objClassmate("overall_Points")&","
			comments = comments&"#,"&objClassmate("connection")&",#"
			if instrrev(objClassmate("requirement"), "|", -1 , 1)>0 then
				requirement = split(objClassmate("requirement"), "|", -1, 1)
				requirementValue = requirement(1)
			end if
			suggestions = objClassmate("suggestions")
			compliment = objClassmate("compliment")
			isnup = true '如果已做過評鑑就不能再做評鑑
		end if
		objClassmate.close
		Set objClassmate = nothing
		
		if not isEmptyOrNull(comments) then comments = split(comments,"#")
		Set objClassmate = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
		sql = "SELECT qql_number,qql_question FROM questionary_question_list WHERE ql_number='"&qlNumberOfEvaluation&"' ORDER BY ql_sort_sn"
		intSqlWriteResult = objClassmate.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
		valueOfQuestion = 1
		while not objClassmate.eof
			'顧問教學品質,教材品質,通訊品質
			if objClassmate("qql_question") = strLangSessionFeedback00024 then imgIdName = "teacherscore"
			if objClassmate("qql_question") = strLangSessionFeedback00003 then imgIdName = "mertirialscore"
			if objClassmate("qql_question") = strLangSessionFeedback00004 then imgIdName = "netscore"
			if objClassmate("qql_question") = strLangSessionFeedback00024 or objClassmate("qql_question") = strLangSessionFeedback00003 or objClassmate("qql_question") = strLangSessionFeedback00004 then
				mainQuestionAddTag = "<img src=""/images/sessionFeedback/icon-feedback-info.png"" hspace=""2"" >"
				addId = " id="&imgIdName
				tagClass=" ra-title"
			else
				mainQuestionAddTag = ""
				addId = ""
				tagClass=""
			end if
%>
    <tr>
      <td class="left<%=tagClass%>"<%=addId%>><%=mainQuestionAddTag%><%=objClassmate("qql_question")%></td>
      <td class="right<%=tagClass%>">
<%
			if not isEmptyOrNull(comments) then commentItem = comments(valueOfQuestion) else commentItem=""
			response.write extendTree(qlNumberOfEvaluation,objClassmate("qql_number"),1,Array("0","radio","checkbox","radio"),1,valueOfQuestion,"show",commentItem)
%>
	  </td>
    </tr>
<%
			valueOfQuestion = valueOfQuestion + 1
			objClassmate.movenext
		wend
		objClassmate.Close
		Set objClassmate = nothing
%>
    </table>
    </div>
    <!--After Session END-->
    
    <!--Classmate (Only for RegularSession) START--><!--其他同學上課狀況-->
<%
		'如果是大會堂不顯示其他同學上課狀況
		if (isRegularSession) then
			Set objInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
			sql = " SELECT  ClassRecord.client_sn , "
			sql = sql & "         ISNULL(member.Profile.ProfileFirstName, '') + ' ' + LEFT(ISNULL(member.Profile.ProfileLastName, ''), 1) AS cname , "
			sql = sql & "         classmate_session_feedback.condition , "
			sql = sql & "         classmate_session_feedback.performance "
			sql = sql & " FROM    member.ClassRecord "
			sql = sql & "         LEFT OUTER JOIN classmate_session_feedback ON ClassRecord.client_sn = classmate_session_feedback.client_sn "
			sql = sql & "                                                       AND member.ClassRecord.session_sn = dbo.classmate_session_feedback.session_sn "
			sql = sql & "                                                       AND classmate_session_feedback.inputer IN ( "&intClientSn&" ) "
			sql = sql & "         LEFT JOIN member.Profile ON member.Profile.client_sn = member.ClassRecord.client_sn "
			sql = sql & " WHERE   ( ClassRecordValid = 1 ) "
			sql = sql & "         AND member.ClassRecord.ClassStatusId = 3 " '該同學有出席課程
			sql = sql & "         AND ( member.ClassRecord.session_sn = '"&SessionSn&"' ) "
			sql = sql & "         AND ( NOT ( ClassRecord.client_sn IN ( "&intClientSn&" ) )) "
			sql = sql & " ORDER BY ClassRecord.client_sn "
			'response.write sql
			intSqlWriteResult = objInfo.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
			valueOfClassmate = 1
%>
    <div class="session-desc" id="StudentEvaluationCollection">
    <div class="top1"><%=strLangSessionFeedback00076%><!--其他同學上課狀況-->
		<div class="clearall" style="display:none;top:5px">
			<img style="cursor:hand;" id="ClearStudentEvaluation" src="/images/sessionFeedback/icon-clear.gif"/>
		</div>
	</div>
	<table cellpadding="0" cellspacing="0" class="classmate">
<%
			while not objInfo.eof
				if not isEmptyOrNull(objInfo("condition")) then	condition = objInfo("condition") else condition = ""
				if not isEmptyOrNull(objInfo("performance")) then performance=objInfo("performance") else performance = ""
%>
    <tr>
        <td class="left bg"><%=objInfo("cname")%></td>
        <td class="right bg">
        <!--問題 START-->
<%
				Set objClassmate = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
				sql = "SELECT qql_number,qql_question FROM questionary_question_list WHERE ql_number='"&qlNumberOfClassmate&"' ORDER BY ql_sort_sn"
				'response.write sql
				intSqlWriteResult = objClassmate.excuteSqlStatementEachQuick(sql, CONST_TUTORABC_R_CONN)
				valueOfQuestion = 1
				while not objClassmate.eof
%>
        <div class="environment">
        <h1><%=objClassmate("qql_question")%></h1>
<%
					if valueOfQuestion = 1 then response.write extendTree(qlNumberOfClassmate,objClassmate("qql_number"),1,Array("0","radio","checkbox","radio"),valueOfClassmate,valueOfQuestion,"show",condition)
					if valueOfQuestion = 2 then response.write extendTree(qlNumberOfClassmate,objClassmate("qql_number"),1,Array("0","radio","checkbox","radio"),valueOfClassmate,valueOfQuestion,"show",performance)
%>
        </div>
<%
					valueOfQuestion = valueOfQuestion + 1
					objClassmate.movenext
				wend
				objClassmate.Close
				Set objClassmate = nothing
%>
        <!--問題 END-->
        </td>
    </tr>
<%
				valueOfClassmate = valueOfClassmate + 1
				objInfo.movenext
			wend
%>
	</table>
	</div>
<%
			objInfo.Close
			Set objInfo = nothing
		end if
%>
    <!--Classmate (Only for RegularSession) END-->
    
    <!--FeedBack START-->
    <div class="session-desc">
    <div class="top1"><%=strLangSessionFeedback00084%><!--對顧問的建議與讚美--></div>
    <div class="box2">
    <%=strLangSessionFeedback00085%><!--建議--><br /><textarea name="suggestions" id="textarea" cols="45" rows="5" class="textarea-fb" <%if (isnup) then response.write " disabled=true"%>><%if suggestions<>"" then response.write suggestions%></textarea>
    <%=strLangSessionFeedback00086%><!--讚美--><br /><textarea name="compliment" id="textarea" cols="45" rows="5" class="textarea-fb" <%if (isnup) then response.write " disabled=true"%>><%if compliment<>"" then response.write compliment%></textarea>
	<%if not (isnup) then %>
	<input type="checkbox" name="C1" value="ON" id="C1"><%=strLangSessionFeedback00090%><label id="HDDL" style="visibility:hidden;"> <%=strLangSessionFeedback00091%>
	<%
	mystring = split(strLangSessionFeedback00092, "|", -1, 1)
	%>
	<select size="1" name="ctime" style="font-family: Arial; font-size: 8pt">
	<option value="0"><%=mystring(0)%></option>
	<option value="1"><%=mystring(1)%></option>
	<option value="2"><%=mystring(2)%></option>
	<option value="3"><%=mystring(3)%></option>
	</select>
	</label>
	<%end if%>
    </div>
    </div>
    <!--FeedBack Next-->
	<!--送出評鑑-->
    <div align="center"><input type="submit" name="button" id="button" value="<%=strLangSessionFeedback00087%>" class="btn-feedback" <%if (isnup) then response.write " disabled=true"%>></div>
</div>
<!--Beson 20111213 新版課後評鑑 End-->
        </form>
    </div>
    <!--內容 end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<script type="text/javascript">
$(function() {
	//網頁從資料庫load值的時候，將有check的radio秀出來 開始
	$(".new-feedback").find("#level_1").each(function () {
		tmp_level_2 = $(this).find("#level_2");
        if ($($(this).children()[2]).find(":checked").length == 1) {
			tmp_level_2.hide();
			if(tmp_level_2.length == 2) {
				$(tmp_level_2[1]).show();
			}
			else {
				tmp_level_2.show();
			}
        }
        else {
            tmp_level_2.hide();
			if ($($(this).children()[0]).find(":checked").length == 1) {
				if(tmp_level_2.length == 2) {
					$(tmp_level_2[0]).show();
				}
				else
				{
					tmp_level_2.hide();
				}
			}
        }
	});
	//網頁從資料庫load值的時候，將有check的radio秀出來 結束
	
	//需求面的輸入欄位顯示 開始
	var tmp_level_1_7 = $($(".new-feedback").find("#level_1")[7]);
	tmp_level_1_7.append("<input type='text' name='rn' size='7' <%if (isnup) then response.write " disabled=true"%>/ value='<%=requirementValue%>'>");
	tmp_level_1_7.find('input[name="rn"]').attr("disabled","false");
	//需求面的輸入欄位顯示 結束
	
<%
	'如果不是Regular Session,則[通訊狀態]的[其他學員環境吵雜] 開始
	if not (isRegularSession) then
%>
	var Level_2 = $(".new-feedback").find("#level_2");
	$(Level_2[4]).find("span:last").hide()
<%
	end if
	'如果不是Regular Session,則[通訊狀態]的[其他學員環境吵雜] 結束
%>

<%
	'如果沒有上課, 不顯示評鑑部分 開始
	if bolClassAttend = true then
%>
	var strHtmlTag = "<div class=\"title\"><%=strLangSessionFeedback00025%></div>";
	strHtmlTag = strHtmlTag + "<div align=\"right\">";
	strHtmlTag = strHtmlTag + "<input type=\"button\" name=\"button\" id=\"button\" value=\"<%=strLangSessionFeedback00067%>\" class=\"btn-feedback\" onClick=\"window.open('class_vocabulary.asp?mtl=<%=Course%>')\">";
	strHtmlTag = strHtmlTag + "</div>"
	$(".new-feedback").html(strHtmlTag);
<%
	end if
	'如果沒有上課, 不顯示評鑑部分 結束
%>
	//學生互評鑑部分,第1題,問題選項第三層有勾選才enable,反之disable 開始
	var Level_1 = $(".new-feedback").find("#level_1");	
	if (Level_1.length > 10) {
		for (var i=10;i<=Level_1.length;i=i+2) {
			$(Level_1[i]).find("#level_3").find(":radio").attr("disabled",true);
			$(Level_1[i]).find("#level_2").find(":checkbox").click(function() {
				if ($(this).attr("checked") == true) {
					$(this).parent().next().find(":radio").attr("disabled",false);
				}
				else {
					$(this).parent().next().find(":radio").attr("disabled",true);
					$(this).parent().next().find(":radio").attr("checked",false);
				}
			});
		}
	}
	//學生互評鑑部分,第1題,問題選項第三層有勾選才enable,反之disable 結束

	//清空課後評比 開始
	$("#ClearConsultantEvaluation").click(function() {
		if (confirm("<%=strLangSessionFeedback00093&strLangSessionFeedback00075&strLangSessionFeedback00094%>"))
		{
			$("#ConsultantEvaluationCollection").find(":checked").attr("checked",false);
			$("#ConsultantEvaluationCollection").find(":text").val("");
			$("#ConsultantEvaluationCollection").find(":text").attr("disabled",true);
		}
	});
	//清空課後評比 結束
	
	//清空其他同學上課狀況 開始
	$("#ClearStudentEvaluation").click(function() {
		if (confirm("<%=strLangSessionFeedback00093&strLangSessionFeedback00076&strLangSessionFeedback00094%>"))
		{
			$("#StudentEvaluationCollection").find(":checked").attr("checked",false);
			$("#StudentEvaluationCollection").find("#level_3").find(":radio").attr("disabled",true);
		}
	});
	//清空其他同學上課狀況 結束

<%	
	if not (isnup) then
%>
	$("#ConsultantEvaluationCollection").find(".clearall").show();
	$("#StudentEvaluationCollection").find(".clearall").show();
<%
	end if
%>
});
</script>

<iframe src="refresh.html" frameborder="0" width="10" height="10"></iframe>
<!-- pop up start -->
	<!--Consultant Rating pop START-->
    <div id="ConsultantRating" class="rating-pop-desc local-rating-con" style="display:none">
    <div class="pop-desc-top" style="padding-left:10px;"></div>
            <div class="pop-desc-body">
            <div class="pop-rating-title"><%=strLangSessionFeedback00026 %></div><!--顧問教學品質 評比說明-->
                    <div class="pop-rating-area">
                    <span>10<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00027 %><br /><!-- 非常滿意，傑出的教學品質-->
                    &nbsp;&nbsp;<span>9<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00028 %><br /><!-- 很滿意，讓我的英文進步很多-->
                    &nbsp;&nbsp;<span>8<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00029 %><br /><!-- 滿意，讓我的英文有所成長-->
                    &nbsp;&nbsp;<span>7<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00030 %><br /><!-- 尚可，略有學習效果-->
                    &nbsp;&nbsp;<span>6<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00031 %><br /><!-- 普通，學習效果可接受-->
                    </div>
                                
                    <div class="pop-rating-area">
                    <span>5<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00032 %><br /><!-- 不甚滿意，學習效果不顯著-->
                    <span>4<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00033 %><br /><!--  有些不滿意，學習效果待加強 -->
                    <span>3<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00034 %><br /><!-- 不滿意，學習效果欠佳-->
                    <span>2<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00035 %><br /><!-- 非常不滿意，學習效果極差 -->
                    <span>1<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00036 %><!-- 完全無法接受，毫無學習效果-->
                    </div>
                                
                    <div class="clear"></div>
            </div>
    </div>
    <!--Consultant Rating pop END-->
                
    <!--Material Rating pop START-->
    <div id="MaterialRating" class="rating-pop-desc local-rating-matr" style="display:none;">
    <div class="pop-desc-top" style="padding-left:35px;"></div>
            <div class="pop-desc-body">
            <div class="pop-rating-title"><%=strLangSessionFeedback00037 %></div><!--教材品質 評比說明-->
                    <div class="pop-rating-area">
                    <span>10<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00038 %><br /><!-- 非常滿意，教材內容完全符合需求和興趣-->
                    &nbsp;&nbsp;<span>9<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00039 %><br /><!--  很滿意，教材內容對我很有幫助 -->
                    &nbsp;&nbsp;<span>8<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00040 %><br /><!--  滿意，教材內容實用且生動有趣 -->
                    &nbsp;&nbsp;<span>7<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00041 %><br /><!-- 尚可，教材內容適中且符合需求 -->
                    &nbsp;&nbsp;<span>6<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00042 %><br /><!-- 普通，教材內容尚符合需求 -->
                    </div>
                                
                    <div class="pop-rating-area">
                    <span>5<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00043 %><br /><!-- 不甚滿意，勉強可接受-->
                    <span>4<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00044 %><br /><!-- 有些不滿意，教材內容略為簡單或太難 -->
                    <span>3<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00045 %><br /><!-- 不滿意，教材內容不太符合我的需求-->
                    <span>2<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00046 %><br /><!-- 非常不滿意，根本不符合我的需求-->
                    <span>1<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00047 %><!-- 完全不滿意，對教材內容毫無興趣-->
                    </div>
                                
                    <div class="clear"></div>
            </div>
    </div>
    <!--Material Rating pop END-->
                
    <!--Net Rating pop START-->
    <div id="NetRating" class="rating-pop-desc local-rating-net"  style="display:none;">
    <div class="pop-desc-top" style="padding-left:88px;"></div>
            <div class="pop-desc-body">
            <div class="pop-rating-title"><%=strLangSessionFeedback00048 %></div><!--通訊品質 評比說明-->
                    <div class="pop-rating-area">
                    <span>10<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00049 %><br /><!-- 非常滿意，跟實際面對面上課一樣-->
                    &nbsp;&nbsp;<span>9<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00050 %><br /><!-- 很滿意，聲音和影像十分清晰 -->
                    &nbsp;&nbsp;<span>8<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00051 %><br /><!--  滿意，聲音及影像無延遲或斷訊 -->
                    &nbsp;&nbsp;<span>7<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00052 %><br /><!-- 尚可，通訊品質不錯，可正常上課-->
                    &nbsp;&nbsp;<span>6<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00053 %><br /><!-- 普通，聲音及影像略有延遲-->
                    </div>
                                
                    <div class="pop-rating-area">
                    <span>5<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00054 %><br /><!-- 不甚滿意，聲音及影像略有斷訊-->
                    <span>4<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00055 %><br /><!-- 有些不滿意，聲音及影像有延遲斷訊-->
                    <span>3<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00056 %><br /><!-- 不滿意，聲音及影像常延遲斷訊，影響上課效果-->
                    <span>2<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00057 %><br /><!-- 非常不滿意，聲音及影像持續延遲斷訊，上課效果差-->
                    <span>1<%=strLangSessionFeedback00080 %></span><%=strLangSessionFeedback00058 %><!-- 完全無法接受，根本無法上課-->
                    </div>
                                
                    <div class="clear"></div>
            </div>
    </div>
    <!--Net Rating pop END-->
<!-- pop up end -->
