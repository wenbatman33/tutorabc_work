<!--[if lt IE 7.]>
<script defer type="text/javascript" src="/lib/javascript/pngfix.js"></script>
<![endif]-->
<link type="text/css" href="/css/order_check_dialoguebox.css" rel="stylesheet" />
<link type="text/css" href="/lib/javascript/JQuery/css/ui-lightness_white/jquery-ui-1.8.7.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="/program/member/reservation_class/javascript/reservation_class.js"></script>
<script type="text/javascript" src="/program/member/reservation_class/javascript/lobby_class.js"></script>
<script type="text/javascript" language="javascript">
<!--
//自建Javascript-right函數
function Right(str, n)
{
    if (n <= 0)
        return "";
    else if (n > String(str).length)
        return str;
    else {
        var iLen = String(str).length;
        return String(str).substring(iLen, iLen - n);
    }
}

//預設開啟燈箱
var strDialogueBoxOpen = "true";
$(document).ready(function ()
{
	//解決IE與其它瀏覽器 抓到的燈箱 高度不同的問題
	var boxHeight = 450;  //FireFox/Chrome/Safari 
	var isIE = navigator.userAgent.search("MSIE");
	var isIE7 = navigator.userAgent.search("MSIE 7"); //IE8可供用
	if ( (isIE7) > -1 || (isIE) > -1 )
	{
		boxHeight = 850;
	}

	$("#dialog").dialog("destroy");
	$("#check_lobby_session").dialog({
		//設定maxHeight,minHeight:解決燈箱高度在IE瀏覽器不會因大會堂內容多,外框就縮小；內容少,外框變大的問題.
		maxHeight: boxHeight, 
		minHeight: boxHeight,
		autoOpen: false,
		width: 730,
		modal: true,
		//closeOnEscape false按esc不會跳出
		closeOnEscape: false
	});
});

// 開啟燈箱 function
function orderCheckDialogueBox(p_strOrderSessionData, p_strChkNoOrderId, p_strClientEmail, p_strDebug)
{
    //訂課checkbox勾選後
	if ( $("#"+p_strChkNoOrderId).attr("checked") )
	{
		//Ajax傳入訂課時間，確認此時段是否有課程，若有課程，回傳 html 大會堂資訊 (topic +<br> +introduction)，若無課程，則不顯示燈箱。
		if ( p_strOrderSessionData != "" )
		{
			$.ajax({
				type: "post",
				url: "/program/member/reservation_class/include/ajax_dialoguebox_session.asp",
				cache: false,
				data:
				{
					nodate: new Date(),
					strDebug: p_strDebug,
					strOrderSessionData: p_strOrderSessionData,
					strClientEmail: p_strClientEmail
				},
				error: function (xhr)
				{
					if ( "yes" == p_strDebug )
					{
						alert(xhr + "\n" + "error:这个时间大会堂没有开课");
					}
				},
				success: function (result)
				{   
					strDialogueBoxOpen = "true";//strDialogueBoxOpen 判斷燈箱 開或關用
					if ( result == "" )//無大會堂資料 不show燈箱
					{
						strDialogueBoxOpen = "false";
					}
					else//有大會堂資料 show燈箱
					{
						strDialogueBoxOpen = "true";
					}
					
					if ( "yes" == p_strDebug )
					{
						alert("result=" + result);
						alert("strDialogueBoxOpen = " + strDialogueBoxOpen);
						alert("p_strChkNoOrderId = " + p_strChkNoOrderId);                           
					}

					//燈箱 放入大會堂的課程資訊
					$("#div_show_lobby").empty();
					$("#div_show_lobby").append(result);

					//燈箱 顯示課程時間
					var strSessionTime = p_strOrderSessionData.substring(10,12);
					$("#div_lobby_time").empty();
					$("#div_lobby_time").append(strSessionTime + ":30");

					//在1對X(1/多) 的 radio value 值設為"checkbox 的id",作為後續要取消"checkbox勾選"用
					$("#div_radio").empty();
					$("#div_radio").append("<input type='radio' name='input_name_session' id='InputOneToX_Id' value='" + p_strChkNoOrderId + "' />");

					if ( "true" == strDialogueBoxOpen )
					{
						//================ 燈箱 start ================
						$("#check_lobby_session").dialog("open");
						//================ 將原燈箱UI的關閉鈕拿掉 Start ================
						$("a").removeClass("ui-dialog-titlebar-close ui-corner-all ui-state-focus");
						$("span:contains('close')").removeClass().empty();
						//================ 將原燈箱UI的關閉鈕拿掉 End ================
						$("#check_lobby_session").dialog("option", "resizable", false);
						$("#check_lobby_session").dialog("option", "draggable", false);
						//================ 燈箱 End ================
					} 
					else//false == strDialogueBoxOpen
					{
						//正常執行原程式，所以什麼事都不做
					}
				}//success: function (result)
			})//$.ajax({
		}//if( p_strOrderSessionData != "" )
	}//if ( $("#"+p_strChkNoOrderId).attr("checked") )
}//function orderCheckBoxFunction()

//燈箱 按 "取消" 或 X 時，關閉燈箱
function CancelSelectSession()
{
	//取消 原選課頁面的 checkbox
	strBookingCheckboxId = $("#InputOneToX_Id").val();
	$("#" + strBookingCheckboxId).attr("checked", false);

	//關閉燈箱
	$("#check_lobby_session").dialog("close");
	return false;
}

//燈箱 大會堂訂課
function chkLobby()
{
	var obj = document.chkSessionForm;
	var strOrderChkMsg = "";
	var strOrderChkMsg2 = "";
	var intSessionValue = '';
    var strJustCheckRemind = "";
    var strCheckId = $("input:radio[name='input_name_session']:checked").attr('id');         
	
    //若勾選一對x選項
    if ( strCheckId == "InputOneToX_Id" )
	{
		//關閉燈箱
		$("#check_lobby_session").dialog("close");
		alert("亲爱的客户，提醒您订课尚未完成喔! \n请您再次确认预约课程是否正确，点选下方送出键后即可完成订课。");
				
		//勾選 原選課頁面的checkbox
        strBookingCheckboxId = $("#InputOneToX_Id").val();
		$("#" + strBookingCheckboxId).attr("checked", true);
		return false;
	} 
    //若不等於沒勾
	else if ( strCheckId != undefined )
	{
        var  strLobbySessionInfo = $('input[name=input_name_session]:checked').val();
		strInfoData = strLobbySessionInfo.substr(2,16); //ex. 99 2011 01 31 20 3798 (99:大會堂, 20110131:日期, 20:時, 3798:lobby_session sn)
		var strSessionDate = strInfoData.substr(0,4) + "/" + strInfoData.substr(4,2) + "/" + strInfoData.substr(6,2);
		var strSessionStartTime = strInfoData.substr(8,2) + ":30";
		var strSessionEndTime = (parseInt(strInfoData.substr(8,2),10) == 23) ? "00:30" : eval(parseInt(strInfoData.substr(8,2),10)+1) + ":30"; //parseInt('08') bug
        strSessionEndTime = Right("0" + strSessionEndTime,5);
		var intChargeSession = ""

		strOrderChkMsg2 += strSessionDate + ' ' + strSessionStartTime + ' - ' + strSessionEndTime + ' <%="(一般订位 -- 新增订位)" %> '; //( 一般訂位 -- 新增訂位 )<br>
		var strLobbySn = strLobbySessionInfo.substr(12,6);
        //算扣堂
		$.ajax({
			url: '/program/member/reservation_class/include/ajax_check_charge_session.asp',
			data: { lobby_sn: strLobbySn },
			async: false,
			cache: false,
			error: function (xhr)
			{ 
				//alert("錯誤" + xhr.responseText)
			},
			success: function (result)
			{
				intChargeSession = result;
			}
		});
		strOrderChkMsg2 += '，收取 ' + intChargeSession + ' 堂\n';
        intSessionValue = strLobbySessionInfo;
	}//if ( strCheckId == "InputOneToX_Id" )

	strOrderChkMsg = "<%="您新增 预约上课的时间如下:"%> \n\n";

	if ( strOrderChkMsg2 != "" ) 
	{ 
		strOrderChkMsg += strOrderChkMsg2 + "\n\n"; 
	}

	if ( strOrderChkMsg2 != "" ) 
	{
		if ( confirm(strOrderChkMsg) )//確認訂大會堂課程
		{  
			//關閉燈箱
			$("#check_lobby_session").dialog("close");
			g_obj_reservation_lobby_class_ctrl.addNewLobbyClass(strLobbySn, false);
            g_obj_reservation_class_ctrl.submitReservationClass();
            $("#hdn_total_reservation_class_dt").val(strLobbySessionInfo);
            $("#form_reservation_class_module").submit();
		} 
		else//if ( confirm(strOrderChkMsg) )
		{
			return false;
		}//if ( confirm(strOrderChkMsg) )
	}
	else//if (strOrderChkMsg2 != "")
	{
        //先判斷客戶是否勾選不要提醒
	    if ( $('#cbNoRemind').attr('checked') ) 
	    {
		    //勾選不提醒
		    $.ajax({
			    url: '/program/member/reservation_class/include/ajax_lobby_remind.asp?do=noremind',
			    async: false,
			    cache: false,
			    error: function (xhr) 
			    {
				    alert("网页忙录中，请再次操作!");
			    },
			    success: function (result) 
			    {
				    //alert("您好,TutorABC提醒您：系統將不再提醒您大會堂課程\n\n");
				    window.location.reload();
			    }
		    });
	    }
        else
        {
		    alert("您好,VIPABC提醒您：请选择您欲上课的课程\n\n");
        }
		return false;
	}//if (strOrderChkMsg2 != "")
};//function checksubmint()

$("#open_remind").click(function() {
    var strChkMsg = "是否想了解同时段还有哪些大会堂课程"
	if ( confirm(strChkMsg) ) 
	{
		$.ajax({
			url: '/program/member/reservation_class/include/ajax_lobby_remind.asp?do=remind',
			async: false,
			cache: false,
			error: function (xhr) 
			{
				alert("网页忙录中，请再次操作!");
			},
			success: function (result) 
			{
				location.reload(true);
			}
		});
    }
});
//-->
</script>
<%
strFName = getSession("client_fname", CONST_DECODE_NO) 
if ( isEmptyOrNull(strFName) ) then
	strFName ="客户"
end if
%>
<!--灯箱 内容 start-->
<div class="order_check_box ui-dialog-content ui-widget-content" id="check_lobby_session" style="display: none;">
    <div id="altmsg"></div>
    <div class="toppp">
        <div class="ttpp"><img src="/images/reservation_class/dialogue_box/order2010_top.png"></div>
        <div class="ttpp2">&nbsp;</div>
        <div class="ttpp"><a id="x_close_dialog" href="javascript:void(0);" onclick="return CancelSelectSession();">
                <img src="/images/reservation_class/dialogue_box/order2010_close.png"></a></div>
        <div class="f-clear"></div>
    </div>
    <div class="bodyyy">
        <div class="desccc">Dear<%=strFname%>, 太棒了！您选择的同一时间<span class="color_red font_bold" id="div_lobby_time"></span>也有精彩的大会堂课程，快来瞧瞧：</div>
        <div class="title_ssss">&nbsp;</div>
        <div class="ss_descccbox" id="div_show_lobby">
            <!--Ajax (append result) html-->
            <div class="f-clear"></div>
        </div>
        <% 
        'if ( cstr(str_class_type) = "03" ) then '03:一對多 
	    '    strCssName = "title_1to6"
        if ( cstr(str_class_type) = "07" ) then '07:一對一
            strCssName = "title_1to1"
        else
             strCssName = "title_1to6"
        end if
        %>
		<div class="<%=strCssName%>" id="div_radio"><!--append html--></div>
        <div class="f-clear"></div>
        <div class="btnnnnbox">
            <%
                    '检查是否要显示 如果没有全时段就显示 
                        Dim int_run_result
                        Dim obj_sql_opt_read : Set obj_sql_opt_read = server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
                        int_run_result = obj_sql_opt_read.excuteSqlStatementEach("select count(*) as sum from  lobby_dialoguebox where valid=1 and client_email='" & strClientEmail & "'",  null , CONST_TUTORABC_R_CONN)
                        If ( int_run_result <> CONST_FUNC_EXE_SUCCESS ) then
                                'Response.write "SQL for Debug:" & obj_sql_opt_read.FullSqlStatement &"<br>"
                                'Response.write "RESULT:" & obj_sql_opt_read.ExcuteResult
                        end if 
                        if (obj_sql_opt_read("sum") <= 0) then
            %>
            <p id="remind" name="remind" style="color: #aaaaaa; font-size: 12px;"><input type="checkbox" name="cbNoRemind" id="cbNoRemind" />谢谢，暂不需要提醒我大会堂课程</p>
            <p id="remind" name="remind" style="color: #aaaaaa; font-size: 10px;">(本系统在勾选关闭时，一个月后将会再自动开启)</p>
            <% end if %>
            <input type="button" name="chkClsButton" id="chkClsButton" value="确定" class="btnnnn" onclick="return chkLobby();" />
            <input type="button" name="button" id="button" value="取消" class="btnnnn" onclick="return CancelSelectSession();" />
        </div>
    </div>
    <div class="toppp">
        <div class="ttpp"><img src="/images/reservation_class/dialogue_box/order2010_left_under.png"></div>
        <div class="ttpp3"> &nbsp;</div>
        <div class="ttpp"><img src="/images/reservation_class/dialogue_box/order2010_right_under.png"></div>
        <div class="f-clear"></div>
    </div>
</div>
<!--燈箱 內容 end-->

