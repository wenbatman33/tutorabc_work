<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<script type="text/javascript" src="/program/member/reservation_class/javascript/ClassRecordJR.js"></script>
<% 
dim bolDebugMode : bolDebugMode = false 
dim bolTestCase : bolTestCase = true
%>
<!--#include virtual="/program/member/reservation_class/include/getComboProductData.asp"-->
<%
'會有六個session 五個變數
'bolComboProduct, bolComboHaving1On1, bolComboHaving1On3, bolComboHavingLobby, bolComboHavingRecord
if ( true = bolDebugMode ) then
    response.Write "bolComboProduct: " & bolComboProduct & "<br />"
    response.Write "bolComboHaving1On1: " & bolComboHaving1On1 & "<br />"
    response.Write "bolComboHaving1On3: " & bolComboHaving1On3 & "<br />"
    response.Write "bolComboHavingLobby: " & bolComboHavingLobby & "<br />"
    response.Write "bolComboHavingRecord: " & bolComboHavingRecord & "<br />"
end if
%>
<!--#include virtual="/program/class/lobby/CourseUtility.asp" -->
<body lang="zh">
<link href="/css/session_20120807.css" rel="stylesheet" type="text/css" />
<%

dim arrParam : arrParam = null
dim strSql : strSql = "" 'sql
dim objLobbySessionData : objLobbySessionData = null
dim strTotalSession : strTotalSession = "" '剩餘堂數 
dim strCheckCalOrList : strCheckCalOrList = getRequest("type", CONST_DECODE_NO) '確認訂課後 判斷載入頁面
'大會堂 近期預定課程 
'取得客戶編號
if ( Not isEmptyOrNull(g_var_client_sn) ) then
    int_client_sn = g_var_client_sn
end if

'20120316 阿捨新增 無限卡產品判斷	
Dim bolUnlimited : bolUnlimited = false
if ( Not isEmptyOrNull(str_now_use_product_sn) ) then
    if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (str_now_use_product_sn & ",")) > 0 ) then
        bolUnlimited = true 
    end if
end if

'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493     
'20121112 阿捨新增 套餐產品改版
'改讀/program/member/reservation_class/getComboProductData.asp

'20121112 阿捨新增 套餐產品改版 沒有大會堂上限 導回課程選擇頁面
if ( true = bolComboProduct ) then
    if ( true = bolComboHavingLobby ) then
    else
        Response.Redirect "class_choose.asp"
        Response.End
    end if
end if
%>
<!--燈箱載入css-->
<link type="text/css" href="/lib/javascript/JQuery/css/redmond/jquery-ui-1.8.2.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="javascript/reservation_class<%=strVIPABCJR%>.js"></script>
<script type="text/javascript" src="javascript/lobby_class.js"></script>
<div id="conimg" style="display:none"></div>
<div class="oneColFixCtr">
    <div id="container">
        <div id="mainContent">

            <div class='page_title_1'><h2 class='page_title_h2'>Special Session 大会堂</h2></div>
            <div class="box_shadow" style="background-position:-80px 0px;"></div> 
            <div id="session_booking">
                <div class="session_booking_head">
<!--課程列表start-->
<ul id="ul_reservation_class_info" class="class_list32" style="display:none;">
    <li>
        <h1>&nbsp;</h1>
        <h2><%=str_msg_font_date%></h2><!--日期-->
        <h3><%=str_msg_font_time%></h3><!--时间-->
        <h4><%=str_msg_font_consultant%></h4><!--顾问-->
        <h5><%=str_msg_font_lobby_class_list_head_2%></h5><!--主题-->
        <% if ( true = bolUnlimited ) then %>
        <% else %>
        <h2><%=str_msg_font_lobby_class_list_head_3%></h2><!--预计扣除课时数-->
        <% end if %>
    </li>      
</ul>
<!--課程列表end-->
<!--內容start-->
<div id="div_reservation_class_result" class="list_txt" style="display:none;">
    <form id="form_reservation_class_module" method="post"  name = "myform" onsubmit="javascript:return g_obj_reservation_class_ctrl.submitReservationClass();" >
        <%'hdn_reservation_class_count 隱藏值 是為了要紀錄已新增了幾筆課程%>
        <input type="hidden" id="hdn_reservation_class_count" name="hdn_reservation_class_count" value="0" />
        <%'hdn_prev_select_class_date 隱藏值 是為了要紀錄前一次選擇的日期%>
        <input type="hidden" id="hdn_prev_select_class_date" value="" />        
        <%'hdn_total_reservation_class_dt 隱藏值 客戶已預訂的課程日期和時間 YYYYMMDDHH%>
        <input type="hidden" id="hdn_total_reservation_class_dt" name="hdn_total_reservation_class_dt" value="" />
        <%'hdn_order_or_reorder_type: 1=預定 2=重新預定%>
        <input type="hidden" id="hdn_order_or_reorder_type" name="hdn_order_or_reorder_type" value="1" />
        <input type="hidden" id="page" name="page" value="<%=intNowPage%>" />
        <p class="color_blue">
        <span class="ball_blue">●</span><%=str_msg_font_lobby_class%><!--大会堂--></p>
        <p>
        <%=str_msg_font_order_result_1%><!--您此次预订课时共--> <span id="span_reservation_class_count"></span> <%=str_msg_font_order_result_2%><!--次-->
        <span class="m-left20">
        <span class="color_5"><%=str_msg_font_lobby_class_list_head_3%><!--预计扣除课时数--></span>
        <span id="span_reservation_class_total_subtract_session"></span>
        <!--<span class="color_5"><%=str_msg_font_order_result_3%>--><!--可用课时数--><!--</span>
        <span id="span_client_total_session"></span>-->
        </span>
        </p>
    </form>
        <br class="clear" />
</div>
                    <div class="session_booking_head_t">
                    </div>
                    <div class="notice">
                        <div class="notice_t">预约 / 取消课程时请注意</div>
                        <ol>
                            <li>新增预约课程：课程开始的 24 小时前完成</li>
                            <li>取消预约课程：课程开始的 4 小时前完成</li>
                            <li class="red">为提供您学习在线英语演讲的便利，现在「大会堂」课程开放上课前 5 分钟完成订位!</li>
                        </ol>
                    </div>
                    <!-- end. notice -->
                </div>
                <!-- end.session_booking_head -->
            </div>
            <!-- end.session_booking -->

            <div class="session_booking_body">
                <div class="head_bar clearfix">
                    <div class="filter">
                        <select id="lobby_suit_lev" name="lobby_suit_lev">
                            <option value="">适用程度</option>
                            <%
                            dim intI : intI = 0
                            dim intMaxLV : intMaxLV = 12
                            dim strSelected : strSelected = ""
                                               
                            for intI = 1 to intMaxLV
                                if ( intI = intLobbySuitLevel ) then
                                    strSelected = "selected"
                                else
                                    strSelected = ""
                                end if 
                            %>
                            <option value="<%=intI%>" <%=strSelected%>>L<%=intI%></option>
                            <% next %>
                        </select>
                        <!--選擇主題-->
                        <select id="lobby_topic" name="lobby_topic">
                            <option value="">选择主题</option>
                            <%
                            dim strTypeName : strTypeName = ""
                            dim arrTypeName : arrTypeName = null
                            dim intLobbySessionSn : intLobbySessionSn = 0
                            dim strTopicName : strTopicName = ""

					        strSql = " SELECT DISTINCT dbo.lobby_session_type.sn AS sn, dbo.lobby_session_type.type_name_cn "
				            strSql = strSql & " FROM  lobby_session "
					        strSql = strSql & " LEFT JOIN dbo.lobby_session_type ON dbo.lobby_session.lobby_type = dbo.lobby_session_type.sn "
					        strSql = strSql & " WHERE lobby_session.valid = 1 "
					        strSql = strSql & " AND lobby_session.sn <> 2622 "
					        strSql = strSql & " AND lobby_session.session_date BETWEEN CONVERT(VARCHAR(10),GETDATE(),111) AND CONVERT(VARCHAR(10),GETDATE()+7,111) " '七天內的符合條件
                            '20130331 阿捨新增 VIPABCJR判斷 FOR JR大會堂
                            if ( true = bolVIPABCJR ) then
                                 strSql = strSql & " AND lobby_session.BrandType&2 = 2 "
                            else
                                strSql = strSql & " AND lobby_session.BrandType <> 2 "  'type 1:TUTPRABC VIPABC/2:JR/3:all
                            end if 
					        strSql = strSql & " ORDER BY lobby_session_type.type_name_cn "
                            Set objLobbySessionData = excuteSqlStatementReadEach(strSql, "", CONST_VIPABC_R_CONN)
                            if ( true = bolDebugMode ) then
                                response.Write("錯誤訊息Sql for 一周大會堂主題:" & g_str_sql_statement_for_debug & "<br>")
                            end if
                            while not objLobbySessionData.eof
                                strTypeName = objLobbySessionData("type_name_cn")
                                strTopicName = ""

                                if ( Not isEmptyOrNull(strTypeName) ) then
                                    strTypeName = replace(strTypeName, "–","-")
						            arrTypeName = split(strTypeName,"-")
                                    if ( 1 = Ubound(arrTypeName) ) then
                                        strTopicName = arrTypeName(1)
                                    end if
                                end if
                                if ( isEmptyOrNull(strTopicName) ) then
                                    strTopicName = strTypeName
                                end if

                                intLobbySessionSn = objLobbySessionData("sn")
                                if ( sCDbl(intLobbyTopic) = sCDbl(intLobbySessionSn) ) then
                                    strSelected = "selected"
                                else
                                    strSelected = ""
                                end if
                            %>
                            <option value="<%=intLobbySessionSn%>" <%=strSelected%>><%=strTopicName%></option>
                            <%
                                objLobbySessionData.movenext
                            wend
                            %>
                        </select>
                        <!--選擇顧問-->
                        <select id="con_sn" name="con_sn">
                            <option value="">选择顾问</option>
                            <%
                            dim strConsultantName : strConsultantName = ""
                            dim intDataConsultantSn : intDataConsultantSn = ""

						    strSql = " SELECT DISTINCT con_basic.con_sn, con_basic.basic_fname + ' ' + con_basic.basic_lname AS ConsultantName "
						    strSql = strSql & " FROM lobby_session "
						    strSql = strSql & " INNER JOIN con_basic ON lobby_session.consultant = con_basic.con_sn "
						    strSql = strSql & " WHERE lobby_session.valid = 1 AND lobby_session.sn <> 2622 "
						    strSql = strSql & " AND lobby_session.session_date BETWEEN CONVERT(VARCHAR(10),GETDATE(),111) AND CONVERT(VARCHAR(10),GETDATE()+7,111)  " '七天內的符合條件
                           '20130331 阿捨新增 VIPABCJR判斷 FOR JR大會堂
                            if ( true = bolVIPABCJR ) then
                                 strSql = strSql & " AND lobby_session.BrandType&2 = 2 "
                            else
                                strSql = strSql & " AND lobby_session.BrandType <> 2 "  'type 1:TUTPRABC VIPABC/2:JR/3:all
                            end if 
						    strSql = strSql & " ORDER BY ConsultantName "
						    Set objLobbySessionData = excuteSqlStatementReadEach(strSql, "", CONST_VIPABC_R_CONN)
                            if ( true = bolDebugMode ) then
                                response.Write("錯誤訊息Sql for 一周大會堂顧問" & g_str_sql_statement_for_debug & "<br>")
                            end if
                            while not objLobbySessionData.eof
					            strConsultantName = objLobbySessionData("ConsultantName")
                                intDataConsultantSn = objLobbySessionData("con_sn")
                                if ( sCDbl(intDataConsultantSn) = sCDbl(intConsultantSn) ) then
                                    strSelected = "selected"
                                else
                                    strSelected = ""
                                end if
                            %>
                            <option value="<%=intDataConsultantSn%>" <%=strSelected%>><%=strConsultantName%></option>
                            <%
                                objLobbySessionData.movenext
                            wend
                            %>
                        </select>
                        <input type="button" value="查询" class="btn_1"onclick="SearchSumit();" /> <!--for list-->
                    </div>
                    <div class="display_type">
                        <div class="type_btn" value="cal">
                            <a class="display_type_cal" title="切換成行事曆模式" onclick="toCalendar();">行程表</a>
                        </div>
                        <div class="type_btn" value="list">
                            <a class="display_type_list_current" title="切換成列表模式" onclick="toList();">列表</a>
                        </div>
                    </div>
                    <div class="lesson_qty">
                    <%  '若無限卡客戶則不show剩餘堂數
                        if ( true = bolUnlimited ) then 
                        else
                    %>
                        您尚可订<span id="total_session"> <span class="qty red" ><%
                        '20121110 阿捨新增 套餐產品設計 這隻程式只for大會堂剩餘堂數顯示寫死99
                        if ( true = bolComboProduct ) then
                            dim intHavingSession : intHavingSession = getComboProductSession(int_member_client_sn, str_now_use_contract_sn, "99", 1, 0, CONST_VIPABC_R_CONN)
                            '20130102 阿捨新增 套餐產品判斷 for 贈送堂數
                            dim intComboGiftSession : intComboGiftSession = getComboProductGift(str_now_use_contract_sn, "", "99", "", 0, CONST_VIPABC_R_CONN)
                            response.Write (sCDbl(intHavingSession) + sCDbl(intComboGiftSession))
                        else
                            response.Write flt_client_total_session + flt_client_week_session + flt_client_bonus_session
                        end if
                        %></span> </span>堂咨询 <!--id名稱訂課時更新剩餘堂數-->
                    <% end if %>
                    </div>
                </div>
                <!-- end .head_bar -->
                <div class="session_booking_list">
                <!--取得課程列表 Allison 20121002-->
                    <div id="div_lobby_class_info">
                    </div>
                </div>
                <!--end. session_booking_list-->
            </div>
            <!--end. session_booking_body-->
            </form>
            <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
        </div>
        <!--end. mainContent-->
    </div>
    <!--end. container-->
</div>
<!--end. oneColFixCtr-->
<!--燈箱拉到外面 開始-->
<div id="dialog-form" style="display: none;">
<!--为适应播放视频，原div布局改为tabel begin 2013/02/28 Paul-->
   <table border="0" width="97%" height="auto" align="center" cellpadding="0" cellspacing="0">
		<tr style="background-image: none; background-attachment: scroll; background-repeat: repeat; background-position-x: 0%; background-position-y: 0%; background-color: rgb(201, 181, 130); color: #ffffff; font-weight: bolder; height:35px">
			<td  style="padding-left: 10px; width:35%">顾问 / 日期</td>
			<td  style="padding-left: 5px;width:45%">课程主题</td>
			<td width="20%">适用程度</td>
        </tr>
		<tr style="background-color: rgb(250, 247, 238);height:30px" >
			<td style="padding-left: 10px; width: 35%;">
				<span id="session_datetime"></span><!--课程时间-->
			</td>
			<td style="padding-left: 5px; padding-top: 5px;width:45%" >
				<span id="topic"></span><!--课程主题-->
			</td>
			<td style="color: #f59048;">
				<span id="lv"></span><!--level级别-->
			</td>
		</tr>
		<tr style="background-color: rgb(250, 247, 238);">
			<td  style="color: #f59048; font-weight: bolder;width:35%; padding-left: 10px;">
				<span id="consultant_name"></span><!--顾问名称-->
			</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr style="background-color: rgb(250, 247, 238);">
			<td style="width: 35%; padding-left: 10px;">
				<span id="material_pic"></span><!--顾问图片-->
			</td>
			<td style="width: 45%; padding-left: 5px; color: #706d68;" vAlign="top">		
				<span id="desc"></span><!--课程介绍-->
			</td>
			<td>&nbsp;</td>
		</tr>
   </table>
   <table border="0" width="98%" align="center" cellpadding="0" cellspacing="5px">
		<tr style="background-color: rgb(250, 247, 238);">
			<td colSpan="3" style="text-align: center; height: 35px;" >
				<input type="button" id="submit_button_dialog" style="cursor:hand;text-decoration:none;" value="确认订位" class="btn_1" onclick="submitButtonDialog();" />
				<input type="button" id="back" value="关闭" class="btn_1" style="cursor:hand;text-decoration:none;" onclick="CloseMsg();"/> 
			</td>
		</tr>
   </table>
   <!--为适应播放视频，原div布局改为tabel end 2013/02/28 Paul-->
</div><!--end .dialog-form-->
<!--燈箱拉到外面 結束-->

<!--內容end-->
<!--把script搬到最下面執行，因intTotal後面才有值-->
<script type="text/javascript" language="javascript">
<!--
var intTotal = "<%=intTotal%>";
var intTotalPage = "<%=intTotalPage%>";
var intLobbySuitLevel = "<%=intLobbySuitLevel%>";
var intLobbyTopic = "<%=intLobbyTopic%>";
var intConsultantSn = "<%=intConsultantSn%>";
var strCheckCalOrList = "<%=strCheckCalOrList%>"; //查詢預設頁list SearchSubmit  //確認訂課頁面 導頁判斷
//var str_in_four_hours = "" //判斷四小時底色控制
var strClassIDString = ""; //訂課字串

//列表 確認訂課
function submitReservate(){
    document.body.style.cursor="progress";
    var bolCheckOk = g_obj_reservation_class_ctrl.submitReservationClass();
    if ( true == bolCheckOk )
    {
        $('#form_reservation_class_module').submit();
        //重新load 剩餘堂數
        //RefreshSession();
    }
}	

/// <summary>
/// 20121005  小丸子新增 
/// 換頁呼叫ajax載入
/// 傳入參數：頁數；換頁方式1:列表 2:行事曆
/// </summary>
function ChangePage(intPage, intPageType)
{
    //document.body.style.cursor="progress";
    $("#div_lobby_class_info").html("");
	var intLobbySuitLevel = $("#lobby_suit_lev option:selected").val();  //程度
	var intLobbyTopic = $("#lobby_topic option:selected").val();  //主題
	var intConsultantSn = $("#con_sn option:selected").val();  //顧問

    $("#div_lobby_class_info").html("<img id='img_ajax_loader' src='/lib/javascript/JQuery/images/ajax-loader.gif' border='0' />");

    if ( 2 == intPageType ) //若為2為cal換頁 載入cal
    {
        //document.body.style.cursor="default";
        strPageUrl = "/program/member/reservation_class/special_session_cal.asp";
    }
    else //若為1為list換頁 載入list
    {
        //document.body.style.cursor="default";
        strPageUrl = "/program/member/reservation_class/include/ajax_lobby_class_info<%=strVIPABCJR%>.asp";
    }

	$.ajax({    　　　　　　　  
		url: strPageUrl,
		cache: false,
		async: false,
		data:{
			page: intPage, 
			intTotal : intTotal, 
			intTotalPage : intTotalPage,
			lobby_suit_lev : intLobbySuitLevel,
			lobby_topic : intLobbyTopic, 
			con_sn : intConsultantSn
		},　　　　　　　　　　　　　　　　　　
		success: function(result)
        {
			$("#div_lobby_class_info").html(result);
		},
		error: function(result) {
			//alert("error:"+result);
		}
    }); 
    ChangePageChecked();
    $("#ajax_temp").html("");
}

/// <summary>
/// 20121005  小丸子新增 
/// 查詢送出事件
/// 傳入參數：
/// </summary>
function SearchSumit()
{
	var intLobbySuitLevel = $("#lobby_suit_lev option:selected").val(); //程度
	var intLobbyTopic = $("#lobby_topic option:selected").val(); //主題
	var intConsultantSn = $("#con_sn option:selected").val(); //顧問
    
    if ( "list" == strCheckCalOrList ) //判斷欲查詢的頁面為list
    {
        intSearch_Url = "/program/member/reservation_class/include/ajax_lobby_class_info<%=strVIPABCJR%>.asp";
    }
    else //判斷欲查詢的頁面為cal
    {
        intSearch_Url = "/program/member/reservation_class/special_session_cal.asp";
    }
    $("#div_lobby_class_info").html("<img id='img_ajax_loader' src='/lib/javascript/JQuery/images/ajax-loader.gif' border='0' />");

	$.ajax({    　　　　　　　  
		url: intSearch_Url,  //"/program/member/reservation_class/include/ajax_lobby_class_info.asp", 
		cache: false,
		async: true,
		data:{
			intTotal : intTotal, 
			lobby_suit_lev : intLobbySuitLevel,
			lobby_topic : intLobbyTopic, 
			con_sn : intConsultantSn
		},　　　　　　　　　　　　　　　　　　
		success: function(result) {
			$("#div_lobby_class_info").html(result);
		},
		error: function(result) {
			//alert("error:"+result);
		}
    });
}

/// <summary>
/// 20121009  小丸子新增 
/// 切換行事曆模式
/// 傳入參數：
/// </summary>
function toCalendar() 
{
	var intLobbySuitLevel = $("#lobby_suit_lev option:selected").val(); //程度
	var intLobbyTopic = $("#lobby_topic option:selected").val(); //主題
	var intConsultantSn = $("#con_sn option:selected").val(); //顧問
    strCheckCalOrList = "cal";  //切換成cal模式 查詢SearchSubmit會利用此判斷查詢為cal/list  //導回頁面判斷
    $("#div_lobby_class_info").html("<img id='img_ajax_loader' src='/lib/javascript/JQuery/images/ajax-loader.gif' border='0' />"); //loading圖示
    $.ajax({
        url:"/program/member/reservation_class/special_session_cal.asp",
        cache: false,
		async: true,
        data: {
            lobby_suit_lev : intLobbySuitLevel,
            lobby_topic : intLobbyTopic, 
            con_sn : intConsultantSn
        },
        success:function(result) {
            $("#div_lobby_class_info").html(result);
        },
        error:function(result) {
            //alert("error:"+result);
        }
    });
    $(".display_type_cal").attr("className","display_type_cal_hover"); //行事曆變亮
    $(".display_type_list_current").attr("className","display_type_list_hover"); //列表變暗
    $("#ul_reservation_class_info").html("");
    document.myform.action = "/program/member/reservation_class/reservation_class_exe<%=strVIPABCJR%>.asp?session=lobby&type=" + strCheckCalOrList; //導回原訂課完成之頁面
}

/// <summary>
/// 20121009  小丸子新增 
/// 切換列表模式
/// 傳入參數：
/// </summary>
function toList() 
{
	var intLobbySuitLevel = $("#lobby_suit_lev option:selected").val(); //程度
	var intLobbyTopic = $("#lobby_topic option:selected").val(); //主題
	var intConsultantSn = $("#con_sn option:selected").val(); //顧問
    strCheckCalOrList = "list"; //切換成list模式 查詢SearchSumit會利用此判斷該跳去哪頁  //導回頁面判斷
    $("#div_lobby_class_info").html("<img id='img_ajax_loader' src='/lib/javascript/JQuery/images/ajax-loader.gif' border='0' />"); //loading圖示
    $.ajax({
        url:"/program/member/reservation_class/include/ajax_lobby_class_info<%=strVIPABCJR%>.asp",
        cache: false,
		async: true,
        data: {
            lobby_suit_lev : intLobbySuitLevel,
            lobby_topic : intLobbyTopic, 
            con_sn : intConsultantSn
        },
        success:function(result) {
            $("#div_lobby_class_info").html(result);
        },
        error:function(result) {
            //alert("error:"+result);
        }
    });
    $(".display_type_cal_hover").attr("className","display_type_cal"); //行事曆變暗
    $(".display_type_list_hover").attr("className","display_type_list_current"); //列表變亮
    $("#ul_reservation_class_info").html("");
    document.myform.action = "/program/member/reservation_class/reservation_class_exe<%=strVIPABCJR%>.asp?session=lobby&type=" + strCheckCalOrList; //導回原訂課完成之頁面
}

/// <summary>
/// 20121018 小丸子新增
/// 列表取消訂課
/// 傳入參數：checkbox的idName, 課程名稱, 課程日期, 課程時間, 
/// </summary>
function cancelCheck(strSessionIdName, strSessionTopic, strSessionSDate, intSessionSTime, intClientAttendSn, intSessionSn, intPage, strChecked) 
{
    //排除特殊字元 topic
    var strSessionTopic = $("#" + intSessionSn + "_topic").val();
    var strSessionId = document.getElementById(strSessionIdName);
    //var strSessionImgID = strSessionIdName+"_IMG"; //checkbox圖
    //var strCss = $("#"+strSessionImgID).attr('class');
//    if ( "checkbox check_on" == strCss ) {
//        $("#"+strSessionIdName).attr("checked", false);
//        $("#"+strSessionImgID).removeClass("check_on");
//        $("#"+strSessionImgID).addClass("check_off"); 
//    }
//    else
//    {
//        $("#"+strSessionIdName).attr("checked", true);
//        $("#"+strSessionImgID).removeClass("check_off");
//        $("#"+strSessionImgID).addClass("check_on"); 
//    }
    var strMsg = "课程名称："+strSessionTopic.replace("<br>","\n") + "\n" + "课程时间："+ strSessionSDate + " " + intSessionSTime + ":30";
    if ( !strSessionId.checked && "checked" == strChecked )  //原本狀態有訂 現在取消
    {
        //document.body.style.cursor="progress"; //滑鼠loading
        if ( confirm("您取消预约的课程资讯如下：\n" + strMsg +"\n\n注意! 按下确认之后即取消课程!") ) 
        {
            //直接取消課程 for list( intPageType : 1)
            g_obj_reservation_class_ctrl.cancelClass(intClientAttendSn, strSessionSDate, intSessionSTime, intPage, 1, strSessionIdName);
            //strSessionId.checked = false;
            //$("#"+strSessionIdName).attr("checked", false); //移至cancelClass 取消成功才將checkbox取消勾選
            RefreshSession(); //取消課程後 重新load剩餘堂數
        }
        else 
        {
            //strSessionId.checked = true;
            //$("#"+strSessionImgID).removeClass("check_off");
            //$("#"+strSessionImgID).addClass("check_on"); 
            $("#"+strSessionIdName).attr("checked", true);
            //CheckedColor(strSessionIdName); //異動checkbox背景色
        }
        //document.body.style.cursor="default"; //滑鼠loading
    }
    else
    {
        //訂課到暫存區
        g_obj_reservation_lobby_class_ctrl.addNewLobbyClass(intSessionSn, false); 
        CheckedColor(strSessionIdName); //異動checkbox背景色
    }
    CheckCheckBox(strSessionIdName); //判斷異動checkbox function

    //if ( "true"==str_in_four_hours ) //四小時內無法退課
    //{
    //    strClassIDString = strClassIDString.replace("," + strSessionIdName, "");
    //    str_in_four_hours = "";
    //}
}

/// <summary>
/// 20121022 小丸子新增
/// 重整 剩餘堂數
/// 列表： submit_button、cancelCheck()；行事曆：submit_button_dialog 燈箱 訂課取消共用
/// </summary>
function RefreshSession(){
    $.ajax({
        url: "/program/member/reservation_class/refresh_class.asp",
        cache: false,
        async: false,
        success: function (result) {
            $('#total_session').html(result); 
        },
        error: function (result) {
            //alert("error:" + result);
        }
    });
}

function consns(con_sn){
$("#material_pic").html(
  "<object width='320' height='266' classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' codebase='http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,4,5,715' standby='Loading Microsoft Windows Media Player components...' type='application/x-mplayer2'>"+
  "<param name='_ExtentX' value='11233'>"+
  "<param name='_ExtentY' value='8916'>"+
  "<param name='AUTOSTART' value='1'>"+
  "<param name='SHUFFLE' value='0'>"+
  "<param name='PREFETCH' value='0'>"+
  "<param name='NOLABELS' value='0'>"+
  "<param name='SRC' value='http://www.muchtalk.com/consultant/video/" + con_sn + ".wmv'>"+
  "<param name='CONTROLS' value='Imagewindow'>"+
  "<param name='CONSOLE' value='clip1'>"+
  "<param name='LOOP' value='0'>"+
  "<param name='NUMLOOP' value='0'>"+
  "<param name='ShowTracker' value='0'>"+
  "<param name='CENTER' value='0'>"+
  "<param name='MAINTAINASPECT' value='0'>"+
  "<param name='BACKGROUNDCOLOR' value=''>"+
  "<param name='ShowPositionControls' value='0'>"+
  "<param name='ShowAudioControls' value='0'>"+
  "<param name='ShowControls' value='1'>"+
  "<param name='ShowDisplay' value='0'>"+
  "<param name='ShowStatusBar' value='0'>"+
  "<param name='ShowGotoBar' value='0'>"+
  "<param name='ShowCaptioning' value='0'>"+
  "<param name='TransparentAtStart' value='0'>"+
  "<param name='EnableContextMenu' value='0'>"+
  "<embed  src='http://www.muchtalk.com/consultant/video/" + con_sn + ".wmv' width='320' height='266' ShowTracker='false' ShowGotoBar='false' ShowControls='true' ShowDisplay='false' ShowStatusBar='false' ShowCaptioning='false' TransparentAtStart='false' ShowPositionControls='false' ShowAudioControls='false' EnableContextMenu='0' align='middle' prefetch='0' nolabels='0' controls='Imagewindow' console='clip1' numloop='0' maintainaspect='0' name='player'>"+
  "</embed>"+
  "</object>");
  }
  
/// <summary>
/// 20121018 小丸子新增 
/// 打開新的 dialog for 系統訊息
/// 用燈箱顯示課程資訊 確認/取消訂課
/// 傳入參數 : 開課日期yyyy/mm/dd, 開課時間01~23, 顧問名稱, 圖片路徑, 主題, 介紹, 程度, 是否訂課1:訂課 0:為訂課, 課程sn, 已訂位才有值
/// </summary>
function OpenMsg(strSDate, strStime, strConsultantEName, strImageUrl, strSessionTopic1, strSessionIntro1, strTempSessionLv, intSessionStatus, intSessionSn, intClientAttendSn)
{
    //燈箱打開 設定是否可調整大小/拖曳
    $("#dialog-form").dialog("open"); 
    //================ 將原燈箱UI的關閉鈕拿掉 Start ================
	$(".ui-dialog-titlebar").removeClass().empty();
	//================ 將原燈箱UI的關閉鈕拿掉 End ================
    $("#dialog-form").dialog("option", "resizable", false);
    $("#dialog-form").dialog("option", "draggable", true);

    $("#session_datetime").html(strSDate + " <span class='time'>" + strStime + ":30</span");
    $("#consultant_name").html(strConsultantEName);
	
    $("#material_pic").html("<img src=" + strImageUrl + " width='70' />");
    //把特殊字元排除 topic/intro
    var strSessionTopic = $("#" + intSessionSn + "_topic").val();
    var strSessionIntro = $("#" + intSessionSn + "_intro").val();
    $("#topic").html(strSessionTopic);
    $("#desc").html(strSessionIntro);
    $("#lv").html(strTempSessionLv);
    
    $("#session_sn").val(intSessionSn);
    $("#session_status").val(intSessionStatus);
    $("#session_sdate").val(strSDate);
    $("#session_stime").val(strStime);
    $("#client_attend_sn").val(intClientAttendSn); //已訂位才有值

    $("#check_box").html("<input type='checkbox' id='chk_lobby_class_" + intSessionSn + "'><span id='span_lobby_class_date_" + intSessionSn + "'>" + strSDate + "</span><span id='span_lobby_class_time_" + intSessionSn + "'>" + strStime + ":30" + "</span><input type='hidden' id='hdn_lobby_class_con_ename_" + intSessionSn + "' value='" + strConsultantEName + "' /><span id='span_lobby_class_topic_" + intSessionSn + "' style='display:none;'>" + strSessionTopic + "</span>");

    //判斷已訂位
    if ( 1 == intSessionStatus ) 
    {
        $("#submit_button_dialog").attr("value", "取消订位");
        //g_obj_reservation_class_ctrl.removeClassInfo(str_order_class_datetime, g_obj_reservation_class_ctrl.arr_class_type.lobby); //移除暫存區
    }
    else
    {
        $("#submit_button_dialog").attr("value", "确认订位");
        //$("#submit_button_dialog").val("确认订位");
    }
    var strSessionId = document.getElementById("chk_lobby_class_"+intSessionSn);
    //strSessionId.checked = true;
    $("#chk_lobby_class_"+intSessionSn).attr("checked", true);

    //訂課到暫存區
    g_obj_reservation_lobby_class_ctrl.addNewLobbyClass(intSessionSn, false);
}

function OpenMsgS(strSDate, strStime, strConsultantEName, strImageUrl, strSessionTopic1, strSessionIntro1, strTempSessionLv, intSessionStatus, intSessionSn, intClientAttendSn , con_sn)
{
    //燈箱打開 設定是否可調整大小/拖曳
    $("#dialog-form").dialog("open"); 
    //================ 將原燈箱UI的關閉鈕拿掉 Start ================
	$(".ui-dialog-titlebar").removeClass().empty();
	//================ 將原燈箱UI的關閉鈕拿掉 End ================
    $("#dialog-form").dialog("option", "resizable", false);
    $("#dialog-form").dialog("option", "draggable", true);

    $("#session_datetime").html(strSDate + " <span class='time'>" + strStime + ":30</span");
    $("#consultant_name").html(strConsultantEName);
	
	SeachConsultantImg(con_sn);
	var imgurl = $("#conimg").html();
	//顾问图像为空，预设图像
	if(imgurl==0){
	$("#material_pic").html("<img src='" + "/images/no_face_img.jpg" + "' width='320' height='266' onclick='consns(" + con_sn + ")' style='cursor:pointer' />");	
	}else{
	$("#material_pic").html("<img src='/lib/functions/functions_show_consultant_image.asp?con_sn="+con_sn+"&website=<%=CONST_VIPABC_WEBSITE%>' alt='"+ strConsultantEName+"' title='" + strConsultantEName + "' border='0' width='320' height='266' onclick='consns(" + con_sn + ")' style='cursor:pointer' />");	
	}
    //把特殊字元排除 topic/intro
    var strSessionTopic = $("#" + intSessionSn + "_topic").val();
    var strSessionIntro = $("#" + intSessionSn + "_intro").val();
    $("#topic").html(strSessionTopic);
    $("#desc").html(strSessionIntro);
    $("#lv").html(strTempSessionLv);
    $("#session_sn").val(intSessionSn);
    $("#session_status").val(intSessionStatus);
    $("#session_sdate").val(strSDate); 
    $("#session_stime").val(strStime);
    $("#client_attend_sn").val(intClientAttendSn); //已訂位才有值
    $("#check_box").html("<input type='checkbox' id='chk_lobby_class_" + intSessionSn + "'><span id='span_lobby_class_date_" + intSessionSn + "'>" + strSDate + "</span><span id='span_lobby_class_time_" + intSessionSn + "'>" + strStime + ":30" + "</span><input type='hidden' id='hdn_lobby_class_con_ename_" + intSessionSn + "' value='" + strConsultantEName + "' /><span id='span_lobby_class_topic_" + intSessionSn + "' style='display:none;'>" + strSessionTopic + "</span>");

    //判斷已訂位
    if ( 1 == intSessionStatus ) 
    {
        $("#submit_button_dialog").attr("value", "取消订位");
        //g_obj_reservation_class_ctrl.removeClassInfo(str_order_class_datetime, g_obj_reservation_class_ctrl.arr_class_type.lobby); //移除暫存區
    }
    else
    {
        $("#submit_button_dialog").attr("value", "确认订位");
        //$("#submit_button_dialog").val("确认订位");
    }
    var strSessionId = document.getElementById("chk_lobby_class_"+intSessionSn);
    //strSessionId.checked = true;
    $("#chk_lobby_class_"+intSessionSn).attr("checked", true);

    //訂課到暫存區
    g_obj_reservation_lobby_class_ctrl.addNewLobbyClass(intSessionSn, false);
}
//根据顾问con_sn获取顾问图像 20130308 paul
function SeachConsultantImg(con_sn){
    $.ajax({
        url: "/lib/functions/functions_search_consultant_img.asp",
        cache: false,
        async: false,
		data:{"consn": con_sn},
        success: function (result) {
            //alert(result);
			$("#conimg").html(result);
		},
        error: function (result) {
            alert("error:" + result);
        }
    });
	//return result;
}

function GetImgData(result){
var img=result;
alert(img);
}

/// <summary>
/// 20121017 小丸子新增 
/// 關閉新的 dialog
/// </summary>
function CloseMsg() 
{
    var intSessionSn = $("#session_sn").val();
    /*
    var strSessionId = document.getElementById("chk_lobby_class_" + intSessionSn);
    strSessionId.checked = false;
    g_obj_reservation_lobby_class_ctrl.addNewLobbyClass(intSessionSn, false);
    */
    str_class_date = $("#span_lobby_class_date_" + intSessionSn).text(); //開課日期
    str_class_time = $("#span_lobby_class_time_" + intSessionSn).text(); //開課時間
    str_order_class_datetime = str_class_date.replace(new RegExp("/", "gm"), "") + str_class_time.substring(0, 2); //  訂課日期時間 YYYYMMDDHHmm
    g_obj_reservation_class_ctrl.removeClassInfo(str_order_class_datetime, g_obj_reservation_class_ctrl.arr_class_type.lobby); //移除暫存區
    //$("#ul_reservation_class_info").html("");
    $("#dialog").dialog("destroy");
    $("#dialog-form").dialog("close"); //做完動作再關閉燈箱	
	$("#material_pic").html("");
}

/// <summary>
/// 20121024 小丸子新增 
/// 檢查是否有新增訂課(checkbox異動) 可跨頁訂課
/// </summary>
//click checkbox時觸發此funciton 用字串串課程
function CheckCheckBox(strCheckIDName) 
{
    var strSessionId = document.getElementById(strCheckIDName); //取得checkbox ID
    
    if ( strSessionId.checked ) //新增勾選課程
    {
        strClassIDString = strClassIDString + "," + strCheckIDName;
    }
    else //取消勾選課程
    {
        strClassIDString = strClassIDString.replace( "," + strCheckIDName,"");
    }
    
}

/// <summary>
/// 20121024 小丸子新增 
/// 換頁時 有異動課程的checkbox也要同步
/// </summary>
//剛才勾選的課程 換頁再換回此頁也要勾選 放在ChangePage裡
function ChangePageChecked()
{
    var arrClassID = strClassIDString.split(","); //把有異動的字串sn變成陣列存放
    for ( var intI=1; intI < arrClassID.length; intI = intI+1 )  //有異動的sn存成陣列
    {
        var strIdCheck = document.getElementById(arrClassID[intI]); 
        if ( strIdCheck != undefined )
        {
            //$("#"+arrClassID[intI]+"_IMG").removeClass("check_off");
            //$("#"+arrClassID[intI]+"_IMG").addClass("check_on");
            $("#"+arrClassID[intI]).attr("checked", true);
            //strIdCheck.checked = true;  //有在異動字串內的checkbox都要打勾
            CheckedColor(arrClassID[intI]); //異動checkbox背景色
        }
    }
}

/// <summary>
/// 20121026 小丸子新增 
/// checkbox打勾就要變色 取消勾選就變回來
/// </summary>
function CheckedColor(strSessionIdName)
{
    var strSessionId = document.getElementById(strSessionIdName);
    if ( strSessionId.checked ) 
    {
        $("#"+strSessionIdName + "_color" ).addClass( " CheckedColor ");
    }
    else
    {
        $("#"+strSessionIdName + "_color" ).removeClass(" CheckedColor "); //變色
    }
}

$(document).ready(function()
{
    //預設關閉
    $("#dialog").dialog("destroy");

    $("#dialog-form").dialog({
        closeText: "hide",
        autoOpen: false,
        height: 455, //燈箱大小
        width: 900,
        modal: true,
        //closeOnEscape false按esc不會跳出
		closeOnEscape: false,
        close: function () {
            //CloseMsg();
        }
    });

    var str_lobby_date = "<%=trim(request("lobby_date"))%>";
    var str_lobby_time = "<%=trim(request("lobby_time"))%>";
    var str_con_ename = "<%=trim(unescape(request("con_ename")))%>";
    var str_lobby_topic = "<%=trim(unescape(request("lobby_topic")))%>";
    var str_lobby_sn = "<%=trim(request("lobby_sn")) %>"
    var str_now_page_url = "<%=Request.ServerVariables("PATH_INFO")%>";
    var str_contract_end_date = "<%=str_service_end_date%>";    // 合約結束日

    if (str_contract_end_date == "")
    {
        // 合約已到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
        //alert(g_obj_class_data.arr_msg.contract_edate_invalid_0 + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);      
        $("#remind_contract_start_msg").text("<%=getWord("WELCOME_JOIN_VIPABC")%>" + "! " + g_obj_class_data.arr_msg.contract_edate_invalid_0 + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
        $("#img_ajax_loader").hide();
    }
    else
    {
        // ajax 動態載入 搜尋大會堂開課資訊介面
        //$("#div_lobby_class_search").load(g_obj_reservation_lobby_class_ctrl.arr_page_url.ajax_lobby_class_search, {ndate : new Date(), page_url : str_now_page_url, from_other_page_con_sn : str_lobby_sn});
        $("#div_lobby_class_info").html("<img id='img_ajax_loader' src='/lib/javascript/JQuery/images/ajax-loader.gif' border='0' />");

        // ajax 動態載入 確認訂課後載入list/cal頁面判斷
        if ( "list" == strCheckCalOrList ) //判斷確認訂課頁面為list
        {
            toList();
        }
        else //判斷確認訂課頁面為cal
        {
            toCalendar();
        }

        // 從搜索頁來的預定大會堂
        /*
        if (str_lobby_date != "" && str_lobby_time != "" && str_con_ename != "" && str_lobby_topic != "")
        {
            return g_obj_reservation_lobby_class_ctrl.addNewLobbyClass("", true, str_lobby_date, str_lobby_time, str_con_ename, str_lobby_topic, str_lobby_sn);
        }
        */
    }
});
//-->
</script>