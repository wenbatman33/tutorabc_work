<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/lib/include/vipabc_csp_include.asp"-->
<% dim bolDebugMode : bolDebugMode = false 
str_now_use_product_sn = int_product_sn
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
%>
<link href="/css/learning_review.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/plugin/jquery.blockUI.js"></script>
<link type="text/css" href="/lib/javascript/JQuery/css/redmond/jquery-ui-1.8.2.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>
<script type="text/javascript">
<!--
//將測驗卷檔案讀取近來
function loadHomeWork(session_sn, intClientSn, p_index) {
    document.getElementById("a_conunt_" + p_index).style.display = "none";
    document.getElementById("img_count_loading_" + p_index).style.display = "";
    $.ajax({
        type: "POST",
        url: "myhomework_write.asp",
        data: { "client_sn": intClientSn, "session_sn": session_sn },
        beforeSend: function () {
            $.blockUI({ message: "<h1>Please Wait...</h1>" });
        },
        error: function (xhr) { alert("此課程為測試課程，請選擇正式課程") },
        success: function (msg) {
            $(".main_membox").html(msg);
        },
        complete: function () {
            $.unblockUI();
        }
    });
    //$(".main_membox").load("myhomework_write.asp",{"client_sn":intClientSn,"session_sn":session_sn},function(){$.blockUI();})
}

function viewScore(session_sn, intClientSn) {
    $.ajax({
        type: "POST",
        url: "ajax_myhomework_complete.asp",
        data: { "session_sn": session_sn, "client_sn": intClientSn },
        beforeSend: function () {
            $.blockUI({ message: "<h1>Please Wait...</h1>" });
        },
        error: function (xhr) { alert("此課程為測試課程，請選擇正式課程"); },
        success: function (msg) {
            //alert(msg);
            $(".main_membox").html(msg)
        },
        complete: function () {
            $.unblockUI();
        }
    });
}

function popupJoinnetPlayback(str_video_path) {
    window.open(str_video_path);
}

/// <summary>
/// 20121009 阿捨新增 換頁
/// 換頁載入
/// </summary>
function ChangePage(intPage) {
    window.location = "learning_review.asp?page=" + intPage;
}

/// <summary>
/// 20121214 阿捨新增 
/// 打開新的 dialog for 系統訊息
/// 用燈箱顯示課程資訊 確認觀看錄影檔
/// </summary>
function OpenMsg(strRecordName, intClientSn, intRecordType, intCostSession, intRecordSn, intMaterialSn, intVideoType, intControlType) {
   
    //================ 將原燈箱UI的關閉鈕拿掉 Start ================
    $(".ui-dialog-titlebar").removeClass().empty();
    //================ 將原燈箱UI的關閉鈕拿掉 End ================
    $("#dialog-form").dialog("option", "resizable", false);
    $("#dialog-form").dialog("option", "draggable", false);
    $("#record_name").attr("value", strRecordName);
    $("#client_sn").attr("value", intClientSn);
    $("#record_type").attr("value", intRecordType);
    $("#cost_session").attr("value", intCostSession);
    $("#record_sn").attr("value", intRecordSn);
    $("#material_sn").attr("value", intMaterialSn);
    $("#video_type").attr("value", intVideoType);

    //判斷是否時間內重覆觀看之錄影檔
    $.ajax({
        url: "/program/member/ajaxAddRecordList.asp",
        cache: false,
        data: {
            client_sn: intClientSn,
            cost_session: intCostSession,
            record_sn: intRecordSn,
            material_sn: intMaterialSn,
            video_type: intVideoType,
            method_type: 2
        },
        success: function (result) {
            if ( 1 == result ) { //表時間內觀看為免費
                $("#cost_session").val("0");
                $("#btn").attr("value", "免费回顾");
            }
        },
        error: function (result) {
            //alert("error:"+result);
        }
    });

    //若為舊產品, 直接開啟並記錄
    if ( 1 == intControlType ) 
    {
        OpenRecord();
    }
    else 
    {
        //燈箱打開 設定是否可調整大小/拖曳
        $("#dialog-form").dialog("open"); 
    }
}

/// <summary>
/// 20121214 阿捨新增 
/// 關閉新的 dialog
/// </summary>
function CloseMsg() 
{
    $("#dialog").dialog("destroy");
    $("#dialog-form").dialog("close"); //做完動作再關閉燈箱
}

/// <summary>
/// 20121214 阿捨新增 
/// 呼叫打開錄影檔
/// </summary>
function OpenRecord()
{
    var strRecordName = $("#record_name").val();
    var intClientSn = $("#client_sn").val();
    var intRecordType = $("#record_type").val();
    var intCostSession = $("#cost_session").val();
    var intRecordSn = $("#record_sn").val();
    var intMaterialSn = $("#material_sn").val();
    var intVideoType = $("#video_type").val();

    //寫入錄影檔觀看紀錄
    $.ajax({
        url: "/program/member/ajaxAddRecordList.asp",
        cache: false,
        data: {
            client_sn: intClientSn,
            cost_session: intCostSession,
            record_sn: intRecordSn,
            material_sn: intMaterialSn,
            video_type: intVideoType,
            method_type: 1
        },
        success: function (result) {

        },
        error: function (result) {
            //alert("error:"+result);
        }
    }); 
	CloseMsg();
    if ( intRecordType == 1 ) //joinnet
    {
        popupJoinnetPlayback(strRecordName);
    }
    //else if ( intRecordType == 2 ) //tutormeet
    else //tutormeet
    {
        popupTutorMeetPlayback(strRecordName, intClientSn, "vipabc");
    }
}

$(document).ready(function()
{
    //預設關閉
    $("#dialog").dialog("destroy");

    $("#dialog-form").dialog({
        closeText: "hide",
        autoOpen: false,
        height: 170, //燈箱大小
        width: 260,
        modal: true,
        //closeOnEscape false按esc不會跳出
		closeOnEscape: false,
        close: function () {
            //CloseMsg();
        }
    });
});
//-->
</script>

<div id="dialog-form" style="display:none;background-color:#fcf9f0;border: solid 1px #fe9d00;">
    <div id="box">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td height="47" align="center" valign="bottom">
                    <span class="STYLE4" lang="zh">是否观看此复习录影档</span>
                </td>
            </tr>
            <tr>
                <td height="42" align="center">
                    <span class="STYLE6" lang="zh">系统将扣除您</span><span class="STYLE1" id="cost_session">0.25</span><span class="STYLE6" lang="zh">课时</span>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="24%">
                            <input type="hidden" id="record_name"  />
                            <input type="hidden" id="client_sn"  />
                            <input type="hidden" id="record_type"  />
                            <input type="hidden" id="cost_session"  />
                            <input type="hidden" id="record_sn"  />
                            <input type="hidden" id="material_sn"  />
                            <input type="hidden" id="video_type"  />
                            </td>
                            <td width="34%" lang="zh">
                                <input name="button" type="submit" class="btn" id="button" value="确定" onclick="OpenRecord();"/>
                            </td>
                            <td width="42%" lang="zh">
                                <input name="button2" type="reset" class="btn2" id="button2" value="取消" onclick="CloseMsg();"/>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
</div>
<!--end .dialog-form-->

<!--內容start-->
<div class="main_membox">
    <!--上版位課程start-->
    <!--#include virtual="/lib/include/html/include_learning.asp"-->
    <!--上版位課程end-->
<% 
dim bolTestCase : bolTestCase = true '測試狀態

dim arrSqlResult : arrSqlResult = null
dim strSql : strSql = ""
dim strSubSql : strSubSql = ""
dim intI : intI = 0
dim arrParam : arrParam = null

dim intTotal : intTotal = 0

'取出課程相關資訊
dim intClientSn : intClientSn = request("client_sn") 
if ( isEmptyOrNull(intClientSn) ) then
	intClientSn = session("client_sn")
end if

''20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493
'20121112 阿捨新增 套餐產品改版
'改讀/program/member/reservation_class/getComboProductData.asp

strSql = " SELECT COUNT(client_attend_list_2.sn) AS total /* 0 */ " 
strSql = strSql & " FROM  "
strSql = strSql & " ( "
strSql = strSql & " SELECT client_attend_list_1.sn, client_attend_list_1.client_sn, client_attend_list_1.session_sn, client_attend_list_1.stype, client_attend_list_1.consultant, client_attend_list_1.mtel "  
strSql = strSql & " FROM " 
strSql = strSql & " ( "
strSql = strSql & " SELECT sn, client_sn, session_sn, attend_livesession_types AS stype, attend_consultant AS consultant, attend_mtl_1 AS mtel " 
strSql = strSql & " FROM client_attend_list " 
strSql = strSql & " WHERE client_sn = "& intClientSn &" AND refund IN (0,1) "    
strSql = strSql & " AND attend_mtl_1 IS NOT NULL AND valid = 1 "
strSql = strSql & " AND session_sn IS NOT NULL "
strSql = strSql & " ) AS client_attend_list_1 "
strSql = strSql & " LEFT OUTER JOIN cfg_session_points " 
strSql = strSql & " ON client_attend_list_1.stype = cfg_session_points.sn " 
strSql = strSql & " LEFT OUTER JOIN con_basic ON client_attend_list_1.consultant = con_basic.con_sn  "
strSql = strSql & " LEFT OUTER JOIN material ON client_attend_list_1.mtel = material.course "
strSql = strSql & " ) AS client_attend_list_2 "
strSql = strSql & " LEFT OUTER JOIN client_session_evaluation "
strSql = strSql & " ON client_attend_list_2.client_sn = client_session_evaluation.client_sn "
strSql = strSql & " AND client_attend_list_2.session_sn = client_session_evaluation.session_sn " 
strSql = strSql & " LEFT OUTER JOIN client_homework "
strSql = strSql & " ON client_attend_list_2.session_sn = client_homework.session_sn " 
strSql = strSql & " AND  client_attend_list_2.client_sn = client_homework.client_sn  " 
arrSqlResult = excuteSqlStatementRead(strSql, "", CONST_VIPABC_R_CONN)
if ( true = bolDebugMode ) then
    response.Write("錯誤訊息Sql for 讀取課程記錄總筆數資訊:" & g_str_sql_statement_for_debug & "<br>")
end if
if ( isSelectQuerySuccess(arrSqlResult) ) then
	'如果有資料
	if ( Ubound(arrSqlResult) >= 0 ) then
       intTotal = arrSqlResult(0,0)
    end if
end if

if ( true = bolDebugMode ) then
    response.Write "intTotal:" & intTotal & "<br />"
end if

dim intPageCount : intPageCount = 10 '每頁筆數
dim intNowPage : intNowPage = getRequest("page", CONST_DECODE_ESCAPE) '目前頁面
dim intTopNum : intTopNum = 0 '開頭筆數
dim intBottomNum : intBottomNum = 0 '結尾筆數
dim intTotalPage : intTotalPage = 0 '總頁數
intTotalPage = ( intTotal / intPageCount ) '總頁數無條件進位
if ( InStr(intTotalPage,".") = 0 ) then
	intTotalPage = intTotalPage 
else 
	intTotalPage = sCInt(Left(sCStr(intTotalPage),InStr(1,sCStr(intTotalPage),".")-1))+ 1 
end if

if ( isEmptyOrNull(intNowPage) ) then
    intNowPage = 1 
end if

intBottomNum = intNowPage * intPageCount
intTopNum = intBottomNum - intPageCount + 1

if ( true = bolDebugMode ) then
    response.Write "intTopNum:" & intTopNum & "<br />"
    response.Write "intBottomNum:" & intBottomNum & "<br />"
    response.Write "intTotalPage:" & intTotalPage & "<br />"
end if

%>
<!--list start-->
<div class="hwork">
    <div class="w">
        <% 
        '20121207 阿捨新增 套餐產品 沒有錄影檔權限
        if ( true = bolComboProduct AND false = bolComboHavingRecord ) then
        else
        %>
        <span class="org">贴心提醒</span><span class="gray">：请点选课程日期并点选教材名称后方<img src="/images/images_cn/video.gif" width="14" height="9" class="m-left5 m-right5">即可进行在线观看！<br>
            <!--录影档下载扣点后，于7天内重复下载相同录影档时，不会被再次扣点！-->
        <% end if %>
        </span>
    </div>
    <div class="wlist">
        <div>
            <li class="title1">日期</li>
            <li class="title2">教材名称</li>
            <li class="title3">成绩</li>
            <div class="clear">
            </div>
        </div>
<%
arrParam = Array(intTopNum, intBottomNum)

strSql = " SELECT * FROM ( " '分頁語法
strSql = strSql & " SELECT client_attend_list_2.sn /* 0 */ " 
strSql = strSql & " ,client_attend_list_2.client_sn /* 1 */ "
strSql = strSql & " ,client_attend_list_2.mtel /* 2 */"
strSql = strSql & " ,client_attend_list_2.session_sn /* 3 */ " 
strSql = strSql & " ,ISNULL(client_attend_list_2.situation, 1) AS situation /* 4 */ "
strSql = strSql & " ,client_attend_list_2.title /* 5 */ "
strSql = strSql & " ,client_attend_list_2.cname /* 6 */ "
strSql = strSql & " ,client_attend_list_2.sdate /* 7 */ "
strSql = strSql & " ,client_attend_list_2.stime /* 8 */ "
strSql = strSql & " ,client_attend_list_2.slevel /* 9 */ "
strSql = strSql & " ,client_attend_list_2.sfee /* 10 */ "
strSql = strSql & " ,client_homework.score /* 11 */ "
strSql = strSql & " ,client_homework.anser /* 12 */ "
strSql = strSql & " ,client_homework.voc_sn /* 13 */ "
strSql = strSql & " ,client_homework.inp_date /* 14 */ " 
strSql = strSql & " ,client_session_evaluation.consultant_points AS cpts /* 15 */ " 
strSql = strSql & " ,client_session_evaluation.materials_points AS mpts /* 16 */ "  
strSql = strSql & " ,client_session_evaluation.overall_points AS opts /* 17 */ " 
strSql = strSql & " ,client_session_evaluation.comments AS comment /* 18 */ "
strSql = strSql & " ,ROW_NUMBER() OVER (ORDER BY client_attend_list_2.sdate DESC, client_attend_list_2.stime DESC) AS row " '分頁語法
strSql = strSql & " FROM  "
strSql = strSql & " ( "
strSql = strSql & " SELECT client_attend_list_1.sn, client_attend_list_1.client_sn, client_attend_list_1.mtel, "
strSql = strSql & " client_attend_list_1.session_sn, client_attend_list_1.situation, material.ltitle AS title, " 
strSql = strSql & " con_basic.basic_fname + ' ' + con_basic.basic_lname AS cname, "
strSql = strSql & " client_attend_list_1.sdate, client_attend_list_1.stime, client_attend_list_1.slevel, "  
strSql = strSql & " cfg_session_points.point_fee AS sfee FROM " 
strSql = strSql & " ( "
strSql = strSql & " SELECT sn, client_sn, CONVERT(VARCHAR, attend_date, 111) AS sdate, "
strSql = strSql & " attend_sestime AS stime, attend_mtl_1 AS mtel,  session_sn AS session_sn, "
strSql = strSql & " session_state AS situation, attend_consultant AS consultant, "
strSql = strSql & " attend_level AS slevel,  attend_livesession_types AS stype " 
strSql = strSql & " FROM client_attend_list " 
strSql = strSql & " WHERE client_sn = "& intClientSn &" AND refund IN (0,1) "    
strSql = strSql & " AND attend_mtl_1 IS NOT NULL AND valid = 1 "
strSql = strSql & " AND session_sn IS NOT NULL "
strSql = strSql & " ) AS client_attend_list_1 "
strSql = strSql & " LEFT OUTER JOIN cfg_session_points " 
strSql = strSql & " ON client_attend_list_1.stype = cfg_session_points.sn " 
strSql = strSql & " LEFT OUTER JOIN con_basic ON client_attend_list_1.consultant = con_basic.con_sn  "
strSql = strSql & " LEFT OUTER JOIN material ON client_attend_list_1.mtel = material.course "
strSql = strSql & " ) AS client_attend_list_2 "
strSql = strSql & " LEFT OUTER JOIN client_session_evaluation "
strSql = strSql & " ON client_attend_list_2.client_sn = client_session_evaluation.client_sn "
strSql = strSql & " AND client_attend_list_2.session_sn = client_session_evaluation.session_sn " 
strSql = strSql & " LEFT OUTER JOIN client_homework "
strSql = strSql & " ON client_attend_list_2.session_sn = client_homework.session_sn " 
strSql = strSql & " AND  client_attend_list_2.client_sn = client_homework.client_sn  " 
strSql = strSql & " ) a "
strSql = strSql & " WHERE row BETWEEN @top AND @bottom " '分頁語法
strSql = strSql & " ORDER BY a.sdate DESC, a.stime DESC "
'response.Write("strSql = " & strSql & "<br>")
arrSqlResult = excuteSqlStatementRead(strSql, arrParam, CONST_VIPABC_R_CONN)
if ( true = bolDebugMode ) then
    response.Write("錯誤訊息Sql for 讀取課程記錄分頁資訊:" & g_str_sql_statement_for_debug & "<br>")
end if
'表格開始
'資料庫讀取
if ( isSelectQuerySuccess(arrSqlResult) ) then
	'如果有資料
	if ( Ubound(arrSqlResult) >= 0 ) then
        intTotal = Ubound(arrSqlResult, 2)
		for intI = 0 to intTotal
            dim strSessionDate : strSessionDate = arrSqlResult(7,intI)
            dim intSessionSn : intSessionSn = arrSqlResult(3,intI)
            dim strSessionTitle : strSessionTitle = arrSqlResult(5,intI)
            dim intSessionLevel : intSessionLevel = arrSqlResult(9,intI)
            dim strConsultantName : strConsultantName = arrSqlResult(6,intI)
            dim intSituation : intSituation = arrSqlResult(4,intI)
            dim intScore : intScore = arrSqlResult(11,intI)
            dim strHWInputDate : strHWInputDate = arrSqlResult(14,intI)
            dim strHWAnswer : strHWAnswer = arrSqlResult(12,intI)
            dim strHWVocSn : strHWVocSn = arrSqlResult(13,intI)
            dim intMaterialSn : intMaterialSn = arrSqlResult(2,intI)
%>
            <div class="onebox">
                <div class="date"><%=strSessionDate%></div>
                <div class="desc">
<%
            dim arrRecordResult : arrRecordResult = null
            dim bolJnrRecord : bolJnrRecord = false 'joinnet錄影檔判斷
            dim strOnclickJS : strOnclickJS = "" '開啟錄影檔javascript
            dim intRecordType : intRecordType = 0 '錄影檔類型
            dim strRecordName : strRecordName = "" '錄影檔名稱
            arrSubParam = Array(intSessionSn)
            strSubSql = " SELECT TOP 1 sessionrecord_fileinfo.file_full_name, /* 0 */ sessionrecord_fileinfo.encodestr, /* 1 */ cfg_sessionfile_path.path_server, /* 2 */ cfg_sessionfile_path.path_text, /* 3 */ " 
            strSubSql = strSubSql & " CONVERT(VARCHAR, sessionrecord_fileinfo.Session_date, 111) AS session_date, /* 4 */ sessionrecord_fileinfo.file_full_name /* 5 */, sessionrecord_fileinfo.sn /* 6 */ "
            strSubSql = strSubSql & " FROM sessionrecord_fileinfo LEFT JOIN cfg_sessionfile_path ON sessionrecord_fileinfo.file_path = CONVERT(CHAR, cfg_sessionfile_path.sn) "
            strSubSql = strSubSql & " WHERE (sessionrecord_fileinfo.session_record_number = @session_sn) AND sessionrecord_fileinfo.encodestr IS NOT NULL "
            arrRecordResult = excuteSqlStatementRead(strSubSql, arrSubParam, CONST_VIPABC_R_CONN)
            if ( true = bolDebugMode ) then
                response.Write("錯誤訊息Sql for 讀取錄影檔資訊:" & g_str_sql_statement_for_debug & "<br>")
            end if
            if ( isSelectQuerySuccess(arrRecordResult) ) then
                if ( Ubound(arrRecordResult) >= 0 ) then
                    dim strRecordFileName : strRecordFileName = arrRecordResult(0,0)
	                if ( Instr(strRecordFileName, ".jnr") > 0 ) then
                        bolJnrRecord = true
                    end if
	                if ( Not isEmptyOrNull(strRecordFileName) ) then 
                        if ( true = bolJnrRecord ) then
                            strRecordUrl = "http://" & arrRecordResult(2,0) & "/" & arrRecordResult(3,0) & replace(arrRecordResult(4,0),"/","_") & "/" & arrRecordResult(5,0)
                            'strOnclickJS = "popupJoinnetPlayback('" & strRecordUrl & "')"
                            intRecordType = 1
                            strRecordName = strRecordUrl
                        else
                            'strOnclickJS = "popupTutorMeetPlayback('" & strRecordFileName & "','" & intClientSn & "','vipabc')"
                            intRecordType = 2
                            strRecordName = strRecordFileName
                        end if

                        dim intCostSession : intCostSession = CONST_12_1210_PRODUCT_MY_RECORD_POINTS
                        dim intVideoType : intVideoType = 1
                        dim intRecordSn : intRecordSn = arrRecordResult(6,0)

                        '20121214 阿捨新增 若為新產品 跳出燈箱扣堂
                        '或開啟測試模式時
                        '20121214 阿捨新增 舊產品也要寫入錄影檔觀看紀錄
                        if ( true = bol1On4Session OR true = bolTestCase) then
                            strOnclickJS = "OpenMsg('" & strRecordName & "','" & intClientSn & "','" & intRecordType & "','" & intCostSession & "','" & intRecordSn & "','" & intMaterialSn & "','" & intVideoType & "','0')"
                        else
                            strOnclickJS = "OpenMsg('" & strRecordName & "','" & intClientSn & "','" & intRecordType & "','0','" & intRecordSn & "','" & intMaterialSn & "','" & intVideoType & "','1')"
                        end if

		                '20121207 阿捨新增 套餐產品 沒有錄影檔權限
                        if ( true = bolComboProduct AND false = bolComboHavingRecord ) then
                        else
%>
                    <span class="bold">
                        <%=strSessionTitle%>
                        <a href="javascript:;" style="cursor:hand;text-decoration:none;" onclick="javascript:<%=strOnclickJS%>;">
                            <img src="/images/images_cn/video.gif" width="14" height="9" border="0" /></a>
                    </span>
<% 
                        end if 'if ( true = bolComboProduct AND false = bolComboHavingRecord ) then
                    end if 'if ( Not isEmptyOrNull(strRecordFileName) ) then
                end if 'if ( Ubound(arrRecordResult) >= 0 ) then
            end if 'if ( isSelectQuerySuccess(arrRecordResult) ) then
%>
                    <p>
                    Level:
                    <span class="bold m-left5"><%=intSessionLevel%></span>
                    <span class="m-left10">顾问
                        <span class="bold"><%=strConsultantName%></span>
                    </span>
                    </p>
                </div>
            <div class="count">
<%
            '如果situation 
	        'if Request("language")="zh-tw" then response.Write "缺席" else response.Write "Absence"
	        'response.Write chklangstr(1192)			
	        '缺席做完作業顯示分數
		    if ( isEmptyOrNull(intScore) ) then
                if ( sCInt(intSituation) = 0 ) then
		            response.Write "<a href='javascript:;' style='cursor:hand;text-decoration:none;' onclick=""loadHomeWork('" & intSessionSn & "','" & intClientSn & "');return false"">缺席</a>" 
                else
                    response.Write "<a href='javascript:;' style='cursor:hand;text-decoration:none;' id='a_conunt_" & intI & "' onclick=""loadHomeWork('" & intSessionSn&"','" & intClientSn & "', '" & intI & "');return false"">尚未填写</a>"
		            response.Write "<img id='img_count_loading_" & intI & "' src='/lib/javascript/JQuery/images/ajax-loader.gif'  width='25' height='25' border='0' style='display:none;'/>"
                end if         
		    else
		        if ( strHWInputDate >= cdate("2006/9/9") ) then
				    if ( isEmptyOrNull(strHWAnswer) OR isEmptyOrNull(strHWVocSn) ) then
					    response.Write intScore
				    else
					    response.Write "<a href='javascript:;' style='cursor:hand;text-decoration:none;' onclick=""viewScore('" & intSessionSn & "','" & intClientSn & "');return false"">" & intScore & "</a>"
				    end if
			    else
				    response.Write intScore
			    end if
		    end if 
%>
                </div>
                <div class="clear"></div>
            </div>
<%
		next
	else
		response.Write("無")	
	end if 'if ( Ubound(arrSqlResult) >0 ) then
else
	'錯誤訊息
	'response.Write("錯誤訊息" & arrSqlResult & "<br>")
end if 'if ( isSelectQuerySuccess(arrSqlResult) ) then
'表格結束	
%>
<div class="page">
<%
if ( true = bolDebugMode ) then
    response.Write "intTotalPage:" & intTotalPage & "<br />"
end if
   
Dim intJ : intJ = 0
for intJ = 1 to sCInt(intTotalPage)
	if ( sCInt(intJ) <> sCInt(intNowPage) ) then
		response.Write "<a href='javascript:;' style='cursor:hand;text-decoration:none;' id='page_" & intJ & "' onclick=""ChangePage('" & intJ & "');"">"
	else
		response.Write "&nbsp;&nbsp;&nbsp;<font color='red'><strong>"
	end if
	'response.Write intJ & " "
    response.Write intJ
	if ( sCInt(intJ) <> sCInt(intNowPage) ) then
		response.Write "</a>"
	else
		response.Write "</strong></font>&nbsp;"
	end if
next
%>
</div>
<!-- <div class="page"><a href="#" class="normal">上一頁</a><a href="#">1</a><a href="#">2</a><a href="#">3</a><a href="#" class="normal">下一頁</a></div>-->
    </div>
     <div class="scale">
<%	
'===========================applet=========================start======================================     		
dim intNumber1 : intNumber1 = 0
dim intNumber2 : intNumber2 = 0
dim intNumber3 : intNumber3 = 0
dim intNumber4 : intNumber4 = 0
strSql = " SELECT SUM((CASE ISNULL(score, - 1) WHEN - 1 THEN 1 ELSE 0 END)) AS num1, SUM((CASE score WHEN 100 THEN 1 ELSE 0 END)) as num2, "
strSql = strSql & " SUM((CASE SIGN(score - 60) WHEN - 1 THEN 1 ELSE 0 END)) AS num3, SUM((CASE WHEN score < 100 and score >= 60 THEN 1 ELSE 0 END)) AS num4 " 
strSql = strSql & " FROM client_attend_list INNER JOIN material ON client_attend_list.attend_mtl_1 = material.course "
strSql = strSql & " LEFT OUTER JOIN con_basic ON client_attend_list.attend_consultant = con_basic.con_sn " 
strSql = strSql & " LEFT OUTER JOIN client_homework ON client_attend_list.session_sn = client_homework.session_sn AND client_attend_list.client_sn = client_homework.client_sn "
strSql = strSql & " WHERE (client_attend_list.client_sn = " & intClientSn & ") AND (client_attend_list.valid = 1) AND (client_attend_list.attend_consultant IS NOT NULL) AND (client_attend_list.attend_mtl_1 IS NOT NULL) "
arrSqlResult = excutesqlStatementReadQuick(strSql, CONST_VIPABC_RW_CONN)		
if ( isSelectQuerySuccess(arrSqlResult) ) then
	if ( Ubound(arrSqlResult) > 0 ) then
		intNumber1 = arrSqlResult(0,0)
		intNumber2 = arrSqlResult(1,0)
		intNumber3 = arrSqlResult(2,0)
		intNumber4 = arrSqlResult(3,0)
	else
		intNumber1 = 0
		intNumber2 = 0
		intNumber3 = 0
		intNumber4 = 0
	end if
else			
end if         
%>
    </div>
<%
'20121207 阿捨新增 套餐產品 沒有錄影檔權限
if ( true = bolComboProduct AND false = bolComboHavingRecord ) then
else
%>
<iframe scrolling="no" frameborder="0" style="border: 0px none white; width: 100%; height: 300px; overflow: hidden;" src="PieChart.asp?w1=<%=server.urlencode(getWord("NOT_COMPLETE_NUMBERS"))%>&w2=<%=server.urlencode(getWord("FULL_MARKS"))%>&w3=<%=server.urlencode(getWord("FAIL_NUMBERS"))%>&w4=<%=server.urlencode(getWord("PASS_BUT_NOT_FULL_MARKS"))%>&n1=<%=intNumber1%>&n2=<%=intNumber2%>&n3=<%=intNumber3%>&n4=<%=intNumber4%>">
</iframe>
<% end if %>
        </div>
        <!--list end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->