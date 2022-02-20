<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--

bolVIPABCJR = session("VIPABCJR") '是否為VIPABCJR客戶
'檢查是否為VIPABCJR客戶
if ( isEmptyOrNull(bolVIPABCJR) ) then
   bolVIPABCJR = isVIPABCJRClient(intclient_sn, CONST_VIPABC_RW_CONN)
end if

Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")   
' 判斷是否 做過喇叭測試 及 網速測試
' 若是 則可以跳著測試功能
' ' ' Set obj_conn_fms = Server.CreateObject("ADODB.Connection")
' ' ' obj_conn_fms.connectionString = tutormeet_connectionString			
' ' ' obj_conn_fms.open	

' ' ' set rs=Server.CreateObject("ADODB.recordset")
' ' ' if not isEmptyOrNull(intclient_sn) then
' ' ' rs.Open " SELECT mic_setting , volume_setting FROM fms_client_audio_setting_info WHERE client_sn = "&intclient_sn  , obj_conn_fms , 1, 1
' ' ' result1 = rs.RecordCount
' ' ' end if
' ' ' 'strUpdateMeterial = " SELECT mic_setting , volume_setting, * FROM fms_client_audio_setting_info WHERE client_sn = "&intclient_sn
			
' ' ' 'obj_conn_fms.Execute strUpdateMeterial
' ' ' obj_conn_fms.close
' ' ' set obj_conn_fms = nothing	


'sql = " SELECT mic_setting , volume_setting, * FROM fms_client_audio_setting_info WHERE client_sn = @client_sn "
'result2 = objPCInfoInsert.excuteSqlStatementEach(strSql, Array(intclient_sn), CONST_TUTORABC_R_CONN)     '有傳入參數用
sql2 = " SELECT * FROM EnvironmentTestRecord WHERE clientSn = @client_sn and (finishTest='Y' or finishTest='E') "
result2 = objPCInfoInsert.excuteSqlStatementEach(sql2, Array(intclient_sn), CONST_TUTORABC_R_CONN)     '有傳入參數用
aa= objPCInfoInsert.RecordCount
if  objPCInfoInsert.RecordCount>0 then 'result2>0 and esult1>0
	Response.Cookies("firstTest").Expires = date - 1
	response.cookies("firstTest") = "N"
	Response.Cookies("firstTest").Expires = date + 1
else
	Response.Cookies("firstTest").Expires = date - 1
	response.cookies("firstTest") = "Y"
	Response.Cookies("firstTest").Expires = date + 1
end if
Set envirService = new EnvironmentTestService
call envirService.setEnvironmentTestInfo(objPCInfoInsert)
status = envirService.getSomeOneStatusIsActiveY(objPCInfoInsert)

'檢查是否做完過首測重測
Dim haveTestOk : haveTestOk = envirService.checkEnvironmentHaveTestOk(objPCInfoInsert)
session("environmentHaveTestOk")=haveTestOk

objPCInfoInsert.close                               '關閉物件
Set objPCInfoInsert = nothing                       '設定物件為空
'查詢客戶預約首冊測試時間---結束---

if haveTestOk ="N" and not isEmptyOrNull( status(0) ) then
	response.redirect "/"+status(1)
end if 
%>
<!--內容start-->
<div id="menber_membox">
			<div align="center">
				<!--  ------------->
				<div class="new_guide_start">
					<div class='page_title_4'><h2 class='page_title_h2'>新手上路</h2></div>
	<div class="box_shadow" style="background-position:-80px 0px;"></div> 
				<div class="start_bg">
				
				<h1>欢迎您加入VIPABC 学习行列！<%session("bolFirstTestQuestionaryhaveTestOk")%></h1>
				<h2>为了确保您未来的上课学习品质，上课前请您先完成个人化偏好设定就可以开始订课了！</h2>
				<div id="steps">
				<!-- test -->
				<% if getSession("environmentHaveTestOk", CONST_DECODE_NO)="Y"  then %>
				<div class="r clearfix">
					<div class="number">1.</div>
					<div class="r_r_01" style="line-height:200%"><a class="title"  href="guide_EnvirStart.asp">测试上课电脑设备</a><span class="time"></span></div>
					<div class="r_r_02 clearfix"><span class="info">为了确保您未来上课的电脑环境品质</span><div class="btn_go" style="display:block;"><a href="guide_EnvirStart.asp">前往</a></div></div>
				</div>
				<% else %>
				<div class="r clearfix">
					<div class="number">1.</div>
					<div class="r_r_01" style="line-height:200%"><a class="title" href="guide_EnvirStart.asp">测试上课电脑设备</a><span class="time">(3mins)</span></div>
					<div class="r_r_02 clearfix"><span class="info">为了确保您未来上课的电脑环境品质</span><div class="btn_go" style="display:block;"><a href="guide_EnvirStart.asp">前往</a></div></div>
				</div>
				<% end if%>
				<!-- dcgs -->
                <% 
                dim strDCGSUrl : strDCGSUrl = ""
                if ( true = bolVIPABCJR ) then
                    strDCGSUrl = "/program/aspx/AspxLogin.asp?url_id=5"
                else
                    strDCGSUrl = "/program/member/learning_dcgs.asp"
                end if
                %>
				<% if getSession("bolkompletDcgs", CONST_DECODE_NO)="1"  then %>
				<div class="r clearfix">
					<div class="number">2.</div>
					<div class="r_r_01" style="line-height:200%"><a class="title" href="<%=strDCGSUrl%>">DCGS学习偏好设定</a><span class="time"></span></div>
					<div class="r_r_02 clearfix"><span class="info">量身订作您的个人化教材</span><div class="btn_done" style="display:block;"><a href="<%=strDCGSUrl%>">檢視</a></div></div>
				</div>
				<% else %>
				<div class="r clearfix">
					<div class="number">2.</div>
					<div class="r_r_01" style="line-height:200%"><a class="title" href="<%=strDCGSUrl%>">DCGS学习偏好设定</a><span class="time">(0.5mins)</span></div>
					<div class="r_r_02 clearfix"><span class="info">量身订作您的个人化教材</span><div class="btn_go" style="display:block;"><a href="<%=strDCGSUrl%>">前往</a></div></div>
				</div>
				<% end if%>
				<!-- 語言程度分析 -->
				<% if getSession("bolComplateLevelTest", CONST_DECODE_NO)="1"  then %>
				<div class="r clearfix">
					<div class="number">3.</div>
					<div class="r_r_01" style="line-height:200%"><a class="title" href="/program/member/learning_level.asp">语言程度分析</a><span class="time"></span></div>
					<div class="r_r_02 clearfix"><span class="info">依您目前的程度为您量身订作难易度适中的教材</span><div class="btn_done" style="display:block;"><a href="/program/member/learning_level.asp">檢視</a></div></div>
				</div>
				<% else %>
				<div class="r clearfix">
					<div class="number">3.</div>
					<div class="r_r_01" style="line-height:200%"><a class="title" href="/program/member/learning_level.asp">语言程度分析</a><span class="time">(20mins)</span></div>
					<div class="r_r_02 clearfix"><span class="info">依您目前的程度为您量身订作难易度适中的教材</span><div class="btn_go" style="display:block;"><a href="/program/member/learning_level.asp">前往</a></div></div>
				</div>
				<% end if%>
				<div class="r clearfix">
					<div class="number">4.</div>
					<div class="r_r_01" style="line-height:200%"><a class="title" href="guide_EnvirVideo1.asp">观看教学影片</a><span class="time">&nbsp;</span></div>
					<div class="r_r_02 clearfix"><span class="info">让您更了解VIPABC的上课方式及功能操作</span><div class="btn_go" style="display:block;"><a href="guide_EnvirVideo1.asp">前往</a></div></div>
				</div>
				</div>

				</div>
				<!--  ------------->
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->