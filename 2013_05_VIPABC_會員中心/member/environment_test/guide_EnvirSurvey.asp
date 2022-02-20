<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/program/class/environment_test/FormEnvirSurvey.asp" -->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<%
Dim intclient_sn    : intclient_sn      = getSession("client_sn", CONST_DECODE_NO)      '客戶編號      'SQL資料參數
'沒有客戶SN先導回首頁--開始--

%>
<%
'本頁功能 新版首測重測 Johnny Liu 2012-04-06
'''''''''''''''''''''''''''''''''''''''''''''''
'初始化
Dim envirService , form1
set envirService = new EnvironmentTestService
set form1 = new FormEnvirSurvey
Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")     
'GET
Dim isPassTest
isPassTest = false
if Request.ServerVariables("REQUEST_METHOD") = "GET" then
	'紀錄使用者目前進入 狀態
	call envirService.insertStatus( objPCInfo , "1" , "program/member/environment_test/guide_EnvirSurvey.asp" )
	'紀錄首測重測 進入點  Log Insert
	set logService = new EnvironmentTestServiceLog	
	call logService.AddLogBaseArrCol( objPCInfo , Array("QuestionStep_in") , Array( Now() ) )

	

	call envirService.checkHaveWriteQuestion(objPCInfo , session( "client_sn") )
	while not objPCInfo.EOF 
		isPassTest = true
		if objPCInfo("qql_number")=285 then
			call form1.setField("q285ComputerKind",objPCInfo("qqi_number") )
		elseif objPCInfo("qql_number")=289 then
			call form1.setField("q289ConnectionMetod",objPCInfo("qqi_number") )
		elseif objPCInfo("qql_number")=288 then
			call form1.setField("q288NetCompany",objPCInfo("qqi_number") )
			call form1.setField("q288NetCompany_other",objPCInfo("qqi_comment") )
		elseif objPCInfo("qql_number")=287 then
			call form1.setField("q287NetKind",objPCInfo("qqi_number") )
		elseif objPCInfo("qql_number")=291 then
			call form1.setField("q291NetShare",objPCInfo("qqi_number") )
		end if
		objPCInfo.moveNext()
	wend
end if
'POST
if Request.ServerVariables("REQUEST_METHOD") = "POST" then
	call form1.checkValid()
	if form1.getValid() = true then
		resultxx = envirService.insertQuestionAnswer(objPCInfo , form1)
		if resultxx >= 5 then
			%>
			<script>
				alert('填写成功！');
				location.href='guide_EnvirDownload.asp';
			</script>
			<%
		else
			%>
			<script>alert('Can not insert data！')</script>
			<%
		end if	
	end if
end if
%>
<script src="/program/member/environment_test/js/logService.js"></script>
<!--內容start-->
<div id="menber_membox">
			<div class='page_title_4'><h2 class='page_title_h2'>新手上路</h2></div>
			<div class="box_shadow" style="background-position:-80px 0px;"></div> 
			<div align="center">
				<!-- --------->
				<form name="formX" action="guide_EnvirSurvey.asp" method="post">
					<div class="new_guide_start"><br />
					<div class="new_quide_test">
					<div class="step-box"><img src="/images/environment_test/new_test_step1.png" usemap="#planetmap"></div>
					<%
					Dim firstTest : firstTest =getCookies("firstTest",CONST_DECODE_NO)																
					if firstTest ="N" then
					%>
					<map name="planetmap">
					  <area shape="rect" coords="0,0,140,65" href="guide_EnvirSurvey.asp?" alt="填写电脑环境问卷" />
					  <area shape="rect" coords="140,0,280,65" href="guide_EnvirDownload.asp?" alt="下载上课软件平台" />
					  <area shape="rect" coords="280,0,420,65" href="guide_EnvirNetwork.asp?" alt="网络联机测试" />
					  <area shape="rect" coords="420,0,600,65" href="guide_EnvirVolumn.asp?" alt="耳机&麦克风测试" />
					</map>
					<%end if%>
					<script>
					
					</script>
					<table cellpadding="0" cellspacing="0" class="envi-tab">
					  <%
						Dim i : i=0
					  
						call envirService.selectQuestionList(objPCInfo , 285)
						%><tr><th colspan="2" class="qu"><span class="mustWrite"><%=form1.getErrorField("q285ComputerKind")%></span>1.<%=objPCInfo("qql_question")%></th></tr><%
						
						while not objPCInfo.EOF
						%>
						<%if i mod 2 = 0 then%><tr><%end if%>
						<td class="ansline"><input type="radio" value="<%=objPCInfo("qqi_number")%>" name="q285ComputerKind" <%if objPCInfo("qqi_number")=form1.getField("q285ComputerKind") then response.write("checked") end if%> ><%=objPCInfo("qqi_answer")%></td>
						<%if i mod 2 = 1 then%></tr><%end if%>
						<%																		
							objPCInfo.moveNext()
							i = i+1
					  wend%>																  
					  <!-- ------->
					  <%
						call envirService.selectQuestionList(objPCInfo , 288)
						%><tr><th colspan="2" class="qu"><span class="mustWrite"><%=form1.getErrorField("q288NetCompany")%></span>2.<%=objPCInfo("qql_question")%></th></tr><%
						i=0
						while not objPCInfo.EOF
						%>
						<%if i mod 2 = 0 then%><tr style="display:none"><%end if%>
						<td ><input type="radio" value="<%=objPCInfo("qqi_number")%>" name="q288NetCompany" <%if objPCInfo("qqi_number")=form1.getField("q288NetCompany") then response.write("checked") end if%> ><%=objPCInfo("qqi_answer")%>
						<% if objPCInfo("qqi_number")=878 then%>
						<input type="text" name="q288NetCompany_other" id="q288NetCompany_other" value="<%=form1.getField("q288NetCompany_other")%>" /><span class="mustWrite"><%=form1.getErrorField("q288NetCompany_other")%></span>
						<% end if %>
						<%if objPCInfo("qqi_description") <> "" then%>&nbsp;<!--<img class="tooltips" src="/images/environment_test/guide_icon_opps.gif" title="<%=objPCInfo("qqi_description")%>" />-->  <%end if%>														
						</td>
						<%if i mod 2 = 1 then%></tr><%end if%>
						<%
							objPCInfo.moveNext()
							i = i+1
						wend%>
						
						<script>
						<%'VIP 選項寫死，再寫到原ims 其他的欄位選項
						%>
						//選其非最後一個選項則消除text內容
						$("input[name='q288NetCompany']").bind('click', function() {
							if  (  $("input[name='q288NetCompany']:checked").val()  !=  $("input[name='q288NetCompany']:last").val() ){
								$("#q288NetCompany_other").val('');
							}
						});
						</script>
						<tr>
							<td><input name="vipq288" id="vipq288_1" onclick="valueToFrom('vipq288_1')" type="radio" value="中国电信"/>中国电信</td>
							<td><input name="vipq288" id="vipq288_2" onclick="valueToFrom('vipq288_2')" type="radio" value="中国移动"/>中国移动</td>
						</tr>
						<tr>
							<td><input name="vipq288" id="vipq288_3" onclick="valueToFrom('vipq288_3')" type="radio" value="中国联通"/>中国联通</td>
							<td><input name="vipq288" id="vipq288_4" onclick="valueToFrom('vipq288_4')" type="radio" value="中国网通"/>中国网通</td>
						</tr>
						<tr>
							<td><input name="vipq288" id="vipq288_5" onclick="valueToFrom('vipq288_5')" type="radio" value="長城寬帶"/>長城寬帶</td>
							<td><input name="vipq288" id="vipq288_6" onclick="valueToFrom('vipq288_6')" type="radio" value="公司网络"/>公司网络</td>
						</tr>
						<tr>
							<td><input name="vipq288" id="vipq288_7" onclick="valueToFrom('vipq288_7')" type="radio" value="宿舍/租屋处/饭店网络 "/>宿舍/租屋处/饭店网络 </td>
							<td><input name="vipq288" id="vipq288_8" onclick="valueToFrom('vipq288_8')" type="radio" value="社区网络"/>社区网络</td>
						</tr>
						<tr>
							<td><input name="vipq288" id="vipq288_9" onclick="valueToFrom('vipq288_9')" type="radio" value="不清楚"/>不清楚</td>
							<td><input name="vipq288" id="vipq288_10" onclick="valueToFrom('vipq288_10')" type="radio" value="其他"/>其他 <input type="text" id="vipq288_10_txt" value=""  maxlength="10"  onclick="valueToFrom('vipq288_10')" onkeyup	="valueToFrom('vipq288_10')"/></td>
						</tr>
						<script>
						function valueToFrom(eleId){
							$("input[name='q288NetCompany']").attr("checked",false);
							$("input[name='q288NetCompany']:last").attr("checked",true);
							$("#q288NetCompany_other").val(  $("#"+eleId ).val() );
							if (  $("#vipq288_10").attr("checked")==false  ){
								$("#vipq288_10_txt").val("");
							}
							if ( ""!=$("#vipq288_10_txt").val() ){
								$("#q288NetCompany_other").val(  $("#vipq288_10_txt").val() );
							}
						}
						if  ( $("#q288NetCompany_other").val()=="中国电信" ){
							$("#vipq288_1").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="中国移动" ){
							$("#vipq288_2").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="中国联通" ){
							$("#vipq288_3").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="中国网通" ){
							$("#vipq288_4").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="長城寬帶" ){
							$("#vipq288_5").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="公司网络" ){
							$("#vipq288_6").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="集体宿舍/出租房/酒店网络" ){
							$("#vipq288_7").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="社区网络" ){
							$("#vipq288_8").attr("checked",true);
						}else if  ( $("#q288NetCompany_other").val()=="不清楚" ){
							$("#vipq288_9").attr("checked",true);
						}else{
							$("#vipq288_10").attr("checked",true);
							 $("#vipq288_10_txt").val(  $("#q288NetCompany_other").val() );
						}
						</script>
					  <tr>
						<td class="ansline"></td>
						<td class="ansline"></td>
					  </tr>
					  <!-- ------->																  
					  <%
						call envirService.selectQuestionList(objPCInfo , 289)
						%><tr><th colspan="2" class="qu"><span class="mustWrite"><%=form1.getErrorField("q289ConnectionMetod")%></span>3.<%=objPCInfo("qql_question")%></th></tr><%
						i=0
						while not objPCInfo.EOF
						%>
						<%if i mod 2 = 0 then%><tr><%end if%>
						<td class="ansline"><input type="radio" value="<%=objPCInfo("qqi_number")%>" name="q289ConnectionMetod" <%if objPCInfo("qqi_number")=form1.getField("q289ConnectionMetod") then response.write("checked") end if%> ><%=objPCInfo("qqi_answer")%>
						<%if objPCInfo("qqi_description") <> "" then%>&nbsp;<!--<img class="tooltips" src="/images/environment_test/guide_icon_opps.gif" title="<%=objPCInfo("qqi_description")%>" />-->  <%end if%>

						</td>
						<%if i mod 2 = 1 then%></tr><%end if%>
						<%
							objPCInfo.moveNext()
							i = i+1
						wend%>																	
					  <!-- ------->																  
					  <%
						call envirService.selectQuestionList(objPCInfo , 287)
						%><tr><th colspan="2" class="qu"><span class="mustWrite"><%=form1.getErrorField("q287NetKind")%></span>4.<%=objPCInfo("qql_question")%></th></tr><%
						i=0
						while not objPCInfo.EOF
						%>
						<%if i mod 2 = 0 then%><tr><%end if%>
						<td ><input type="radio" value="<%=objPCInfo("qqi_number")%>" name="q287NetKind" <%if objPCInfo("qqi_number")=form1.getField("q287NetKind") then response.write("checked") end if%> ><%=objPCInfo("qqi_answer")%>
						<div class="opps-opps"><%=objPCInfo("qqi_description")%></div></td>
						<%if i mod 2 = 1 then%></tr><%end if%><%
							objPCInfo.moveNext()
							i = i+1
						wend%>
					  <tr>
						<td class="ansline"></td>
						<td class="ansline"></td>
					  </tr>
					  <!-- ------->
					  <%
						call envirService.selectQuestionList(objPCInfo , 291)
						%><tr><th colspan="2" class="qu"><span class="mustWrite"><%=form1.getErrorField("q291NetShare")%></span>5.<%=objPCInfo("qql_question")%></th></tr><%
						i=0
						while not objPCInfo.EOF
						%>
						<%if i mod 2 = 0 then%><tr><%end if%>
						<td ><input type="radio" value="<%=objPCInfo("qqi_number")%>" name="q291NetShare" <%if objPCInfo("qqi_number")=form1.getField("q291NetShare") then response.write("checked") end if%> ><%=objPCInfo("qqi_answer")%>
						<div class="opps-opps"><%=objPCInfo("qqi_description")%></div></td>
						<%if i mod 2 = 1 then%></tr><%end if%><%
							objPCInfo.moveNext()
							i = i+1
						wend%>
						<tr>
							<td class="ansline"></td>
							<td class="ansline"></td>
						  </tr>
					</table>
					<div class="opps-box003">
						<dl class="notice clearfix">
							<div class="notice_pic"><img src="/images/environment_test/guide_icon_line_02.gif" /></div>
							<dt>提醒您!<br /> 若您日后上课时有发现网络不稳定的情形，建议您更换网络连接方式：</dt>        
							<dd>
								<ul>
									<li>选择有结网络连接方式</li>
									<li>不与其他电脑共享网络宽带</li>
									<li>上课时保持网络畅通，关闭可能影响上课宽带的软件<br />例如：下载软件、观看网络影片等…</li>
								</ul>
							</dd>
						</dl>
					</div>
						<div class="t-right">
						<% if isPassTest=true then%>
							<input type="button" value="已填写过" onclick="location.href='guide_EnvirDownload.asp?';" class="btn-org QRB" />
						<% end if%>		
						<input type="submit" value="填好送出" class="btn-org QSB"/>
						</div>
					</div>
					</div>
				</form>
				<!-- --------->
                                                                            
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->
<%
objPCInfo.close                               '關閉物件
Set objPCInfo = nothing                       '設定物件為空
%>