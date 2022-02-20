<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/backoffice/inc/conn.asp"-->
<!--#include virtual="/lib/functions/functions_free_lobby_session.asp"-->
<div class="okbox">
    <div class="ok">
	<input type="hidden" name="expoURL" id="expoURL" value="<%=Trim(Getrequest("expoURL",CONST_DECODE_NO))%>" />
        <%
        if (isEmptyOrNull(g_var_client_sn)) then
            Call alertGo("您尚未登入会员,请前往登入或立即加入免费会员", "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/member/member_login.asp", CONST_ALERT_THEN_REDIRECT)
            Response.End
        end if
		'接續訂課頁面的session值
		dim str_org_session_sn: str_org_session_sn = Trim(getRequest("booking",CONST_DECODE_NO))


		'利用session值取得訂課時間
		dim attend_date: attend_date = left(str_org_session_sn,4)&"/"&mid(str_org_session_sn,5,2)&"/"&mid(str_org_session_sn,7,2)
		dim attend_time: attend_time = mid(str_org_session_sn,9,2)
		
		dim bol_is_double_booking ,bol_is_booking_full, p_int_connect_type
		dim p_arr_data, special_sn, mailbody, class_yesno, csh_source_website, csh_source_code
        dim valid : valid = "1"
		dim int_client_sn : int_client_sn = g_var_client_sn
        dim arrFreeLobbySessionData : arrFreeLobbySessionData = null
        '抓出對應的相關資訊
        arrFreeLobbySessionData = getFreeLobbySessionData1(attend_date)
        strValidTime = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_TIME)
        strValidClassName = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_NAME)
        attend_room = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_ROOM_NUMBER)
        strValidDescription1 = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION1)
        'strValidDescription2 = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION2)
        attend_material = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_MATERIAL)
        attend_consultant = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CONSULTANT)
		'檢查是否有訂過課，有則去將其lead_sn找出來(解決寄信的lead_sn不見問題)
		dim p_str_lead_sn, sql, arr_result,arr_param_data
		sql = " SELECT sn FROM lead_basic WHERE client_sn = @client_sn"
        arr_param_data = Array(int_client_sn)
		arr_result = excuteSqlStatementRead(sql, arr_param_data, CONST_VIPABC_R_CONN)
		'response.write sql
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then
				p_str_lead_sn = arr_result(0,0)
			end if 
		end if

		dim Client_name: Client_name = Session("client_cname")&" ( "&Session("client_fname")&" ) "
		
		'檢查客戶是否有登入
		if (isEmptyOrNull(int_client_sn)) then
			Call alertGo(getWord("PLEASE_LOGIN"), "/program/member/member_login.asp", CONST_ALERT_THEN_REDIRECT)
			Response.End
		end if
		
		'檢查是否有重覆訂課
		bol_is_double_booking  =  checkClientDoubleBooking(int_client_sn, CONST_VIPABC_R_CONN)
		
		'檢查課堂人數是否已滿
		'Eddie 20100513 修訂5/14課堂人數上限
		bol_is_booking_full  =  checkFullClientBooking(attend_date, attend_time, CONST_VIPABC_R_CONN, CONST_LIMIT_STUDENTS_OF_FREE_EXPO_RESERVATION)
		
		'判斷是否有選取訂課時間，沒有，則回到訂課頁面
		'20100810 阿捨暫修訂強制將課程判斷已滿額 先註解,之後會改回應有的判斷
        if (isEmptyOrNull(str_org_session_sn) and isEmptyOrNull(attend_consultant)) then
				response.write "您尚未选取课程，请回到订课页面选取您喜爱的上课时间。"
		else
			if bol_is_double_booking = true then
				response.write "您已经预约过生活英语大会堂课程.如需要免费在线真人体验请洽 "& CONST_SERVICE_PHONE &" ,我们将有语言顾问为您服务。"
					
            elseif (true = bol_is_booking_full) then
			    response.write "此堂课已满额,建议您选择其他时段!"
		    else	
		
				p_arr_data = Array(int_client_sn, attend_date, attend_time, attend_consultant, attend_room, attend_material, str_org_session_sn, valid, csh_source_website, csh_source_code)
				insertFreeMemberAttendList p_arr_data, CONST_VIPABC_RW_CONN			
                Response.Write "您已成功预订" & "<br/>" '您已成功预订
                Response.Write "<font color = red>"&attend_date &"  "& attend_time &":30-21:30 " & " 大会堂</font>，" & "<br/>"
				Response.write "课程共1课时共计扣除时数<font color = red> 0 </font> "
				
				'訂課成功則呼叫summer寫好的預存程序將資料寫入tutormeet資料庫
					dim insert_sql, result					
					insert_sql = " exec Sch_Meet_up_free_member_attend_list '"& attend_date &"', '"& attend_time &"', '"& int_client_sn &"', 'i' "
					'if attend_time <> 0 and int_client_sn <> 0 and not isEmptyOrNull(attend_date) then
						result = excuteProcAndFuncQuick(insert_sql, CONST_VIPABC_R_CONN)
					'end if

				'訂課成功則寄信給kevin及eric
					mailbody = "Dear Kevin :<br>"
					mailbody = mailbody & "您的客戶 " & Client_name
					mailbody = mailbody & "(<a href=""http://" & CONST_VIPABC_IMSHOST_NAME & "/asp/bd/client_2.asp?client="&int_client_sn&"&lead="&p_str_lead_sn&"&search=yes"">客戶業務追蹤明細</a>)"
					mailbody = mailbody & " 在" & date & "於官網預約生活英语大會堂,預約時段為"
					mailbody = mailbody & attend_date & " "&strValidTime&"，請主動去電客戶安排生活英语大會堂，謝謝!<br>"
					
					sendEmailQuick "kevinlai@vipabc.com;" & CONST_MAIL_TO_RD_PG_VIPABC_ACCOUNT & ";"  & "","","VIPABC客戶"&Client_name&"生活英语大會堂訂課通知",mailbody, CONST_VIPABC_WEBSITE
				    '清空
				dim strReadClientSql ,arrPara, arrResult, strClientMail,P_strClientMail
				arrPara = array(int_client_sn)
				strReadClientSql = "select email from client_basic where sn = @client_sn"
				arrResult = excuteSqlStatementRead(strReadClientSql, arrPara, CONST_VIPABC_RW_CONN)
				if (isSelectQuerySuccess(arrResult)) then
					if (Ubound(arrResult) >= 0) then
						P_strClientMail = arrResult(0,0)
					end if
				end if
				'訂課成功則通知客戶
					mailbody= " <font size=""2"" face=""Arial, Helvetica, sans-serif"" color=""#3e3e3e"">尊敬的 "&Client_name&" 您好<br /><br />"
					mailbody = mailbody & " 您刚刚已经成功预约免费大会堂上课时间为：<font color=""#df0017""><b>"&attend_date&" "&strValidTime&"</b></font><br /><br />"
					mailbody = mailbody & " 请您于当天 <b>20:25分</b> 登入会员并前往会员专区点选进入免费大会堂即可体验<br /><br />"
					mailbody = mailbody & " 请您提前点选下方网址，来确认您可以顺利地进入教室<br>"
					mailbody = mailbody & " <a href=""http://www.tutormeet.com/mictest/mictest.html?lang=3&compstatus=vipabc"" target=""_blank""><font color=""#3e3e3e"">http://www.tutormeet.com/mictest/mictest.html?lang=3&compstatus=vipabc</font></a><br><br>"
					mailbody = mailbody & " 感谢您对VIPABC的支持与爱护，祝您学习愉快！<br><br>"
					mailbody = mailbody & " <a href=""http://www.vipabc.com"" target=""_blank""><font color=""#3e3e3e"">VIPABC客户服务团队敬上</font></a><br />"
					mailbody = mailbody & " " & CONST_SERVICE_PHONE & " </font>"
					if (not isEmptyOrNull(P_strClientMail)) then
                        sendEmailQuick P_strClientMail,"","VIPABC大会堂上课通知,请准时出席",mailbody, CONST_VIPABC_WEBSITE
                    end if
			end if
		end if	
        %>
    </div>
		<div class="btn">
			<input type="submit" name="button" id="button" value="+ 回首页" class="btn_1" onClick="location.href='http://<%=CONST_VIPABC_WEBHOST_NAME%>'"/>
		</div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>