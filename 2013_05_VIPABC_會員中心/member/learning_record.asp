<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'選擇喜愛顧問%>
<script language='javascript1.2' type="text/javascript">
function set_fav_con(con_name,con_sn,fov){

	var show_msg ="";
	if(fov ==0){
		show_msg ="<%=getWord("set_con_confirm")%>"+unescape(con_name.replace("_"," "))+"<%=getWord("set_like_con")%>?";
	}else{
		show_msg ="<%=getWord("cancel_like_con")%>"+unescape(con_name.replace("_"," "))+"<%=getWord("ma")%>?";
	}
	
	if (confirm(show_msg)){
		location.href='set_fav_con.asp?language=<%=request("language")%>&con_sn='+con_sn+'&fov='+fov+'';
	}else{
		return false;
	}
}

</script>


<%
Dim str_hour  : str_hour = hour(now)

Dim str_sql
Dim var_arr
Dim arr_result

Dim str_rs_attend_list_sn
Dim str_rs_attend_consultant
Dim str_rs_session_sn
Dim str_rs_attend_date
Dim str_rs_attend_sestime
Dim str_class_end_hour
Dim str_rs_ltitle
Dim str_rs_con_name
Dim str_rs_fav_con_sn

Dim i
Dim str_show_li_class	: str_show_li_class = "class=""C4_BgGreen"""
Dim bol_show_li_class	: bol_show_li_class = false
Dim str_show_title


Dim str_session_year 	
Dim str_session_month 
Dim str_session_day 	
Dim str_session_hour 	
Dim str_session_date 	

Dim str_session_feedback_sn 
'20101222 tree RD10121405 移除未進入教室客戶的課後評薦功能
Dim strClassYesno	

Dim str_btn_collect_value
Dim int_ctrl_collect 
%>

<div id="main_con">


   <!--內容start-->
    <div class="main_membox">
    	<div class='page_title_2'><h2 class='page_title_h2'>上课记录&评分</h2></div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 
    <!--上版位課程start-->
    <!--#include virtual="/lib/include/html/include_learning.asp"--> 
    <!--上版位課程end-->
    <!--上課紀錄評分-->
    <ul class="class_list4">
        
<%
if isEmptyOrNull(g_var_client_sn) then
	g_var_client_sn=session("client_sn")
end if

	'20101222 tree RD10121405 移除未進入教室客戶的課後評薦功能 (補上client_attend_list.class_yesno)
	var_arr = Array(g_var_client_sn, g_var_client_sn, g_var_client_sn, str_hour)

	str_sql = " SELECT client_attend_list.sn, client_attend_list.client_sn, client_attend_list.Attend_consultant, client_attend_list.Session_sn, "
	str_sql = str_sql & " CONVERT(varchar,client_attend_list.Attend_Date, 111) as Attend_Date,client_attend_list.Attend_Sestime,material.LTitle, "
	str_sql = str_sql & " con_basic.basic_fname + ' ' + con_basic.basic_lname AS con_name , isnull(fav_con.ccr_object_sn,0) as fav_con_sn ,client_session_evaluation.sn as session_feedback_sn, client_attend_list.class_yesno, client_attend_list.session_state"
	str_sql = str_sql & ", CASE "
	str_sql = str_sql & "when client_attend_list.attend_livesession_types = 1 OR client_attend_list.attend_livesession_types = 7 THEN '1to1' "
	str_sql = str_sql & "when client_attend_list.special_sn IS NULL THEN '1to6' "
	str_sql = str_sql & "WHEN client_attend_list.special_sn IS NOT NULL THEN 'ss' "
	str_sql = str_sql & "END classType,client_attend_list.Attend_mtl_1 "
	str_sql = str_sql & " FROM client_attend_list "
	str_sql = str_sql & " LEFT  JOIN con_basic ON client_attend_list.Attend_consultant = con_basic.Con_Sn "
	str_sql = str_sql & " LEFT  JOIN material ON client_attend_list.Attend_mtl_1 = material.Course "
'----客戶喜愛的顧問
	str_sql = str_sql & " left join ( "
	str_sql = str_sql & " select ccr_object_sn from dcgs_client_class_member_require_bas where client_sn = @g_var_client_sn_1 "
	str_sql = str_sql & " and ccr_require_object = 'T' and ccr_require_code = 'F' and ccr_valid = 1 "
	str_sql = str_sql & " and (DATEDIFF(d,getdate(),ccr_sdate) <=0 and DATEDIFF(d,getdate(),ccr_edate) >=0) "
	str_sql = str_sql & " )  fav_con "
	str_sql = str_sql & " on fav_con.ccr_object_sn = client_attend_list.Attend_consultant "
	str_sql = str_sql & " left join client_session_evaluation(nolock) on client_session_evaluation.session_sn = client_attend_list.session_sn and client_session_evaluation.client_sn = client_attend_list.client_sn "
	str_sql = str_sql & " WHERE (client_attend_list.client_sn = @g_var_client_sn_2 ) AND (client_attend_list.Attend_consultant IS NOT NULL) "
	str_sql = str_sql & " AND (client_attend_list.Attend_mtl_1 IS NOT NULL) AND (client_attend_list.Session_sn IS NOT NULL)  "
	str_sql = str_sql & " AND (client_attend_list.Attend_Date < getdate()) AND (client_attend_list.Valid = 1)  "
	str_sql = str_sql & " OR (client_attend_list.client_sn = @g_var_client_sn_3 ) "
	str_sql = str_sql & " AND (client_attend_list.Attend_consultant IS NOT NULL) AND (client_attend_list.Attend_mtl_1 IS NOT NULL)  "
	str_sql = str_sql & " AND (client_attend_list.Session_sn IS NOT NULL) AND (client_attend_list.Attend_Date = getdate()) "
	str_sql = str_sql & " AND (client_attend_list.Attend_Sestime < @g_str_hour) "
	str_sql = str_sql & " ORDER BY dbo.client_attend_list.Session_sn DESC "
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
'response.write str_sql
'response.write(g_str_run_time_err_msg) 
'response.end 
'response.write g_str_sql_statement_for_debug

if (isSelectQuerySuccess(arr_result)) then
    '有資料
	 if (Ubound(arr_result) > 0) then
%>
        <li>
        <h1><%=getWord("booking_date")%></h1>
        <h1><%=getWord("booking_time")%></h1>
        <h2 style="width:220px"><%=getWord("booking_ml")%></h2>
		<h3 style="width:40px">类型</h3>
        <h3><%=getWord("CLASS_22")%></h3>
        <h4></h4>
        </li>
<%
		for i = 0 to Ubound(arr_result,2)

			str_rs_attend_list_sn 			= arr_result(0,i)
			str_rs_attend_consultant 		= arr_result(2,i)
			str_rs_session_sn 				= arr_result(3,i)
			str_rs_attend_date 				= arr_result(4,i)
			str_rs_attend_sestime 			= arr_result(5,i)
			if (len(str_rs_attend_sestime) < 2 ) then 
				str_rs_attend_sestime = "0" & str_rs_attend_sestime
			end if
			if left(arr_result(13,i),1) = "5" then
				str_rs_ltitle				= "Special 1 on 1 Session 自备教材"
			else
				str_rs_ltitle				= arr_result(6,i)
			end if
			str_rs_con_name					= arr_result(7,i)
			str_rs_fav_con_sn				= arr_result(8,i)
			str_session_feedback_sn		= arr_result(9,i)
			strClassYesno	 			= arr_result(10,i)
	        strSessionStat	 			= arr_result(11,i)
			if (isnull(str_rs_con_name)) then 
				str_rs_con_name = ""
			end if
		
			
			if (bol_show_li_class) then 
				str_show_li_class = "class=""C4_BgGreen"""
				bol_show_li_class = false
			else
				str_show_li_class = ""
				bol_show_li_class = true
			end if

			sessionDateTime = str_rs_attend_date &" "& str_rs_attend_sestime&":30" '課程日期時間
			newUIOnlineDateTime = "2011/12/28 23:20" '新版面開始使用日期時間
			if sessionDateTime > newUIOnlineDateTime then
				redirectFileName = "session_feedback_new.asp"
				redirectFileName1 = "session_feedback_new.asp"
				status = "&status=save"
			else
				redirectFileName = "learning_record_detail.asp"
				redirectFileName1 = "session_feedback.asp"
				status = ""
			end if

		
%>        
        <li <%=str_show_li_class%> >
          <h1>
			<!--<%=str_rs_Attend_Date%>   jim 2013/4/22 註解-->
			<%'jim 2013/4/22 因short session 改時間顯示方式 start			
			sessiontime = getSessionWholeTime(str_rs_session_sn,1,0,CONST_TUTORABC_RW_CONN)
			Arrsessiontime = Split(sessiontime, " ")
			sessiondate = Arrsessiontime(0)
			sessionstarttoendHour = Arrsessiontime(1) & Arrsessiontime(2) & Arrsessiontime(3)
			'jim 2013/4/22 因short session 改時間顯示方式 end	
			%>
			<%=sessiondate%><!-- jim 2013/4/22 add -->

		</h1>
        <h1>
				<!--@<%=str_rs_attend_sestime%>:30   jim 2013/4/22 註解
				~ 
				<%
					if (str_rs_attend_sestime = "23") then  '23(小時)+1 為00
						response.Write "00"
					else
						str_class_end_hour = cint(str_rs_attend_sestime)+1
						if (len(str_class_end_hour) < 2) then 
							response.Write "0" & str_class_end_hour 
						else
							response.Write str_class_end_hour
						end if
					end if
				%>
				:15  <br>  -->

			<%=sessionstarttoendHour%> <!-- jim 2013/4/22 add -->
        </h1>
        <h2 style="width:220px">
			<%
            '判斷可以不是填寫表單的時間
            
            str_session_year 	= left(str_rs_session_sn,4)
            str_session_month 	= mid(str_rs_session_sn,5,2)
            str_session_day 	= mid(str_rs_session_sn,7,2)
            str_session_hour 	= mid(str_rs_session_sn,9,2)
            str_session_date 	= DateValue(str_session_year&"/"&str_session_month&"/"&str_session_day)
    
            if ( (datediff("d", str_session_date, now()) > 1 ) or (datediff("d", str_session_date, now())=1 and hour(now())>9) or str_session_feedback_sn<>"" ) then 
                
            %>	
            <a href="<%=redirectFileName%>?session_sn=<%=str_rs_session_sn&status%>"><%=str_rs_ltitle%></a>
            <%	
            else
            %>
			<%=str_rs_ltitle%>
				<%'20101222 tree RD10121405 移除未進入教室客戶的課後評薦功能 --開始--%>
				<%if (((not isEmptyOrnull(strClassYesno)) and (strClassYesno <> 0)) or ((not isEmptyOrnull(strSessionStat)) and (strSessionStat = 1 or strSessionStat = 2))) then %>
					<input type="button" name="btn_score" id="btn_score" value="+ <%=getWord("grade")%>" class="btn_cancel" style="width:50px;" onclick="javascript:location.href='<%=redirectFileName1%>?session_sn=<%=str_rs_session_sn%>'"/>
				<%else%>
					<input type="button" value="+ <%=getWord("grade")%>" class="btn_cancel" style="width:50px;" onclick='alert("您未出席本课程，无须填写课后评语")'/>
				<%end if%>
				<%'20101222 tree RD10121405 移除未進入教室客戶的課後評薦功能 --結束--%>
			
			
            
            <%
            end if
            %>
        </h2>
		<h3 style="width:40px"><img src="/images/sessionFeedback/feedback_class_<%=arr_result(12,i)%>_zh-cn.gif"></h3>
        <h3><%=str_rs_con_name%></h3>
        <h4>
		<%
	
        if str_rs_fav_con_sn = str_rs_attend_consultant then 
			str_show_title = getWord("cancel_con")
			str_btn_collect_value =	getWord("cancel_con")
			int_ctrl_collect = 1
		else
			str_show_title = getWord("collect_con")&chr(13)&getWord("collect_con_1")
			str_btn_collect_value =	getWord("collect_consultant")
			int_ctrl_collect = 0
		end if
			
        %>
        <input type="button" name="btn_collect" id="btn_collect" value="<%=str_btn_collect_value%>" class="btn_cancel" title="<%=str_show_title%>" onclick="javascript: set_fav_con('<%=escape(replace(str_rs_con_name," ","_"))%>',<%=str_rs_attend_consultant%>,<%=int_ctrl_collect%>);"/>

        </h4>
        </li>

<%
		next
	end if
else


end if
	
%>        
    </ul>
<!--        <div class="page"><a href="#">上一頁</a><a href="#">1</a><a href="#">2</a><a href="#">3</a><a href="#">4</a><a href="#">5</a><a href="#">6</a><a href="#">7</a><a href="#">8</a><a href="#">下一頁</a></div>
-->        <!--上課紀錄評分end-->
    </div>
    <!--內容end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>

</div>
