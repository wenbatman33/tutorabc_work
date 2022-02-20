<!--#include virtual="/lib/include/html/template/template_change.asp"-->

<%
Dim str_Session_sn : str_Session_sn = request("Session_sn")
if (str_Session_sn="") then 
	response.Redirect "learning_record.asp"
end if

Dim str_sql
Dim var_arr
Dim arr_result

Dim rs_consultant
Dim rs_attend_date
Dim rs_attend_sestime
Dim rs_material_title
Dim rs_material_descript
Dim rs_attend_level
Dim rs_course
Dim rs_con_sn
Dim rs_nlevel
Dim material_level_class

Dim rs_participation	
Dim rs_pronunciation	
Dim rs_comprehension	
Dim rs_creativity	
Dim rs_facility	
Dim rs_con_comments	
Dim rs_sel		

'Dim arr_progress : arr_progress = array("", "Outstanding", "Excellent", "Great", "Good", "Needs to participate more")
Dim arr_progress : arr_progress = array("", "Needs to participate more", "Good", "Great", "Excellent", "Outstanding")

Dim i

'20100125 從學習狀態查詢過來的不能填寫課後評語
Dim str_read_only : str_read_only = Trim(Request("read_only"))
%>


<div id="main_con">


       <!--內容start-->
        <div class="main_membox">
        <!--上版位課程start-->
        <!--#include virtual="/lib/include/html/include_learning.asp"--> 
        <!--上版位課程end-->
        
<%
'此節課的課程資訊(教材名稱、上課時間、顧問名稱、課程等級、教材簡述)
	var_arr = Array(str_Session_sn, g_var_client_sn)

	str_sql="SELECT con_basic.basic_fname + ' ' + con_basic.basic_lname AS consultant, client_attend_list.Attend_Date,  "
	str_sql=str_sql&"client_attend_list.Attend_Sestime, material.LTitle, material.LD, client_attend_list.Attend_Level, "
	str_sql=str_sql&" material.Course, con_basic.Con_Sn,client_basic.nlevel "
	str_sql=str_sql&" FROM client_attend_list(nolock) "
	str_sql=str_sql&" LEFT OUTER JOIN material(nolock) ON client_attend_list.Attend_mtl_1 = material.Course "
	str_sql=str_sql&" LEFT OUTER JOIN con_basic(nolock) ON client_attend_list.Attend_consultant = con_basic.Con_Sn "
	str_sql=str_sql&" LEFT JOIN client_basic on client_basic.sn = client_attend_list.client_sn "
	str_sql=str_sql&" WHERE (client_attend_list.Session_sn = @Session_sn) and client_attend_list.valid = 1 "
	str_sql=str_sql&" and (client_sn = @client_sn)"

	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)

if (isSelectQuerySuccess(arr_result)) then
    '有資料
	 if (Ubound(arr_result) > 0) then

		 rs_consultant		= arr_result(0,0)
		 rs_attend_date		= arr_result(1,0)
		 rs_attend_sestime	= arr_result(2,0)
		 rs_material_title	= arr_result(3,0)
		 rs_material_descript	= arr_result(4,0)
		 rs_attend_level		= arr_result(5,0)
		 rs_course			= arr_result(6,0)
		 rs_con_sn			= arr_result(7,0)
		 rs_nlevel			= arr_result(8,0)

%>        
        
        
        <div class="red_title"><%=getWord("CLASS_15")%></div>
        <div>
            <br />
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="score_bd">
          <tr>
            <th>
            <div class="Div_relative">
            </div>
            <%=getWord("BOOKING_ML")%></th>
            <td ><span class="color_1"><%=rs_material_title%></span></td>
            </tr>
          <tr>
            <th><%=getWord("attend_time")%></th>
            <td><%=rs_attend_date%> - <%=rs_attend_sestime%>:30</td>
            </tr>
          <tr>
            <th><%=getWord("consultant_name")%></th>
            <td><%=rs_consultant%></td>
          </tr>
          <tr>
            <th><%=getWord("class_level")%></th>
            <td>
            
            <%
				for i = 1 to 12
					material_level_class = ""
					if (i = rs_attend_level) then 
						material_level_class = "green_ballNum"
					else
						material_level_class = "gray_ballNum"
					end if
					Response.Write "<p class="&material_level_class&">"&i&"</p>"
				next
			%>
            
            
            </td>
          </tr>
          <tr>
            <th><%=getWord("class_mtl_excerpt")%></th>
            <td><%=rs_material_descript%></td>
          </tr>
        </table>
        <div class="clear"></div>
        </div>
<%
		'顾问课后评价
		var_arr = Array(str_Session_sn, g_var_client_sn, str_Session_sn, g_var_client_sn, str_Session_sn, g_var_client_sn, g_var_client_sn, str_Session_sn)
		
		str_sql = " select "
		str_sql = str_sql & " progress.session_sn,progress.client_sn, participation, pronunciation, "
		str_sql = str_sql & " comprehension, creativity, facility, c_comments,sel "
		str_sql = str_sql & " from ( "
		str_sql = str_sql & "	select session_sn,client_sn, participation, pronunciation, "
		str_sql = str_sql & "	comprehension, creativity, facility, c_comments "
		str_sql = str_sql & "	from Client_Progress_Report "
		str_sql = str_sql & "	where session_sn = @session_sn_1 and client_sn = @client_sn_1 "
		str_sql = str_sql & "	and not exists( "
		str_sql = str_sql & "		select * from ( "
		str_sql = str_sql & "		select client_sn from Client_Progress_Report_history "
		str_sql = str_sql & " 		where session_sn = @session_sn_2 and client_sn = @client_sn_2 "
		str_sql = str_sql & "	)Report_history "
		str_sql = str_sql & "	where Client_Progress_Report.client_sn=Report_history.client_sn "
		str_sql = str_sql & " ) "
		str_sql = str_sql & " union all "
		str_sql = str_sql & " select session_sn, client_sn, participation, pronunciation, "
		str_sql = str_sql & " comprehension, creativity, facility, c_comments "
		str_sql = str_sql & " from Client_Progress_Report_history "	
		str_sql = str_sql & "		where sn = ( "
		str_sql = str_sql & "			select max(sn) as sn "
		str_sql = str_sql & "			from Client_Progress_Report_history "
		str_sql = str_sql & "			where session_sn = @session_sn_3 and client_sn = @client_sn_3 "
		str_sql = str_sql & "		) "
		
		str_sql = str_sql & " ) as progress "
		str_sql = str_sql & " LEFT OUTER JOIN ( "
		str_sql = str_sql & "	SELECT client_trial_process.session_sn, "
		str_sql = str_sql & "	CASE isnull(client_trial_process.agree, 3) WHEN 1 THEN '同意' WHEN 0 THEN '不同意' ELSE '' END + '' + CASE client_trial_list.selection WHEN 'Upgrade' THEN '升級' WHEN 'Downgrade' THEN '降級' ELSE '' END AS sel, "
		str_sql = str_sql & "	client_trial_list.client_sn FROM client_trial_process "
		str_sql = str_sql & "	LEFT OUTER JOIN client_trial_list  "
		str_sql = str_sql & "	ON client_trial_process.tsn = client_trial_list.sn "
		str_sql = str_sql & "	WHERE (client_trial_list.client_sn = @client_sn_4) AND "
		str_sql = str_sql & "	(client_trial_process.session_sn = @session_sn_4) "
		str_sql = str_sql & " ) AS derivedtbl_1 "
		str_sql = str_sql & " ON progress.client_sn = derivedtbl_1.client_sn "
		str_sql = str_sql & " AND progress.session_sn = derivedtbl_1.session_sn  "
	
		arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
	'response.write(g_str_run_time_err_msg) 
	'response.end 
	'response.write g_str_sql_statement_for_debug &"<br>"
	'response.write isSelectQuerySuccess(arr_result) &"<br>"
		if (isSelectQuerySuccess(arr_result)) then
			'有資料
			 if (Ubound(arr_result) > 0) then
		
			 rs_participation = arr_result(2,0)
             '阿捨修正空值別帶入陣列
             if ( isEmptyOrNull(rs_participation) ) then
                rs_participation = 0
             end if
			 
             rs_pronunciation = arr_result(3,0)

             if ( isEmptyOrNull(rs_pronunciation) ) then
                rs_pronunciation = 0
             end if

			 rs_comprehension = arr_result(4,0)

             if ( isEmptyOrNull(rs_comprehension) ) then
                rs_comprehension = 0
             end if

			 rs_creativity = arr_result(5,0)
             
             if ( isEmptyOrNull(rs_creativity) ) then
                rs_creativity = 0
             end if

			 rs_facility = arr_result(6,0)

             if ( isEmptyOrNull(rs_facility) ) then
                rs_facility = 0
             end if

			 rs_con_comments	= arr_result(7,0)
			 rs_sel			= arr_result(8,0)
			 
%>        
        
        <div class="red_title"><%=getWord("con_feedback")%></div>
        <div>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="bought_bd">
          <tr>
            <th><%=getWord("con_feedback_participate")%></th>
            <th><%=getWord("CLASS_VOCABULARY_8")%></th>
            <th><%=getWord("con_feedback_intellect")%></th>
            <th><%=getWord("con_feedback_create")%></th>
            <th class="rightLine"><%=getWord("con_feedback_flowing")%></th>
          </tr>
          
          <tr>
            <td><%=arr_progress(rs_participation)%><%'=rs_participation%></td>
            <td><%=arr_progress(rs_pronunciation)%><%'=rs_pronunciation%></td>
            <td><%=arr_progress(rs_comprehension)%><%'=rs_comprehension%></td>
            <td><%=arr_progress(rs_creativity)%><%'=rs_creativity%></td>
            <td class="rightLine"><%=arr_progress(rs_facility)%><%'=rs_facility%></td>
          </tr>
          <tr>
            <th colspan="5" class="rightLine"><%=getWord("con_feedback_word")%></th>
          </tr>
          <tr>  
            <td colspan="5" class="rightLine"><%=rs_con_comments%></td>
          </tr>
        </table>
        
          

        <div class="clear"></div>
        </div>
<%
			end if
		end if
		
		'20100125 從學習狀態查詢過來的不能填寫課後評語
		If (str_read_only="y") Then
			str_read_only = "style=""display:none;"""
		Else
			str_read_only = ""
		End If
%>        
        <div class="page">
        <input type="button" name="btn_score" id="btn_score" value="+ <%=getWord("go_feedback_and_something_another")%>" class="btn_red3" onclick="javascript:location.href='session_feedback.asp?session_sn=<%=str_Session_sn%>'" <%=str_read_only%> />
        </div>
        
<%
	end if
end if
%>        
        
        </div>
        <!--內容end-->
                <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>

        </div>
