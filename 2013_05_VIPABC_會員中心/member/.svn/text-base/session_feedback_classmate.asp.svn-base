<!--#include virtual="/lib/include/html/template/template_change.asp"-->

<%
Dim str_session_sn 	: str_session_sn = request("Session_sn")
if (str_Session_sn="") then 
	response.Redirect "learning_record.asp"
elseif (g_var_client_sn = "") then
	'如果session 掉了 導回首頁
	response.Redirect "http://"&CONST_VIPABC_WEBHOST_NAME&"/index.asp"
end if
Dim mtl 				'小幫手提供的conn開法所需變數
Dim str_sql 			'小幫手提供的conn開法所需變數
Dim var_arr			'小幫手提供的conn開法所需變數
Dim str_tmp_sql		'小幫手提供的conn開法所需變數
Dim arr_result		'小幫手提供的conn開法所需變數

Dim bol_feedback_disable : bol_feedback_disable = false		'預設都可以填

Dim str_session_year 	: str_session_year 	= left(str_session_sn,4)
Dim str_session_month : str_session_month 	= mid(str_session_sn,5,2)
Dim str_session_day 	: str_session_day		= mid(str_session_sn,7,2)
Dim str_session_hour 	: str_session_hour	= mid(str_session_sn,9,2)
Dim str_session_date 	: str_session_date	= DateValue(str_session_year&"/"&str_session_month&"/"&str_session_day)

Dim i
Dim str_rs_client_sn		 
Dim str_rs_cname		 
Dim str_rs_condition	 
Dim str_rs_performance 


%>

<div class="main_con">
       <!--內容start-->
       <form action="session_feedback_exe.asp" name="frm1" method="post">
       <div id="issue_membox">
         <!--問卷資料start-->
         <div class="main_mylist">
           <ul>
             <li><%=getWord("feedback_other")%></li>
             </ul>
         </div>
         <!--問卷資料end-->
         
<%


'判斷可以不是填寫表單的時間
'客戶反應2/15那天無法填寫評鑑，所以特地開放讓客戶填
if (("2011021521404" = str_session_sn) and ("244276" = g_var_client_sn)) then
	bol_feedback_disable = false
elseif ( (datediff("d", str_session_date, now()) > 1 ) or (datediff("d", str_session_date, now())=1 and hour(now())>9) ) then 
	bol_feedback_disable = true
end if


'判斷是否填過其他同學的feedback
	var_arr = Array(str_session_sn,g_var_client_sn)
	str_sql = "select session_sn,client_sn from classmate_session_feedback where session_sn = @session_sn and inputer = @inputer"
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
if (isSelectQuerySuccess(arr_result)) then  '有資料代表有填過
	if (Ubound(arr_result) > 0) then
		if ( not (isnull(arr_result(0,0)) or arr_result(0,0)="")  ) then  
			bol_feedback_disable = true	'填過不能再填 
		end if
	end if
end if

'response.write g_str_sql_statement_for_debug
'撈會員對此堂課其他同學的feedback
	var_arr = Array(g_var_client_sn, str_session_sn, g_var_client_sn)
	str_sql = "SELECT client_attend_list.client_sn, "
	str_sql = str_sql&" ISNULL(client_basic.fname, '') + ' ' + left(ISNULL(client_basic.lname, ''),1) AS cname,  "
	str_sql = str_sql&" classmate_session_feedback.condition, classmate_session_feedback.performance  "
	str_sql = str_sql&" FROM client_attend_list "
	str_sql = str_sql&" LEFT OUTER JOIN classmate_session_feedback ON client_attend_list.client_sn = classmate_session_feedback.client_sn AND client_attend_list.session_sn = classmate_session_feedback.session_sn and classmate_session_feedback.inputer IN (@client_sn1) "
	str_sql = str_sql&" LEFT OUTER JOIN client_basic ON client_attend_list.client_sn = client_basic.sn WHERE (client_attend_list.valid = 1) AND (client_attend_list.session_sn = @session_sn) AND (NOT (client_attend_list.client_sn IN (@client_sn2)))  "
	str_sql = str_sql&" ORDER BY client_attend_list.attend_room " 
	
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
	
'	response.write(g_str_run_time_err_msg) 
'	response.end 
'	response.write g_str_sql_statement_for_debug

'Customer_client_sn_S -> 计算机环境
'Customer_client_sn_SxxF -> 计算机环境細項
'Customer_client_sn_P -> 上课表现



if (isSelectQuerySuccess(arr_result)) then  '有資料
	if (Ubound(arr_result) > 0) then
	
	
		for i = 0 to Ubound(arr_result,2)
		
			str_rs_client_sn		= arr_result(0,i)	 
			str_rs_cname			= arr_result(1,i)		 
			str_rs_condition		= arr_result(2,i)	 
			str_rs_performance 		= arr_result(3,i)

			if (isnull(str_rs_condition) )then 
				str_rs_condition= ""
			end if
			if (isnull(str_rs_performance) )then 
				str_rs_performance= ""
			end if
%>         
         <!--  其他會員名字標頭    -->
         <div class="con_title">
           <div class="con_header_left"><img src="/images/images_cn/arrow_path.gif" alt="" width="4" height="12" hspace="2" /><%=str_rs_cname%></div>
           <div class="con_header_right "></div>
         </div>
         <div class="clear"></div>
         <!--   其他會員名字標頭end    -->
         <!-- 职业别 start-->
        <div class="iss_main_memfiled_all">
        <div class="iss_main_memfiled_left"><%=getWord("computer_environment")%><br />(<%=getWord("net_and_sound_effects")%>)</div>
        <div class="iss_main_memfiled_right">
        <ul>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_S" value="S1" <%if instrrev(str_rs_condition, ", S1,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("regular")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_S" value="S2" <%if instrrev(str_rs_condition, ", S2,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><span class="iss_width_2"><%=getWord("Customer_feedback_bad_1")%></span><input type="radio" name="Customer<%=arr_result(0,i)%>_S2F" value="S2C" <%if instrrev(str_rs_condition, ", S2C,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("continued")%><input type="radio" name="Customer<%=arr_result(0,i)%>_S2F" value="S2S" <%if instrrev(str_rs_condition, ", S2S,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("banished")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_S" value="S3" <%if instrrev(str_rs_condition, ", S3,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><span class="iss_width_1"><%=getWord("Customer_feedback_bad_2")%></span><input type="radio" name="Customer<%=arr_result(0,i)%>_S3F" value="S3C" <%if instrrev(str_rs_condition, ", S3C,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("continued")%><input type="radio" name="Customer<%=arr_result(0,i)%>_S3F" value="S3S" <%if instrrev(str_rs_condition, ", S3S,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("banished")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_S" value="S4" <%if instrrev(str_rs_condition, ", S4,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><span class="iss_width_1"><%=getWord("Customer_feedback_bad_3")%></span><input type="radio" name="Customer<%=arr_result(0,i)%>_S4F" value="S4C" <%if instrrev(str_rs_condition, ", S4C,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("continued")%><input type="radio" name="Customer<%=arr_result(0,i)%>_S4F" value="S4S" <%if instrrev(str_rs_condition, ", S4S,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("banished")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_S" value="S5" <%if instrrev(str_rs_condition, ", S5,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><span class="iss_width_2"><%=getWord("Customer_feedback_bad_4")%></span><input type="radio" name="Customer<%=arr_result(0,i)%>_S5F" value="S5C" <%if instrrev(str_rs_condition, ", S5C,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("continued")%><input type="radio" name="Customer<%=arr_result(0,i)%>_S5F" value="S5S" <%if instrrev(str_rs_condition, ", S5S,", -1 , 1)>0 then response.write " checked " end if %> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("banished")%></li>
        </ul>
        <div class="clear"></div>
        </div>
        <div class="clear"></div>
        </div>
        <!--职业别 end-->
        <div class="iss_main_memfiled_line"></div>
        <!--學生學習狀況-->
        <div class="iss_main_memfiled_all">
        <div class="iss_main_memfiled_left"><%=getWord("class_display")%></div>
        <div class="iss_main_memfiled_right_2">
        <ul>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P1" <%if instrrev(str_rs_performance, ", P1,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("fine")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P4" <%if instrrev(str_rs_performance, ", P4,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_1")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P2" <%if instrrev(str_rs_performance, ", P2,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_2")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P9" <%if instrrev(str_rs_performance, ", P9,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_3")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P7" <%if instrrev(str_rs_performance, ", P7,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_4")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P5" <%if instrrev(str_rs_performance, ", P5,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_5")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P3" <%if instrrev(str_rs_performance, ", P3,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_6")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P10" <%if instrrev(str_rs_performance, ", P10,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_7")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P8" <%if instrrev(str_rs_performance, ", P8,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_8")%></li>
        <li><input type="checkbox" name="Customer<%=arr_result(0,i)%>_P" value="P6" <%if instrrev(str_rs_performance, ", P6,", -1 , 1)>0 then response.write " checked " end if%> <%if bol_feedback_disable then %>disabled="disabled"<%end if%> /><%=getWord("class_display_feedback_9")%></li>
        
        </ul>
        <div class="clear"></div>
        </div>
        <div class="clear"></div>
        </div>
        <!--學生學習 end-->
<%
		next
        
	end if
end if
%>         
        
        
        <div class="t-right m-top25 ">
          <input type="button" name="btn_back" id="btn_back" value="+ <%=getWord("back_to_feedabck")%>" class="btn_1" onClick="javascript:history.back();"/>
          <input type="submit" name="btn_submit" id="btn_submit" value="+ <%=getWord("BUY_1_40")%>" class=" m-left5 btn_1"  <%if bol_feedback_disable or request("session_feedback_disable") then %>disabled="disabled"<%end if%>  />
        </div>
         
       </div>
        <input name="session_sn" id="session_sn" type="hidden" value="<%=str_Session_sn%>" />
        <!-- 記錄session_feedback.asp post過來的參數 -->
        <input name="consultant" type="hidden" value="<%=request("consultant")%>">
        <input name="speed" type="hidden" value="<%=request("speed")%>">
        <input name="distribution" type="hidden" value="<%=request("distribution")%>">
        <input name="tskill" type="hidden" value="<%=request("tskill")%>">
        <input name="behavior" type="hidden" value="<%=request("behavior")%>">
        <input name="materials" type="hidden" value="<%=request("materials")%>">
        <input name="difficulty" type="hidden" value="<%=request("difficulty")%>">
        <input name="requirement" type="hidden" value="<%=request("requirement")%>">
        <input name="overall" type="hidden" value="<%=request("overall")%>">
        <input name="connection" type="hidden" value="<%=request("connection")%>">
        <input name="suggestions" type="hidden" value="<%=request("suggestions")%>">
        <input name="compliment" type="hidden" value="<%=request("compliment")%>">
        <input name="help_note" type="hidden" value="<%=request("help_note")%>">
		<input name="C1" type="hidden" value="<%=request("C1")%>">
		<input name="ctime" type="hidden" value="<%=request("ctime")%>">
       </form>
       <!-- 內容end -->
                <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
       
    </div>
