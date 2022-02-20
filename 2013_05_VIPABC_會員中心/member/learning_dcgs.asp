<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type='text/javascript' src="/lib/javascript/framepage/js/learning_dcgs.js"></script>

 <!--內容start-->

        <div class="main_membox">

        <div class='page_title_3'><h2 class='page_title_h2'>DCGS学习偏好设置</h2></div>
		<div class="box_shadow" style="background-position:-80px 0px;"></div> 
        <!--上版位課程start-->
		<% 
		'如果是從新手上路過來的 就不秀tab的內容 Johnny 2012-08-22
		if not "y" = getRequest("noTab", CONST_DECODE_NO) then %>
        <!--#include virtual="/lib/include/html/include_learning.asp"--> 
		<% end if %>
        <!--上版位課程end-->
        <!--勾選start-->
        <form name="frm_learning_dcgs" action="learning_dcgs_exe.asp" method="post" onsubmit="return checkForm();" >
        <%
		
		
		Dim str_industry : str_industry = ""
		Dim str_role : str_role = ""
		Dim str_interest : str_interest = ""
		Dim str_focus : str_focus = ""
		
		'塞資料的相關設定值
		Dim str_sql
		Dim var_arr 			'傳excuteSqlStatementRead的陣列
		Dim arr_result		'接回來的陣列值
		
		Dim i
		Dim str_check				'checked
		
		Dim int_industry_sn		'職業別 lesson_industry.sn
		Dim str_industry_name		'職業別名稱 lesson_industry.cname
		Dim str_role_sn			'職位別 lesson_role.sn
		Dim str_role_name			'職位別名稱 lesson_role.cname
		Dim str_interest_sn		'興興 cfg_interest.sn
		Dim str_interest_name		'興趣名稱 cfg_interest.cname
		
		str_sql = " select sn,email,focus,industry,role,interest "		'DCGS偏好設定
		str_sql = str_sql & "  from client_basic where sn = @sn"

		var_arr = Array(g_var_client_sn)
		arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
		'response.write str_sql
		'Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
		'Response.Write("錯誤訊息" & g_int_run_time_err_code & "<br>")
		'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
		if (isSelectQuerySuccess(arr_result)) then
		   
    		if (Ubound(arr_result) > 0) then
				'有資料
				str_focus = Trim(arr_result(2,0)) 
				str_industry = arr_result(3,0)
				str_role = arr_result(4,0)
				str_interest = arr_result(5,0)	
			else
				'無資料
			end if 
		end if  
		
		
		%>
        <div>
        <div class="line01"></div>
        <!--最想加强的是 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("dcgs_5")%></div>
        <div class="main_memfiled_right">
        <input type="radio" id="radio" name="focus" value="4" <%IF (str_focus = "4") Then response.write "checked" end if %>/><%=getWord("dcgs_6")%>
        <input type="radio" id="radio" name="focus" value="5" <%IF (str_focus = "5") Then response.write "checked" end if %>/><%=getWord("dcgs_7")%>
        <input type="radio" id="radio" name="focus" value="2" <%IF (str_focus = "2") Then response.write "checked" end if %>/><%=getWord("CLASS_VOCABULARY_7")%>
        <input type="radio" id="radio" name="focus" value="1" <%IF (str_focus = "1") Then response.write "checked" end if %>/><%=getWord("dcgs_8")%>
        <input type="radio" id="radio" name="focus" value="6" <%IF (str_focus = "6") Then response.write "checked" end if %>/><%=getWord("dcgs_9")%>
        <input type="radio" id="radio" name="focus" value="3" <%IF (str_focus = "3") Then response.write "checked" end if %> /><%=getWord("CLASS_VOCABULARY_8")%></div>
        <div class="clear"></div>
        </div>
        <!--最想加强的是 end-->
        <div class="main_memfiled_line2"></div>
        <!--职业别 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("ROLE")%><br />(<%=getWord("dcgs_10")%>1 ~ 2 项)</div>
        <div class="main_memfiled_right">
        <ul>
		<%
		
		
		
		
		arr_result = excuteSqlStatementReadQuick("select sn,sname from lesson_industry", CONST_VIPABC_RW_CONN)
		'Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
'		Response.Write("錯誤訊息" & g_int_run_time_err_code & "<br>")
'		Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")


		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) > 0) then   
			'有資料
				For i = 0 to Ubound(arr_result,2)
					str_check = ""
					int_industry_sn = Trim(arr_result(0,i))
					str_industry_name = Trim(arr_result(1,i))
					
					if ( not isEmptyOrNull(str_industry) ) then 
						if (InstrRev(str_industry,"-"&trim(int_industry_sn)&"-") >0) then
							str_check = " checked=""checked"""
						end if 
					end if 
					
					
				%>
                <li><input type="checkbox" name="industry" id="checkbox" value='<%=int_industry_sn%>' <%=str_check%>/><%=str_industry_name%></li>

				<%
				Next
			else
			'無資料
			end if 
		end if 
		%>
        </ul>
        <div class="clear"></div>
        </div>
        <div class="clear"></div>
        </div>
        <!--职业别 end-->
        <div class="main_memfiled_line"></div>
        <!--职位别 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("dcgs_11")%><br />(<%=getWord("dcgs_10")%>1 ~ 2 项)</div>
        <div class="main_memfiled_right">
        <ul>
		<%
		arr_result = excuteSqlStatementReadQuick("select sn,sname from lesson_role", CONST_VIPABC_RW_CONN)
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) > 0) then   
			'有資料
				For i = 0 to Ubound(arr_result,2)
					str_check = ""
					str_role_sn = Trim(arr_result(0,i))
					str_role_name = Trim(arr_result(1,i))
					
					if ( not isEmptyOrNull(str_role) ) then 
						if (InstrRev(str_role,"-"&trim(str_role_sn)&"-") >0) then
							str_check = " checked=""checked"""
						end if 
					end if 
					
					
				%>
                <li><input type="checkbox" name="role" id="checkbox" value='<%=str_role_sn%>' <%=str_check%>/><%=str_role_name%></li>

				<%
				Next
			else
			'無資料
			end if 
		end if 
		%>
        </ul>
        <div class="clear"></div>
        </div>
        <div class="clear"></div>
        </div>
        <!--职位别 end-->
        <div class="main_memfiled_line"></div>
        <!--兴趣 start-->
        <div class="main_memfiled_all">
        <div class="main_memfiled_left"><%=getWord("dcgs_12")%><br />
          (<%=getWord("dcgs_10")%>3 ~ 5 项)</div>

        <div class="main_memfiled_right">
        <ul>
        <%
        '201304017 阿捨修正 21 22 for VIPABCJr
		arr_result = excuteSqlStatementReadQuick(" SELECT sn, sname FROM cfg_interest WHERE sn <> '21' AND sn <> '22' ", CONST_VIPABC_RW_CONN)
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) > 0) then   
			'有資料
				For i = 0 to Ubound(arr_result,2)
					str_check = ""
					str_interest_sn = Trim(arr_result(0,i))
					str_interest_name = Trim(arr_result(1,i))
					
					if ( not isEmptyOrNull(str_interest) ) then 
						if (InstrRev(str_interest,"-"&trim(str_interest_sn)&"-") >0) then
							str_check = " checked=""checked"""
						end if
					end if 
					 
				%>
                <li><input type="checkbox" name="interest" id="checkbox" value='<%=str_interest_sn%>' <%=str_check%>/><%=str_interest_name%></li>

				<%
				Next
			else
			'無資料
			end if 
		end if 
		%>
        </ul>
        <div class="clear"></div>
        </div>
        <div class="clear"></div>
        </div>
        <!--兴趣 end-->
        <div class="main_memfiled_line"></div>
        <div class="t-right" style='margin-left:110px;'><input type="submit" name="button" id="button" value="+ <%=getWord("BUY_1_40")%>" class="btn_1"/></div>
        </div>
        </form>
        
        
        <div class="top_block"></div>
        <!--勾選end-->
        <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
        </div>
        <!--內容end-->
        
