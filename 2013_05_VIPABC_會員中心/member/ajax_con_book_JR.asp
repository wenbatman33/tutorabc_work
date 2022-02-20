<!--#include virtual="/lib/include/global.inc"-->
<!--教材資訊-->
 <%
	Dim str_session_sn : str_session_sn = request("input_session_sn")
	Dim str_material : str_material = request("input_material")
	Dim str_today_session : str_today_session = request("input_today_session")
	Dim var_arr, str_sql, arr_result, obj_fs
	Dim str_material_rating, str_material_ltitle, str_material_ld, str_attend_level, str_foldername
	'如果有教材 且 今天有課程
	Dim str_stimeHM : str_stimeHM = ""	'系統資料上課 年月日時分秒
	Dim ntimeHM
	ntimeHM = Now()	'抓取現在的時間

	if (str_material <> "" and str_today_session <> "") then 
	 
		 var_arr = Array(str_session_sn,str_material,str_material)
		 'str_sql = "select " 
		 'str_sql = str_sql & " client_atteind_list.attend_mtl_1,material_rating.rating, "   
		 'str_sql = str_sql & " material.ltitle,material.ld,client_atteind_list.attend_level, "
		 'str_sql = str_sql & " client_atteind_list.session_sn "
		 'str_sql = str_sql & " from ( "
		 'str_sql = str_sql & " 	 select attend_mtl_1,session_sn,attend_level "
		 'str_sql = str_sql & " 	 from client_attend_list where session_sn = @session_sn "
		 'str_sql = str_sql & " ) as client_atteind_list"
		 'str_sql = str_sql & " left join ( "
		'						'--rating
		 'str_sql = str_sql & " 	SELECT client_attend_list.attend_mtl_1, "
		 'str_sql = str_sql & "  LEFT((SUM(client_session_evaluation.materials_points) + 0.0) / COUNT(client_session_evaluation.sn), 4) AS rating "
		 'str_sql = str_sql & "  FROM client_session_evaluation "
		 'str_sql = str_sql & " 	LEFT JOIN (  "
		 'str_sql = str_sql & " 		SELECT DISTINCT session_sn, attend_mtl_1 "
		 'str_sql = str_sql & "		FROM client_attend_list  "
		 'str_sql = str_sql & " 		WHERE (valid = 1) and attend_mtl_1 =@attend_mtl_1 "
		 'str_sql = str_sql & " 	) AS client_attend_list "
		 'str_sql = str_sql & " ON client_session_evaluation.session_sn = client_attend_list.session_sn  "
		 'str_sql = str_sql & " GROUP BY client_attend_list.attend_mtl_1  "
		 'str_sql = str_sql & " )as material_rating "
		 'str_sql = str_sql & " on client_atteind_list.attend_mtl_1 = material_rating.attend_mtl_1 "
		 'str_sql = str_sql & " left join ( "
		 'str_sql = str_sql & " 	select course,ltitle,ld from material where course = @course "
		 'str_sql = str_sql & " ) as material "
		 'str_sql = str_sql & " on material.course = client_atteind_list.attend_mtl_1 "
		
         str_sql = " SELECT " 
		 str_sql = str_sql & " ClassRecord.ClassMaterialNumber, material_rating.rating, "   
		 str_sql = str_sql & " material.ltitle, material.ld, ClassRecord.ProfileLearnLevel, "
		 str_sql = str_sql & " ClassRecord.session_sn "
		 str_sql = str_sql & " FROM ( "
		 str_sql = str_sql & " 	 SELECT ClassMaterialNumber, session_sn, ProfileLearnLevel "
		 str_sql = str_sql & " 	 FROM member.ClassRecord WHERE session_sn = @session_sn "
		 str_sql = str_sql & " ) AS ClassRecord "
		 str_sql = str_sql & " LEFT JOIN ( "
		 str_sql = str_sql & " 	SELECT ClassRecord.ClassMaterialNumber, "
		 str_sql = str_sql & "  LEFT((SUM(client_session_evaluation.materials_points) + 0.0) / COUNT(client_session_evaluation.sn), 4) AS rating "
		 str_sql = str_sql & "  FROM client_session_evaluation "
		 str_sql = str_sql & " 	LEFT JOIN (  "
		 str_sql = str_sql & " 		SELECT DISTINCT session_sn, ClassMaterialNumber "
		 str_sql = str_sql & "		FROM member.ClassRecord  "
		 str_sql = str_sql & " 		WHERE ( ClassStatusId = 2 OR ClassStatusId = 3 ) AND (ClassRecordValid = 1) AND ClassMaterialNumber =@ClassMaterialNumber "
		 str_sql = str_sql & " 	) AS ClassRecord "
		 str_sql = str_sql & " ON client_session_evaluation.session_sn = ClassRecord.session_sn  "
		 str_sql = str_sql & " GROUP BY ClassRecord.ClassMaterialNumber  "
		 str_sql = str_sql & " ) AS material_rating "
		 str_sql = str_sql & " ON ClassRecord.ClassMaterialNumber = material_rating.ClassMaterialNumber "
		 str_sql = str_sql & " LEFT JOIN ( "
		 'str_sql = str_sql & " 	SELECT course, ltitle, ld FROM material WHERE course = @course "
         str_sql = str_sql & " 	SELECT abcjr.dbo.lesson_plan_Temp.course, abcjr.dbo.lesson_plan_Temp.title AS ltitle, abcjr.dbo.lesson_description_temp.description AS ld FROM abcjr.dbo.lesson_plan_Temp "
         str_sql = str_sql & "   LEFT JOIN abcjr.dbo.lesson_description_temp ON abcjr.dbo.lesson_plan_Temp.course = abcjr.dbo.lesson_description_temp.course "
         str_sql = str_sql & "   WHERE abcjr.dbo.lesson_plan_Temp.course = @course "
		 str_sql = str_sql & " ) AS material "
		 str_sql = str_sql & " ON material.course = ClassRecord.ClassMaterialNumber "

		arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)    
		'response.write g_str_sql_statement_for_debug
		if (isSelectQuerySuccess(arr_result)) then
		   
			if (Ubound(arr_result) >= 0) then
				str_material_rating = arr_result(1,0)
				str_material_ltitle= arr_result(2,0)
				str_material_ld = arr_result(3,0)
				str_attend_level = arr_result(4,0)	
				'20110314 tree RD11031101 JR和VIPABC加入預覽教材功能 儲存系統的年月日時分秒
				str_stimeHM = Mid(arr_result(5,0),1,4) + "/" + Mid(arr_result(5,0),5,2) + "/" + Mid(arr_result(5,0),7,2) + " " + Mid(arr_result(5,0),9,2) + ":30:00"
			else
			
			end if 
		end if 
	 '-------------- 判斷是否有檔案存在 -------------------
		str_foldername = "E:\web\www.vipabc.com\images\mtl_thumb\" & str_material & ".jpg"
		Set obj_fs = CreateObject ("Scripting.FileSystemObject")
		If obj_fs.FileExists (str_foldername) Then
			str_foldername = str_material & ".jpg"
		else
			str_foldername = "mov.jpg"
		end if 
		Set obj_fs = Nothing
		
	end if

    
    '---客戶資料 COM 物件
    Dim obj_member_opt_exe : Set obj_member_opt_exe = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
    Dim int_member_result_exe : int_member_result_exe = obj_member_opt_exe.prepareData(g_var_client_sn, CONST_VIPABC_RW_CONN)

    if (int_member_result_exe = CONST_FUNC_EXE_SUCCESS) then        
	    intClientLevel = obj_member_opt_exe.getData(CONST_CUSTOMER_NOW_LEVEL) '取得客戶的等級
    end if
    set int_member_result_exe = nothing
    set obj_member_opt_exe = nothing
%>
 <!--
 <div class="book_left">
               
              
			   <img src="/images/mtl_thumb/<%=str_foldername%>" width="158" height="118" border="0" /><br />
                 Level : <span class="color_1"><%=intClientLevel%></span> Avg Rating : <span class="color_1"><%=str_material_rating%></span></div>-->
               <div class="book_right"><span class="big bold"><%=str_material_ltitle%></span><br />
                 <%=getWord("CLASS_15")%>：<!-- /
                 <select name="select2" id="select2" class="login_box4 width3">
                   <option><%=getWord("CLASS_17")%></option>
                 </select>--><br /><br />
                 <%=str_material_ld%><br /><br />
					<input type="submit" name="button2" id="button2" value="+<%=getWord("CLASS_12")%>" class="btn_1 m-left5" onclick="javascript:location.href='class_vocabulary.asp';"/>
					<input type="button" name="button2" id="button2" value="+ <%=getWord("CLASS_14")%>" class="btn_1 m-left5" style="display:none" onclick="javascript:location.href='class_notice.asp';"/>
					
					<%
					'20110314 tree RD11031101 JR和VIPABC加入預覽教材功能---開始---
					'(下次上課時間 和 現在系統時間的時間秒差) >= -1800秒
					%>
					<%if ( DateDiff("S", str_stimeHM,  ntimeHM ) >= -1800 ) then %>
						<input type="submit" name="button3" id="button3" value="+教材预览" class="btn_1 m-left5" onclick='dialog_open(); return false;'/>
					<%end if%>	
					<%
					'20110314 tree RD11031101 JR和VIPABC加入預覽教材功能---結束---
					%>

			   </div>  