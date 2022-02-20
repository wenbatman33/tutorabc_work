
<!--#include virtual="/lib/include/html/template/template_change.asp"-->

<div id="temp_contant">
       
       
       <!--內容start-->
        <div class="main_membox">
        
        <div class='page_title_3'><h2 class='page_title_h2'>程度分析</h2></div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 
        <!--上版位課程start-->
        <!--#include virtual="/lib/include/html/include_learning.asp"--> 
        <!--上版位課程end-->
        <!--程度分析-->
        
         <%
		 '求得要秀字樣的陣列值
		function getArrayString(ByVal p_str_strel) 
			Dim int_strel_num 
			Dim int_array_num
			int_strel_num = p_str_strel / 2
			'無條件進位
			If (InStr(int_strel_num,".")=0) Then
				int_array_num=int_strel_num
			Else
				int_array_num=cint(Left(cstr(int_strel_num),InStr(1,cstr(int_strel_num),".")-1))+1
			end If
			getArrayString = int_array_num
		end function 

		redim stre1(5)
		redim stre2(5)
		
		
		
		Dim str_sql : str_sql = ""
		Dim i
		Dim evaluation_type : evaluation_type =""
		Dim client_level : client_level =""

		Dim str_client_fname : str_client_fname = Session("client_fname")
		
		Dim c : c= 0
		Dim gp : gp= 0
		Dim lp : lp= 0
		Dim pp : pp= 0
		Dim sp : sp= 0
		Dim vp : vp= 0
		Dim rp : rp= 0
		Dim placementt : placementt = 0 
		'塞資料的相關設定值
		Dim var_arr 								'傳excuteSqlStatementRead的陣列
		Dim arr_result							'接回來的陣列值
		Dim str_tablecol 							'table欄位
		Dim str_tablevalue 						'table值
		
		'------------- 模擬考部份 -----------------------
		Dim int_act_i
		Dim str_act_inpdate   '輸入時間
		Dim str_act_message3	'結束時間
		Dim str_act_message2	'分數
		Dim str_act_note		'題目
			
		Dim Dinpdate			'模擬考測驗日期
		Dim Hinpdate			'模擬考開始--時
		Dim Minpdate			'模擬考開始--分
		
		Dim HDone				'模擬考結束--時
		Dim MDone				'模擬考結束--分
		Dim Done				'模擬考結束--時:分
		Dim DoneSocre			'模擬考分數
		Dim DDone				'之前沒有紀錄的加45分鐘
		
		if session("productno")=40 then 
			'歌倫比亞客戶
			str_sql = "SELECT sn,listening, speaking, reading, pronunciation, grammar, vocabulary  "
			str_sql = str_sql & " FROM columbia.dbo.client_basic(nolock) "
			str_sql = str_sql & " RIGHT JOIN (SELECT columbia_sn FROM client_basic(nolock) WHERE (Sn = @sn)) DERIVEDTBL"
			str_sql = str_sql & " ON client_basic.sn = DERIVEDTBL.columbia_sn "
				
			var_arr = Array(g_var_client_sn)
			arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
		
			if (isSelectQuerySuccess(arr_result)) then
			
				'--------- 塞適性測驗 ------------ start ------------
				if (Ubound(arr_result) > 0) then 
					'有資料
					
					stre1(0)=arr_result(5,0)	'Grammar
					stre1(1)=arr_result(6,0)  'Vocabulary
					stre1(2)=arr_result(4,0)  'Pronunciation
					stre1(3)=arr_result(2,0)	'Speaking
					stre1(4)=arr_result(3,0)	'Reading
					stre1(5)=arr_result(1,0)	'Listening
					
					placementt=6
					
				else
					'無資料
				end if 
				'--------- 塞適性測驗 ------------ end ------------
			end if    	
		
		else
			'TutorABC客戶 (最後一筆)
			str_sql = "select lesson_group.client_sn, lesson_group.expr1, isNull(lesson_question.client_level,0) client_level,  "
			str_sql = str_sql & " isNull(lesson_question.evaluation_type,0) evaluation_type  "
			str_sql = str_sql & " from "
			str_sql = str_sql & " 	(select client_sn, max(test_sn) as expr1  "
			str_sql = str_sql & " 	from "
			str_sql = str_sql & " 		(select client_sn, test_sn  "
			str_sql = str_sql & " 		from lesson_question_temp(nolock)  "
			str_sql = str_sql & " 		group by client_sn, test_sn   "
			str_sql = str_sql & " 		having (count(sn) >= 5) and client_sn=@client_sn  "
			str_sql = str_sql & " 	) as lesson_temp  "
			str_sql = str_sql & " 	group by client_sn having (client_sn = @client_sn_1)  "
			str_sql = str_sql & " 	) as lesson_group   "
			str_sql = str_sql & " 	left join lesson_question_temp(nolock) as lesson_question   "
			str_sql = str_sql & " 	on lesson_group.expr1 = lesson_question.test_sn  "
		
			var_arr = Array(g_var_client_sn,g_var_client_sn)
		
			arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
		
		
			if (isSelectQuerySuccess(arr_result)) then
		
				'--------- 塞適性測驗 ------------ start ------------
				if (Ubound(arr_result) > 0) then 
					'有資料
					placementt = Ubound(arr_result,2) + 1
					'placementt = 1
					for i = 0 to Ubound(arr_result,2) 
					
						evaluation_type = arr_result(3,i)
						client_level = arr_result(2,i)
						
						if evaluation_type=1 then stre1(0)=client_level	 'Grammar
						if evaluation_type=2 then stre1(1)=client_level  'Vocabulary
						if evaluation_type=3 then stre1(2)=client_level  'Pronunciation
						if evaluation_type=4 then stre1(3)=client_level	 'Speaking
						if evaluation_type=5 then stre1(4)=client_level	 'Reading
						if evaluation_type=6 then stre1(5)=client_level	 'Listening
		
					next
					
				else
					'無資料
				end if 
				'--------- 塞適性測驗 ------------ end ------------
			end if    	
		
		end if 	
		
		
		'若沒值，表示此客戶並未做過適性測驗
		if stre1(0) = "" then stre1(0) = 0
		if stre1(1) = "" then stre1(1) = 0
		if stre1(2) = "" then stre1(2) = 0
		if stre1(3) = "" then stre1(3) = 0
		if stre1(4) = "" then stre1(4) = 0
		if stre1(5) = "" then stre1(5) = 0	
			
		
		
		'-------------------- 尚未完成適性測驗的時候 --------------- start -------------
		'response.write placementt
		if (placementt < 6) then
	   	
			Call alertGo("", "learning_level.asp", CONST_NOT_ALERT_THEN_REDIRECT )
			Response.end
		'-------------------- 尚未完成適性測驗的時候 --------------- end --------------
		else
		
			Dim arr_g_string 	'文法
			Dim arr_v_string 	'字彙
			Dim arr_p_string 	'發音
			Dim arr_l_string 	'聽力
			Dim arr_r_string  '閱讀
			
			reDim arr_g_string(5)
			arr_g_string(0)= getWord("evaluation_grammar_string_1")
			arr_g_string(1)= getWord("evaluation_grammar_string_2")
			arr_g_string(2)= getWord("evaluation_grammar_string_3")
			arr_g_string(3)= getWord("evaluation_grammar_string_4")
			arr_g_string(4)= getWord("evaluation_grammar_string_5")
	 
			reDim arr_v_string(5)
			arr_v_string(0)= getWord("evaluation_word_string_1")
			arr_v_string(1)= getWord("evaluation_word_string_2")
			arr_v_string(2)= getWord("evaluation_word_string_3")
			arr_v_string(3)= getWord("evaluation_word_string_4")
			arr_v_string(4)= getWord("evaluation_word_string_5")
	 
			reDim arr_p_string(5)
			arr_p_string(0)= getWord("evaluation_pronounce_string_1")
			arr_p_string(1)= getWord("evaluation_pronounce_string_2")
			arr_p_string(2)= getWord("evaluation_pronounce_string_3")
			arr_p_string(3)= getWord("evaluation_pronounce_string_4")
			arr_p_string(4)= getWord("evaluation_pronounce_string_5")
			
			reDim arr_l_string(5)
			arr_l_string(0)= getWord("evaluation_listening_string_1")
			arr_l_string(1)= getWord("evaluation_listening_string_2")
			arr_l_string(2)= getWord("evaluation_listening_string_3")
			arr_l_string(3)= getWord("evaluation_listening_string_4")
			arr_l_string(4)= getWord("evaluation_listening_string_5")
	 
			
			reDim arr_r_string(5)
			arr_r_string(0)= getWord("evaluation_read_string_1")
			arr_r_string(1)= getWord("evaluation_read_string_2")
			arr_r_string(2)= getWord("evaluation_read_string_3")
			arr_r_string(3)= getWord("evaluation_read_string_4")
			arr_r_string(4)= getWord("evaluation_read_string_5")
			%>
			
			<!--英語能力測試內容start-->
			<div class="main_classcontbox ">
			<!--英語能力測試過程列start-->
			<div class="main_classtest_processbar">
			<%=getWord("YOU_PLAY_NOW")%><br/>
			<ul class="process7">
			<li><%=getWord("WORD")%></li>
			<li><%=getWord("DCGS_9")%></li>
			<li><%=getWord("GRAMMAR")%></li>
			<li><%=getWord("DCGS_7")%></li>
			<li><%=getWord("CLASS_VOCABULARY_8")%></li>
			<li><%=getWord("FINISH")%></li>
			</ul>
			<div class="clear"></div>
			</div>
			<!--英語能力測試過程列end-->
			<!--測試結果內容start-->
			<div class="resultbox">
			<!--受測人名稱顯示start-->
			<div class="ana_name"><%=Trim(str_client_fname)%> ’s Strength Analysis</div>
			<!--受測人名稱顯示end-->
			<!--受測人結果量表start-->
			<div class="ranklist">
			<div class="ranktr1">
			<div class="ranklevel<%=stre1(0)%>"><%=getWord("GRAMMAR")%></div>
			<div class="clear"></div>
			</div>
			<div class="ranktr2">
			<div class="ranklevel<%=stre1(1)%>"><%=getWord("WORD")%></div>
			<div class="clear"></div>
			</div>
			<div class="ranktr2">
			<div class="ranklevel<%=stre1(3)%>"><%=getWord("CLASS_VOCABULARY_8")%></div>
			<div class="clear"></div>
			</div>
			<div class="ranktr2">
			<div class="ranklevel<%=stre1(5)%>"><%=getWord("DCGS_7")%></div>
			<div class="clear"></div>
			</div>
			<div class="ranktr2">
			<div class="ranklevel<%=stre1(4)%>"><%=getWord("DCGS_9")%></div>
			<div class="clear"></div>
			</div>
			</div>
			<!--受測人結果量表end-->
			<!--評量結果解說start-->
			<div class="explist">
	
			<div class="rescate"><%=getWord("GRAMMAR")%>:</div>
			<div class="rlevel"><%=stre1(0)%></div>
			<div class="clear"></div>
			<%%>
			<div class="restext"><%=arr_g_string(getArrayString(stre1(0)))%></div>
			<div class="rescate"><%=getWord("WORD")%>:</div>
			<div class="rlevel"><%=stre1(1)%></div>
			<div class="clear"></div>
			<div class="restext"><%=arr_v_string(getArrayString(stre1(1)))%></div>
			
			<div class="rescate"><%=getWord("CLASS_VOCABULARY_8")%>:</div>
			<div class="rlevel"><%=stre1(3)%></div>
			<div class="clear"></div>
			<div class="restext"><%=arr_p_string(getArrayString(stre1(3)))%></div>
			
			<div class="rescate"><%=getWord("DCGS_7")%>:</div>
			<div class="rlevel"><%=stre1(5)%></div>
			<div class="clear"></div>
			<div class="restext"><%=arr_l_string(getArrayString(stre1(5)))%></div>
			
			<div class="rescate"><%=getWord("DCGS_9")%>:</div>
			<div class="rlevel"><%=stre1(4)%></div>
			<div class="clear"></div>
			<div class="restext"><%=arr_r_string(getArrayString(stre1(4)))%></div>
			</div>
			<!--評量結果解說end-->
			<div class="clear"></div>
			</div>        
			<div class="d_contact"><%=getWord("real_man_test_description")%>  <input type="submit" name="button" id="button" value="+ <%=getWord("contact_us")%>" class="btn_1" onclick="location.href='../contact/contact_us.asp'"/></div>
			<!--測試結果內容end-->
			</div>
			<!--英語能力測試內容end-->      
		<%
		end if 
	   %>

        </div>
        <!--內容end-->
        
  <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>        
        
</div>
