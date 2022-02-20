<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_vocbulary.asp" -->
<script type="text/javascript">
<!--
//只要call這隻都是已經完成作業了
$(document).ready(function(){
    $("input:radio").attr("disabled","disabled");
    //上方的button CSS變更
    $("#a_learning_review").css({ color: "#FFFFFF"}); 
    $("#a_learning_review").removeClass();
    $("#a_learning_review").addClass("top_class_N3");
})
//-->
</script>	
  <div class="main">
    <div class="main_con">
        <div class="main_membox">
<!--上版位課程start-->
	<!--#include virtual="/lib/include/html/include_learning.asp"--> 
<!--上版位課程end-->	
<% 
	'=================開始將表頭列出來======================
	Dim sql 'sql 語法
	Dim g_arr_result '陣列接收資料庫回傳值
	Dim session_sn : session_sn = request("session_sn")'該堂課
	Dim client_sn : client_sn = request("client_sn")'該學生
	Dim g_int_row, g_int_col '列、欄
	Dim i '變數	
	session_sn = request("session_sn")
	'如果沒有client_sn 則去抓session中的sn
	if client_sn="" then
		client_sn = session("client_sn")
	else
		client_sn = request("client_sn")
	end if
	'列出所有上課的相關資訊 課程代號 顧問 標題 上課時間 其他相關資訊 主要！
	sql="SELECT client_attend_list.Session_sn, dbo.client_attend_list.client_sn, dbo.client_attend_list.Attend_Date, dbo.client_attend_list.Attend_Sestime, dbo.client_attend_list.Attend_consultant, dbo.client_attend_list.Attend_mtl_1, dbo.client_attend_list.Attend_Level, dbo.con_basic.basic_fname + ' ' + dbo.con_basic.basic_lname AS Expr1, dbo.material.LTitle, dbo.material.LD, dbo.client_homework.Score, dbo.client_homework.anser,dbo.client_homework.rdn, dbo.client_homework.Sn, dbo.client_homework.voc_sn FROM dbo.client_attend_list LEFT OUTER JOIN dbo.client_homework ON dbo.client_attend_list.client_sn = dbo.client_homework.Client_Sn AND dbo.client_attend_list.Session_sn = dbo.client_homework.Session_Sn LEFT OUTER JOIN dbo.material ON dbo.client_attend_list.Attend_mtl_1 = dbo.material.Course LEFT OUTER JOIN dbo.con_basic ON dbo.client_attend_list.Attend_consultant = dbo.con_basic.Con_Sn WHERE (dbo.client_attend_list.Session_sn = '"&session_sn&"') AND (dbo.client_attend_list.client_sn = "&client_sn&")"
	g_arr_result = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)
	'如果存取成功
	if (isSelectQuerySuccess (g_arr_result)) then
		For g_int_row = 0 To Ubound(g_arr_result, 2)
			For g_int_col = 0 To Ubound(g_arr_result)
				'Response.write(g_arr_result(g_int_col, g_int_row) & "<br>")    ' col代表筆數, row代表欄位
			Next
		Next
	else
	'錯誤訊息
		Response.Write("錯誤訊息" & g_arr_result & "<br>")
	end if
	'根據資料庫 定義各表頭需要欄位
	Dim title : title = g_arr_result(8,0)'該堂課的標題
	Dim answer : answer = g_arr_result(11,0)'答案
	Dim voc_sn : voc_sn = g_arr_result(14,0)'
	Dim sessionTime : sessionTime = g_arr_result(2,0)&" - "&g_arr_result(3,0)&":30" '上課時間
'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  Start '''''''''''''''
		sessionTime1=sessionTime
	    arrsessionTime = Split(getSessionWholeTime(session_sn,3,0,CONST_TUTORABC_RW_CONN), " ")
		sessionTime = arrsessionTime(0) &" - " &arrsessionTime(1)
'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  End '''''''''''''''
	Dim consultant : consultant = g_arr_result(7,0) '上課的老師
	Dim level : level = g_arr_result(6,0) '該學生的程度
	'Dim client_level : client_level = rs("attend_level")'該學生的程度
	Dim description : description = g_arr_result(9,0) '描述
	Dim course : course = g_arr_result(5,0) '課名序號
	'=========讀取作答資訊====start============
	Dim answer_arr_result '接受作答的資料
	Dim client_vocsn, client_answer, client_point, client_rnd_order '釋義sn、學生作答、學生分數、亂數釋義順序
	Dim arr_c_vocsn, arr_c_answer, arr_rnd_order '關於學生的釋義sn、學生作答、學生分數 各陣列
	sql = "select * from client_homework where Valid = '1' And session_sn = '"&session_sn&"' And Client_sn = '"&client_sn&"'" 
	answer_arr_result = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)
	'將相關作答資訊存入陣列
	if (isSelectQuerySuccess (answer_arr_result)) then
			if (Ubound(answer_arr_result)>0) then
				client_vocsn = answer_arr_result(9,0)
				client_answer = answer_arr_result(10,0)
				client_point = answer_arr_result(12,0)
				client_rnd_order = answer_arr_result(4,0)
			end if
		else
		'錯誤訊息
			Response.Write("錯誤訊息" & questions_arr_result & "<br>")
	end if
	arr_c_vocsn = split(client_vocsn,",")
	arr_c_answer = split(client_answer,",")
	arr_rnd_order = split(client_rnd_order,",")
	'=========讀取作答資訊====end==============				
%>        
        
		<!--list start-->
        <div class="hwork">
        <!--table start-->
        <div class="wlist2">
        <div class="title">Session Information<div class="tline"></div></div>
        <div>
        <div class="linebox">
        <div class="plt">Course #</div>
        <div class="prt"><%=course%></div>
        <div class="clear"></div>
        </div>
        <div class="linebox">
        <div class="plt">Title</div>
        <div class="prt"><%=title%></div>
        <div class="clear"></div>
        </div>
        <div class="linebox">
        <div class="plt">Session Time</div>
        <div class="prt"><%=sessiontime%></div>
        <div class="clear"></div>
        </div>
        <div class="linebox">
        <div class="plt">Consultant</div>
        <div class="prt"><%=Consultant%></div>
        <div class="clear"></div>
        </div>
        <div class="linebox">
        <div class="plt">Level</div>
        <div class="prt">
        <!--level-->
		<%
		'level等級圖案 總共12級
		for i =1 to 12
			if i<level then
				response.write "<p class=""green_ballNum2""></p>"
			else
				if i=level then
					response.write "<p class=""green_ballNum_go"">"&i&"</p>"
				else
					response.write "<p class=""gray_ballNum2"">"&i&"</p>"
				end if
			end if
		next
		%>        
        <!--level-->
        </div>
        <div class="clear"></div>
        </div>
        <div class="linebox">
        <div class="plt">Description</div>
        <div class="prt"><%=description%></div>
        <div class="clear"></div>
        </div>
        </div>
        </div>
        <!--table end-->
		

				<%
		'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  Start '''''''''''''''
			checkcalsstype = isShortSession(session_sn, "", 0, CONST_TUTORABC_RW_CONN) '判斷是否為短課程，若不是，回傳false
			dim intQnum:intQnum=5
			if false <> checkcalsstype then  '短課程:3題 ,一般課程:5題
				intQnum=3			
			end if
		'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  End '''''''''''''''	

		%>
		<form name='form' method="POST" action="">
        <!--quest start-->
        <div class="wlist3"><!--這裡頭的radio button應該是灰色不可點選的樣子-->
		
		<div class="title">Your score is<span class="score"><%=client_point%></span></div>
		
        <!--question 1~5-->
		<div class="ps">Please answer these questions: The correct answer is displayed in Orange.</div>
<%
	'=================開始將題目列出來======================
	Dim questions_arr_result '接收題目的相關資訊陣列
	Dim questions_arr '題目陣列
	Dim choice_arr '選項陣列 二維陣列 一筆題目會有五個選項
	Dim questions_answer '題目的答案陣列
	Dim choices_title : choices_title = array("a.","b.","c.","d.") '選項的標題
	Dim j '變數
 	'列出該堂課相關題目，正常會有五筆，分別為題目 與 選項
	sql="SELECT TOP " &intQnum& " * FROM lesson_question WHERE (Course = '"&course&"') and (Question <> '')"
	'response.write "sql="&sql
	questions_arr_result = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)
	'動態宣告陣列大小
	Redim questions_arr(Ubound(questions_arr_result, 2))
	Redim choice_arr(Ubound(choices_title),Ubound(questions_arr_result, 2)) '題目選項最多為四個
	Redim questions_answer(Ubound(questions_arr_result, 2))
	'如果存取成功 則將題目相關資訊存入陣列
	if (isSelectQuerySuccess (questions_arr_result)) then
		For g_int_row = 0 To Ubound(questions_arr_result, 2)
			'將問題塞入陣列
			questions_arr(g_int_row) = questions_arr_result(6,g_int_row)
			'選項為7～10分別為choice1~4
			For g_int_col = 7 To 10
				'將選項塞入陣列 初始值為0 故減7 
				choice_arr(g_int_col-7,g_int_row) = questions_arr_result(g_int_col,g_int_row)
				'============重要！將正確答案轉換成數字選項存入答案陣列中====satrt========
				If questions_arr_result(g_int_col,g_int_row) = questions_arr_result(13,g_int_row) then
					'如果與答案相等 將idex-6回存 因為資料庫中的紀錄答案欄位為數字 故要作此轉換 且a、b、c、d分別代表1、2、3、4
					questions_answer(g_int_row) = g_int_col-6
				End if				
				'============重要！將正確答案轉換成數字選項存入答案陣列中====end========
			Next
		Next
	else
	'錯誤訊息
		Response.Write("錯誤訊息" & questions_arr_result & "<br>")
	end if
%>
		
		<%
			'------------------------上半部作答區------start-------------
			'==============批改作業====================================	
			For i=0 to Ubound(questions_arr)
		%>
		<P>
		<%
		'=======判斷題目答對答錯得圖片=========start========
			if questions_answer(i) = Fix(arr_c_answer(i) )then
		%>
			<img src="/images/images_cn/yes.gif" width="20" height="20" />
		<%
			else			
		%>
			<img src="/images/images_cn/wrong.gif" width="20" height="20" />
		<%
			end if	
		'=======判斷題目答對答錯得圖片=========end========	
		%>
		<!--題目-->
		<%=(i+1)&"."&questions_arr(i)%>
		<br>
			<!--隱藏的input 把正確解答放進來-->
			<input type="hidden" name="right_answer_<%=i%>" id="right_answer_<%=i%>" value="<%=questions_answer(i)%>" />
			<%
				For j=0 to Ubound(choices_title) '選項最多為三個 
					if  j+1 = Fix(arr_c_answer(i))then					
			%>
					
					<input type="radio" name="answer_<%=i%>" id="answer_<%=i%>" value="<%=j+1%>" checked />
				<%
					else
				%>
					<input type="radio" name="answer_<%=i%>" id="answer_<%=i%>" value="<%=j+1%>" />
				<%
					end if
				%>
				<span class="under"><!--選項標題--><%=choices_title(j)%>
				</span>	
				<!--顯示正確的解答start-->
				<%
					if(questions_answer(i)=j+1) then
				%> 
					<!--正確選項--><span class="answer"><%=choice_arr(j,i)%></span>	
				<%
					else
				%>
					<!--選項--><%=choice_arr(j,i)%>	
				<%				
					end if
				<!--顯示正確的解答 end-->
				next
			%> 
		</P>
		<%
			next
			'------------------------上半部作答區------end-------------
		%>
        <!--連連看 start-->
          <div class="selct">
          <div class="tle">Please find the correct definition</div>
			  <div class="p-bottom15 ">
			  <!--左邊s-->
			  <div class="lt ">
				<%			  
					Dim q_def_arr_result '用來接收連連看定義的資料陣列
					Dim q_def_arr '定義單字陣列
					Dim q_explain_arr '單字解釋陣列
					Dim q_def_title : q_def_title = array("A","B","C","D","E")	
					Dim explain_lan '解釋的語系
					Dim input_checked '讓input選項被點選
					Dim explain_right_answer : explain_right_answer = array("","","","","")'釋義的正確答案陣列
						if 3 = intQnum then '短課程需要定義不同長度陣列
							q_def_title = array("A","B","C")		
							explain_right_answer = array("","","")
						end if
					'選出連連看的題目單字 只需要"五"筆？ valid必須為1且中文英文的釋義都需存在
					sql="SELECT TOP " & intQnum & "sn, vocabulary AS vocabulary_keyword, definition FROM material_vocabulary WHERE (course = "&course&" AND c_dictionary_state = '1'  AND valid = '1') OR (course = "&course&" AND e_dictionary_state = '1' AND valid = '1') or (course = "&course&" AND is_checked = '1' AND valid = '1')"
					q_def_arr_result = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)
					Redim q_explain_arr(Ubound(q_def_arr_result,2))
					Redim q_def_arr(Ubound(q_def_arr_result,2))
					'將值塞入單字與單字解釋陣列中
					For i=0 to Ubound(q_def_arr_result,2)
						'將單字塞入q_def_arr陣列中
						q_def_arr(i) = q_def_arr_result(1,i)
						'將單字解釋塞入q_explain_arr陣列中
						q_explain_arr(i) = q_def_arr_result(2,i)
						'根據單字來搜尋解釋
						'若程度level大於4級 以英文撰寫 其餘中文
						if (level>4) then
							'解釋為英文
							explain_lan = "en"
						else
							'解釋為中文
							explain_lan = "zh-cn"
						end if
												
						'20111121 add by JosieWu RD11110708 - VIPABC課後作業取單字解釋的邏輯與TutorABC相同 Begin
						'單字解釋的取得 step1. db -> step2. getVocabularyExplain -> step3. read txt
						if isEmptyOrNull(q_explain_arr(i)) then
							
							Set obj_mtl_voc_opt = Server.CreateObject("TutorLib.MaterialOpt.MaterialVocabularyOpt")
							q_explain_arr(i)  = obj_mtl_voc_opt.getVocabularyExplain(q_def_arr_result(1,i), explain_lan)

							Set obj_mtl_voc_opt = Nothing 
				
						end if
						
						if isEmptyOrNull(q_explain_arr(i)) then
							Set objFSO = CreateObject("Scripting.FileSystemObject")
							'***字彙處理開始*********************************************************************
							IF client_level>1 and client_level<4  Then
									 
								IF isEmptyOrNull(q_explain_arr(i)) then
								   v_path_str=CONST_TUTORABC_DISK_WEB_ADDR&"\www.tutorabc.com\tools\cdic\homework_vocabulary\en-text\" & q_def_arr_result(1,i) & ".txt"
								   Set objFSO = CreateObject("Scripting.FileSystemObject") 
								   IF objFSO.FileExists(v_path_str) Then
									  Set objTextFile = objFSO.OpenTextFile (v_path_str, 1,false,-1)
									  q_explain_arr(i)=objTextFile.readall
								   End IF
								   Set objTextFile=Nothing
								End IF
							Else
								v_path_str=CONST_TUTORABC_DISK_WEB_ADDR&"\www.tutorabc.com\tools\cdic\homework_vocabulary\en-text\" & q_def_arr_result(1,i) & ".txt"
								Set objFSO = CreateObject("Scripting.FileSystemObject") 
								IF objFSO.FileExists(v_path_str) Then
								   Set objTextFile = objFSO.Opentextfile(v_path_str,1,false,-1)
								   q_explain_arr(i)=objTextFile.readall
								End IF
								Set objTextFile=Nothing
								
							End IF
							Set objFSO = Nothing
						end if 
						'20111121 add by JosieWu RD11110708 - VIPABC課後作業取單字解釋的邏輯與TutorABC相同 End
					Next
				'================重要！紀錄下正確的釋義答案============start=========
				
					for i=0 to Ubound(q_def_arr_result,2)
						for j=0 to Ubound(q_def_title)
							If q_explain_arr(arr_rnd_order(j)) = q_explain_arr(i) then
								explain_right_answer(i) = q_def_title(j)
							end if
						next
					next
				'================重要！紀錄下正確的釋義答案============end=========	
				%>
				<%
					if 3 = intQnum then istart=4 else istart = 6
					'----------------------下半部連連看---start---------------------------------------------------
					For i=0 to Ubound(q_def_arr_result,2)
				%>
					<!--1s-->
					<div class="subj">
					<%
						if (explain_right_answer(i) = arr_c_answer(i+5)) then
					%>
						<img src="/images/images_cn/yes.gif" width="20" height="20"  />
					<%
						else
					%>
						<img src="/images/images_cn/wrong.gif" width="20" height="20" />
					<%
						end if
					%>
					<!--序號--><%=i+istart&"."%>
					<span class="subj-font"><!--單字--><%=q_def_arr(i)%></span>
					</div>
					<div class="subj-as">
					<%
						For j=0 to Ubound(q_def_title)
							if (arr_c_answer(i+5) = q_def_title(j)) then
								input_checked = "checked"
							else
								input_checked = ""
							end if							
					%>
							<input type="radio" name="answer_<%=i+5%>" id="answer_<%=i+5%> " value="<%=q_def_title(j)%>" <%=input_checked%> />
						
						<%
						'==============判斷哪個是正確的選項 顯示出來start=======================
							if (explain_right_answer(i) = q_def_title(j)) then
						%>
							<!--正確選項--><span class="answer"><%=q_def_title(j)%></span>
						<%
							else
						%>
						<font >	<!--選項--><%=q_def_title(j)%> </font>
						<%
							end if
							'==============判斷哪個是正確的選項 顯示出來end=======================
						%>
					<%
						Next
					%>
					<!--正確解答-->
					<input type="hidden" name="right_answer_<%=i+5%>" id="right_answer_<%=i+5%>" value="<%=explain_right_answer(i)%>" />
					</div>
				<!--左邊e-->
				<%
					Next
				%>
			</div>
			  <!--右邊s-->
			<div class="rt" >
			  <!--1s-->
				<%
					For i=0 to Ubound(q_explain_arr)
				%>
				  <div class="cho">
				  
				  <h1><!--解釋標題--><%=q_def_title(i)%></h1>
				  <h2><!--單字解釋 亂數排列--><%=q_explain_arr(arr_rnd_order(i))%></h2>
				 
				  <div class="clear"></div>
				  </div>
				<%
					Next
				%>
			  <!--1e-->
			
			
			  </div>
			  <!--右邊e-->
			  <div class="clear"></div>
			  </div>
          
          </div>
        <!--連連看 end--> 
		<div class="bline">
			<input type="button" name="button2" id="button2" value="+ 回上页" class="btn_1" onclick="window.location = 'learning_review.asp'" />
			<input type="button" class="btn_1 m-left5" value="+ 提交" id="send_homework" name="button2" disabled>
		</div>
        </div>
          <%
			  '----------------------下半部連連看---end---------------------------------------------------
		  %>
        <!--quest end-->
		</form>
        </div>
         
        </div>
        <!--list end-->
        
        </div>
        <!--內容end-->
        </div>      
	