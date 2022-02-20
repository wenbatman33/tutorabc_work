<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_vocbulary.asp" -->
<script type="text/javascript">
	function checkAnswer()
	{
		var bol_is_empty_answer;
		//是不是有空值的選項
		bol_is_empty_answer = true;
		$("input:radio").each(function ()
		{
			$(this).parent().removeClass("path");
			//抓出個別名稱
			var str_input_radio_name = $(this).attr("name");
			
			
			//檢查是否有答案沒選
			if ($("input[name=" + str_input_radio_name + "]:checked").size() <= 0 )
			{
				//錯誤的變色
	//			$(this).parent().addClass("path");
	//			bol_is_empty_answer = false;
			}
		});
		if (!bol_is_empty_answer)
		{
	//		alert("请确认答案是否都已经填写");
		}
		return bol_is_empty_answer;
	}
function viewScore(session_sn,client_sn)
{
	$.ajax({
			type: "POST",
			url: "ajax_myhomework_complete.asp",
			data: {"session_sn":session_sn,"client_sn":client_sn},
			beforeSend: function(){
			$.blockUI({ message: "<h1>count your score...</h1>"}); 	
			},
			error: function(xhr) {alert("錯誤"+xhr.responseText)},
			success: function(msg){
				//alert(msg);
				$(".main_membox").html(msg)
			},
			complete:function(){
			$.unblockUI(); 
			}
		});	
}

$(document).ready(function ()
{
    //上方的button CSS變更
    $("#a_learning_review").css({ color: "#FFFFFF" });
    $("#a_learning_review").removeClass();
    $("#a_learning_review").addClass("top_class_N3");
    //當送出時
    $("#send_homework").click(function ()
    {
        var temp_str_answer;
        var temp_str_r_answer;
        var temp_data;
        var i;
        //alert("yes");
        temp_str_answer = "";
        temp_str_r_answer = "";
        for (i = 0; i <= 9; i++)
        {
            temp_str_answer = temp_str_answer + "answer_" + i + "=" + $("input[name=answer_" + i + "]:checked").val() + "&";
            temp_str_r_answer = temp_str_r_answer + "right_answer_" + i + "=" + $("#right_answer_" + i).val() + "&";
        }
        temp_data = temp_str_answer + temp_str_r_answer + "session_sn=" + $("#session_sn").val() + "&" + "client_sn=" + $("#client_sn").val() + "&" + "vocsn=" + $("#vocsn").val() + "&" + "rnd_explain_order=" + $("#rnd_explain_order").val();

		if (checkAnswer())
        {
            //呼叫ajax myhomework_score把作答資料寫入資料庫
            $.ajax({
                type: "POST",
                url: "myhomework_score.asp",
                data: temp_data,
                beforeSend: function ()
                {
                    $.blockUI({ message: "<h1>Please Wait...</h1>" });
                },
                error: function (xhr) { alert("錯誤" + xhr.responseText) },
                success: function (msg)
                {
					
                    viewScore($("#session_sn").val(), $("#client_sn").val())
                    //alert(msg);
                },
                complete: function ()
                {
                    //呼叫顯示的ajax				
                    $.unblockUI();
                }
            }); 
        }
    });
})
</script>
  <div class="main">
    <div class="main_con">
        <div class="main_membox">
<!--上版位課程start-->
	<!--#include virtual="/lib/include/html/include_learning.asp"--> 
<!--上版位課程end-->	
<% 
on error resume next

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
	    arrsessionTime = Split(getSessionWholeTime(session_sn,3,0,CONST_TUTORABC_RW_CONN), " ")
		sessionTime = arrsessionTime(0) &" - " &arrsessionTime(1)
'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  end '''''''''''''''
	Dim consultant : consultant = g_arr_result(7,0) '上課的老師
	Dim level : level = g_arr_result(6,0) '該學生的程度
	'Dim client_level : client_level = rs("attend_level")'該學生的程度
	Dim description : description = g_arr_result(9,0) '描述
	Dim course : course = g_arr_result(5,0) '課名序號

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
		<form name='form' method="POST" action="myhomework_score.asp">
        <!--quest start-->
        <div class="wlist3"><!--這裡頭的radio button應該是灰色不可點選的樣子-->
<%
	'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  Start '''''''''''''''
			checkcalsstype = isShortSession(session_sn, "", 0, CONST_TUTORABC_RW_CONN) '判斷是否為短課程，若不是，回傳false
		
			dim intQnum:intQnum=5
			if false <> checkcalsstype then '短課程
				intQnum=3		
			else
				intQnum=5
			end if
	'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  End '''''''''''''''
%>		
		<div class="title">Your score is<span class="score"></span></div>
		
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
	sql="SELECT Top "& intQnum &" * FROM lesson_question WHERE (Course = '"&course&"') and (Question <> '')" '短課程:3題 ,一般課程:5題
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
			For i=0 to Ubound(questions_arr)
		%>
		<P >
        <img src="/images/images_cn/wrong.gif" width="20" height="20"  style="display:none"/>
		<!--題目-->
		<div ><%=(i+1)&"."&questions_arr(i)%></div>
		<br>
			<!--隱藏的input 把正確解答放進來-->
			<input type="hidden" name="right_answer_<%=i%>" id="right_answer_<%=i%>" value="<%=questions_answer(i)%>" />
			<%
				For j=0 to Ubound(choices_title) '選項最多為三個 
			%>
				<input type="radio" name="answer_<%=i%>" id="answer_<%=i%>" value="<%=j+1%>" />
				<span ><!--選項標題--><%=choices_title(j)%>
				</span><!--選項-->
				<font ><%=choice_arr(j,i)%></font>
		    <%
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
			  <div class="p-bottom15">
			  <!--左邊s-->
			  <div class="lt">
				<%			  
					Dim q_def_arr_result '用來接收連連看定義的資料陣列
					Dim q_def_arr '定義單字陣列
					Dim q_explain_arr '單字解釋陣列
					Dim q_def_title : q_def_title = array("A","B","C","D","E")					
					Dim explain_lan '解釋的語系
					Dim tmp_rnd : tmp_rnd = "" '暫存亂數數字
					Dim tmp_rnd_str '暫存亂數字串
					Dim vocsn '單字的sn 字串串起來
					Dim explain_right_answer : explain_right_answer = array("","","","","")'釋義的正確答案陣列
						if 3 = intQnum then '如果是短課程 定義不同長度陣列
							q_def_title = array("A","B","C")
							explain_right_answer = array("","","")
						end if						
					'選出連連看的題目單字 只需要"五"筆？ valid必須為1且中文英文的釋義都需存在
					'sql="SELECT TOP 5 sn, vocabulary AS vocabulary_keyword, definition FROM material_vocabulary WHERE ((course = "&course&" AND c_dictionary_state = '1' AND valid = '1') OR (course = "&course&" AND e_dictionary_state = '1' AND valid = '1')) AND is_checked = '1'"
					sql = "SELECT TOP  " & intQnum 
					sql = sql & "    sn "
					sql = sql & "    , vocabulary AS vocabulary_keyword "
					sql = sql & "    , definition "
					sql = sql & "FROM material_vocabulary  "
					sql = sql & "WHERE (1 = 1 ) "
					sql = sql & "    AND (course = " & course & ") "
					sql = sql & "    AND (valid = 1) "
					sql = sql & "    AND (c_dictionary_state = 1 OR e_dictionary_state = 1 OR is_checked = 1) "
					'sql = sql & "    AND (definition > '') "
					q_def_arr_result = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)
					Redim q_explain_arr(Ubound(q_def_arr_result,2))
					Redim q_def_arr(Ubound(q_def_arr_result,2))
					'=========================亂數排列釋義==========start===
					Dim rnd_explain '亂數排列後的釋義陣列
					Dim rnd_explain_index : rnd_explain_index = array("0","1","2","3","4") '亂數的index
						if 3 = intQnum then	'如果是短課程 定義不同陣列
							rnd_explain_index = array("0","2","1")
						end if	
					Dim rnd_explain_order : rnd_explain_order = "" '亂數的順序 用來紀錄題目的順序
					Redim rnd_explain(Ubound(rnd_explain_index)) 
					'開始亂數
						if 3 = intQnum then '如果是短課程，取亂數
							inthigh=3
						else
							inthigh=4
						end if					
					For i = 0 to Ubound(rnd_explain_index)
						'Fix((high - low + 1) * Rnd) + low  公式：亂數產生介於high ~low之間的數
						'Randomize '亂數子
						tmp_rnd= Fix((inthigh - 0 + 1) * Rnd) + 0 
						'將第一個值傳到暫存值中
						tmp_rnd_str = rnd_explain_index(tmp_rnd)
						'將亂數的陣列與第一個值調換
						rnd_explain_index(tmp_rnd) = rnd_explain_index(0)
						'將第一個值換到亂數陣列中
						rnd_explain_index(0) = tmp_rnd_str
						'將亂數的順序串成字串
						rnd_explain_order = rnd_explain_order&rnd_explain_index(i)&","
					Next
						rnd_explain_order = left(rnd_explain_order, len(rnd_explain_order) - 1)
					'=========================亂數排列釋義==========end===
					'將值塞入單字與單字解釋陣列中
					For i=0 to Ubound(q_def_arr_result,2)
						'將單字塞入q_def_arr陣列中
						q_def_arr(i) = q_def_arr_result(1,i)
						'將單字解釋塞入q_explain_arr陣列中
						q_explain_arr(i) = q_def_arr_result(2,i)
						'將單字的sn串起來 vocsn 作為後端接收的值
						vocsn = vocsn&q_def_arr_result(0,i)&","
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
					'將voc_sn最右邊的、去除
					vocsn = left(vocsn, len(vocsn) - 1)
				%>
				<%
					'----------------------下半部連連看---start---------------------------------------------------
					For i=0 to Ubound(q_def_arr_result,2)	
					if 3 = intQnum then istart=4 else istart = 6 
				%>
					<!--1s-->			
					<div class="subj "><img src="/images/images_cn/yes.gif" width="20" height="20" style="display:none" /><!--序號--><%=i+istart&"."%>
					<span class="subj-font "><!--單字--><%=q_def_arr(i)%></span>
					</div>
					<div class="subj-as ">
					<%
						For j=0 to Ubound(q_def_title)
							'================重要！紀錄下正確的釋義答案============start=========
							If q_explain_arr(rnd_explain_index(j)) = q_explain_arr(i) then
								explain_right_answer(i) = q_def_title(j)
							end if
							'================重要！紀錄下正確的釋義答案============end=========
					%>
						<input type="radio" name="answer_<%=i+5%>" id="answer_<%=i+5%>" value="<%=q_def_title(j)%>" /><!--選項--><%=q_def_title(j)%>
						
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
			<div class="rt">
			  <!--1s-->
				<%
					For i=0 to Ubound(q_explain_arr)
				%>
				  <div class="cho">
				  
				  <h1><!--解釋標題--><%=q_def_title(i)%></h1>
				  <h2><!--單字解釋 亂數排列--><%=q_explain_arr(rnd_explain_index(i))%></h2>
				 
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
		<!--需要帶給後端的hidden值 start--> 
		<input type="hidden" id="vocsn" name="vocsn" value="<%=vocsn%>" />
		<input type="hidden" id="session_sn" name="session_sn" value="<%=session_sn%>" />
		<input type="hidden" id="client_sn" name="client_sn" value="<%=client_sn%>" />
		<input type="hidden" id="rnd_explain_order" name="rnd_explain_order" value="<%=rnd_explain_order%>" />
		<!--需要帶給後端的hidden值 end-->
        <div class="bline">
			<input type="button" name="button2" id="button2" value="+ 回上页" class="btn_1" onclick="window.location = 'learning_review.asp'" />
			<input type="button" class="btn_1 m-left5" value="+ 提交" id="send_homework" name="button2">
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
	