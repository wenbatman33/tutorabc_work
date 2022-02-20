<!--#include virtual="/lib/include/global.inc"-->
<% 
dim answer(10)
	For i = 0 to 9 
		answer(i)=request("answer_"&i)		
			if "undefined" = answer(i) then
				answer(i) = 0
			end if
	Next

	'=================尚未作答過 計算分數寫入資料==============start=====================
	Dim homework_answer : homework_answer ="" '儲存答案的字串
	Dim homework_right_answer : homework_right_answer ="" '儲存答案的字串
	'將所有使用者的"答案"與"正確答案"串成字串存入資料庫欄位ans
	For i = 0 to 9 
		'最後一個不帶、
		If i<9 then
			homework_answer = homework_answer&answer(i)&","
			homework_right_answer = homework_right_answer&request("right_answer_"&i)&","
		Else
			homework_answer = homework_answer&answer(i)
			homework_right_answer = homework_right_answer&request("right_answer_"&i)
		End If
	Next	
	Dim session_sn : session_sn = request("session_sn")
	Dim client_sn
	Dim arr_homework_answer '使用者填答案的陣列
	Dim arr_homework_right_answer '答案正確解答的陣列
	Dim point : point = 0 '作答總分
	Dim i
	'如果沒找到client_sn or session_sn
	if client_sn="" then
		client_sn = session("client_sn")
	else
		client_sn = request("client_sn")
	end if
	'將字串分為陣列
	arr_homework_answer = split(homework_answer,",")
	arr_homework_right_answer = split(homework_right_answer,",")
	'計算分數
	For i = 0 to ubound(arr_homework_answer)
		If arr_homework_answer(i) = arr_homework_right_answer(i) then
			point = point + 10
		End If
	Next
	
'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  Start '''''''''''''''
	checkcalsstype = isShortSession(session_sn, "", 0, CONST_TUTORABC_RW_CONN) '判斷是否為短課程，若不是，回傳false
	if false <> checkcalsstype then
		point = sCint(point / 0.6)
	end if	
'''''''''''''''2013/04/25 Jim add 短課程所需要的格式  End '''''''''''''''
	
	'=====================將獲取資料寫入資料庫======start===============
	Dim sql, g_arr_result
	sql = "SELECT * FROM dbo.client_homework WHERE (Client_Sn = '"&client_sn&"') AND (Session_Sn = '"&session_sn&"')"
	'response.write "sql="&sql
	'response.end
	g_arr_result = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)
	If (isSelectQuerySuccess(g_arr_result)) then
		If (Ubound(g_arr_result)>=0) then
			response.write "重複作業！"
		Else
			sql= "INSERT INTO client_homework (Client_Sn, Session_Sn,Inp_Date,voc_sn,anser,Score,Status,Valid,Q1) VALUES ('"&client_sn&"','"&session_sn&"','"&getFormatDateTime(Now(), 1)&"','"&request("vocsn")&"','"&homework_answer&"','"&point&"','0','1','"&request("rnd_explain_order")&"')"
		    '將資料寫入
	        Dim g_effected_row '影響的比數
	        g_effected_row = excuteSqlStatementWriteQuick(sql, CONST_VIPABC_RW_CONN)
	        if (g_effected_row >= 0) then
		        Response.Write("影響筆數 " & g_effected_row & "<br>")
	        else
		        Response.Write("影響筆數錯誤訊息" & g_str_run_time_err_msg & "<br>"&g_str_sql_statement_for_debug)
	        end if
        end If
	else
		response.write "資料庫存取錯誤"
	End If
	'=====================將獲取資料寫入資料庫======End=================
	'=================尚未作答過 計算分數寫入資料==============end=====================
%>	