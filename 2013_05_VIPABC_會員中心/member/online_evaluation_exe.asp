<!--#include virtual="/lib/include/global.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>無標題文件</title>
</head>
<body>
    <%

Dim QMN : QMN = 7
Dim maxscore : maxscore = 12+6*(QMN-3)
Dim minscore : minscore = 3
Dim total_score

Dim cqstr : cqstr = Trim(Request("cqstr"))
Dim str_client_sql : str_client_sql = ""

Dim int_g_score 
Dim int_p_score 
Dim int_r_score 
Dim int_l_score 
Dim int_v_score 
Dim int_score
Dim level  		'適性測驗的等級
Dim str_olevel	'cleint.basic.olevel
Dim str_nlevel	'cleint.basic.nlevel
Dim str_test_sn 
Dim i 

Dim arr_Client_Level		'Client_Level(整批塞入)
Dim arr_Total_score		'Total_score(整批塞入)

'塞資料的相關設定值						
Dim var_arr 								'傳excuteSqlStatementRead的陣列
Dim arr_result							'接回來的陣列值
Dim str_tablecol							'update col
Dim str_tablevalue
'beginTranscation
Dim transaction_opt
Dim int_add_sql_state 
Dim int_result_id

'------- G 文法的分數 ---------- start ----------
Dim GC : GC = Int(Request("Score_G"))
total_score = Int(GC)

if 3/QMN <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 then int_g_score=1
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN)/10))*2 then int_g_score=2
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*2 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 then int_g_score=3
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 then int_g_score=4
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 then int_g_score=5
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 then int_g_score=6
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 then int_g_score=7
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 then int_g_score=8
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 <= GC/QMN and GC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 then int_g_score=9
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 <= GC/QMN and GC/QMN <= (3/QMN)+(((36/QMN)-(3/QMN))/10)*10 then int_g_score=10


'------- V 字彙的分數 ---------- start ---------
Dim VC : VC = Int(Request("Score_V"))
total_score = total_score + Int(VC)

if 3/QMN <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 then int_v_score=1
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN)/10))*2 then int_v_score=2
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*2 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 then int_v_score=3
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 then int_v_score=4
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 then int_v_score=5
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 then int_v_score=6
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 then int_v_score=7
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 then int_v_score=8
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 <= VC/QMN and VC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 then int_v_score=9
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 <= VC/QMN and VC/QMN <= (3/QMN)+(((36/QMN)-(3/QMN))/10)*10 then int_v_score=10


'------- P 發音的分數 ---------- start ---------
Dim PC : PC = Int(Request("Score_P"))
total_score = total_score + Int(PC)

if 3/QMN <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*1 then int_p_score=1
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*1 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN)/10))*2 then int_p_score=2
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*2 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*3 then int_p_score=3
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*3 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*4 then int_p_score=4
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*4 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*5 then int_p_score=5
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*5 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*6 then int_p_score=6
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*6 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*7 then int_p_score=7
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*7 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*8 then int_p_score=8
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*8 <= PC/QMN and PC/QMN < (3/QMN)+(((27/QMN)-(3/QMN))/10)*9 then int_p_score=9
if (3/QMN)+(((27/QMN)-(3/QMN))/10)*9 <= PC/QMN and PC/QMN <= (3/QMN)+(((27/QMN)-(3/QMN))/10)*10 then int_p_score=10


'------- L 聽力的分數 ---------- start ---------
Dim LC : LC = Int(Request("Score_L"))
total_score = total_score + Int(LC)

if 3/QMN <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 then int_l_score=1
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN)/10))*2 then int_l_score=2
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*2 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 then int_l_score=3
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 then int_l_score=4
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 then int_l_score=5
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 then int_l_score=6
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 then int_l_score=7
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 then int_l_score=8
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 <= LC/QMN and LC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 then int_l_score=9
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 <= LC/QMN and LC/QMN <= (3/QMN)+(((36/QMN)-(3/QMN))/10)*10 then int_l_score=10


'------- R 閱讀的分數 ---------- start ---------
Dim RC : RC = Int(Request("Score_R"))
total_score = total_score + Int(RC)

if 3/QMN <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 then int_r_score=1
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*1 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN)/10))*2 then int_r_score=2
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*2 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 then int_r_score=3
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*3 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 then int_r_score=4
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*4 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 then int_r_score=5
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*5 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 then int_r_score=6
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*6 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 then int_r_score=7
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*7 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 then int_r_score=8
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*8 <= RC/QMN and RC/QMN < (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 then int_r_score=9
if (3/QMN)+(((36/QMN)-(3/QMN))/10)*9 <= RC/QMN and RC/QMN <= (3/QMN)+(((36/QMN)-(3/QMN))/10)*10 then int_r_score=10

'------- S 會話的分數 ---------- start ---------
Dim SC : SC= total_score/5 

if total_score>=15 and total_score<=22.8 then level=1
if total_score>22.8 and total_score<=30.6 then level=2
if total_score>30.6 and total_score<=46.2 then level=3
if total_score>46.2 and total_score<=61.8 then level=4
if total_score>61.8 and total_score<=77.4 then level=5
if total_score>77.4 and total_score<=93 then level=6
if total_score>93 and total_score<=108.6 then level=7
if total_score>108.6 and total_score<=124.2 then level=8
if total_score>124.2 and total_score<=139.8 then level=9
if total_score>139.8 and total_score<=155.4 then level=10
if total_score>155.4 and total_score<=163.2 then level=11
if total_score>163.2 and total_score<=171 then level=12

	int_score=INT((int(int_g_score)+int(int_p_score)+int(int_r_score)+int(int_l_score)+int(int_v_score))/5)

'20120320 阿捨新增 VIPABCJR客戶判斷 bolVIPABCJR 
%>
<!--#include virtual="/program/member/getJRClient.asp"-->
<%
'---------- inster to lesson_question_temp ------------------------- start  ----------------- 
'1：文法 / 2：字彙 / 3：發音 / 4：會話/ 5：閱讀 / 6：聽力
	arr_Client_Level = Array(int_g_score,int_v_score,int_p_score,int_score,int_r_score,int_l_score)
	arr_Total_score = Array(GC,VC,PC,SC,RC,LC)

	arr_result = excuteSqlStatementRead("select max(test_sn)as max_testsn from lesson_question_temp" , "", CONST_VIPABC_RW_CONN)

if (isSelectQuerySuccess(arr_result)) then
	if (Ubound(arr_result) >= 0) then 
		'有資料，回傳一維陣列
		str_test_sn = arr_result(0,0) + 1
		
		str_tablecol = "Client_Sn,Evaluation_Type,Client_Level,Total_score,question_str,test_sn"
		str_tablevalue = "@Client_Sn,@Evaluation_Type,@Client_Level,@Total_score,@question_str,@test_sn"
		
		Set transaction_opt = beginTranscation(CONST_VIPABC_RW_CONN)
			For i = 0 To 5
				var_arr = Array(g_var_client_sn,i+1,arr_Client_Level(i),arr_Total_score(i),cqstr,str_test_sn)
				int_add_sql_state = addTranscationSQL("Insert into  lesson_question_temp("&str_tablecol&") values("&str_tablevalue&")",var_arr,transaction_opt)
				if(int_add_sql_state=-1) then
					'可看單筆是否成功
					'Response.write("Insert Failed")
				end if
			Next
		int_result_id = endTranscation(transaction_opt)
		
		if(int_result_id<0) then
			'失敗
			'Response.write g_str_run_time_err_msg
		else
			'成功
			'Response.write "Transcation Insert Success" & "<br>"
		end if
		Set transaction_opt =Nothing
	else
		'錯誤訊息
	end if 
end if

'---------- inster to lesson_question_temp ------------------------- end -----------------

'-------將適性測驗後的level記到client_basic的olevel及nlevel --------- start ------------
	str_client_sql = "select sn,olevel,nlevel from client_basic where sn = @sn"
	var_arr = Array(g_var_client_sn)
	arr_result = excuteSqlStatementRead(str_client_sql, var_arr, CONST_VIPABC_RW_CONN)

if (isSelectQuerySuccess(arr_result)) then 
	
	if (Ubound(arr_result) > 0) then 	
		'------ update cleint_basic --------- start -----------
		'有資料
		'2006/9/11 Sam補充:需要在客戶沒有資料的情況才可以新增.要不然會出現客戶新完成的適性成績會記錄到客戶目前上課的紀錄內!

		'--- olevel -----
		if (arr_result(1,0) <> "") then 
			 str_olevel = arr_result(1,0)
		else			
			str_olevel = level
		end if 
		'--- nlevel -----
		if (arr_result(2,0) <> "") then 
			str_nlevel = arr_result(2,0)
		else
			str_nlevel = level 
		end if 
		
		'------ update cleint_basic --------- start -----------
		str_tablecol = "olevel=@olevel,nlevel=@nlevel"
		var_arr = Array(str_olevel,str_nlevel,g_var_client_sn)
		arr_result = excuteSqlStatementWrite("Update client_basic set "&str_tablecol&" where sn=@sn" , var_arr, CONST_VIPABC_RW_CONN)

		if (arr_result > 0) then 
			'成功
		else
			'失敗
		end if 
		'------ update cleint_basic --------- end -----------

        '20120323 阿捨新增 VIPABCJR客戶判斷 bolVIPABCJR
        if ( true = bolVIPABC ) then
		    var_arr = Array(str_nlevel, g_var_client_sn)
		    arr_result = excuteSqlStatementWrite(" UPDATE member.Profile SET ProfileLearnLevel = @ProfileLearnLevel WHERE client_sn = @client_sn " , var_arr, CONST_VIPABC_RW_CONN)

		    if (arr_result > 0) then 
			    '成功
		    else
			    '失敗
		    end if 
        end if
	else
		'無資料
	end if 
end if 
'-------將適性測驗後的level記到client_basic的olevel及nlevel ---------end ----------

Call alertGo("", "/program/member/online_evaluation_finsh.asp", CONST_NOT_ALERT_THEN_REDIRECT )
    %>
</body>
</html>