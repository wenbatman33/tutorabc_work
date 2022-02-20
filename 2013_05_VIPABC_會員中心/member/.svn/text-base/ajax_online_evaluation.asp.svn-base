<!--#include virtual="/lib/include/global.inc"-->
<script type="text/javascript">
//Down();
</script>
<%

function getQuestionString(ByVal p_str_question) 
	
	Dim str_qustion : str_qustion = ""
	
	str_qustion = replace(p_str_question,"<","@")
	str_qustion = replace(str_qustion,">","*")
	str_qustion = replace(str_qustion,"@","<u>")
	str_qustion = replace(str_qustion,"*","</u>")
	str_qustion = replace(str_qustion,"[","<font color=red><b>")
	str_qustion = replace(str_qustion,"]"," </b></font>")
	str_qustion = replace(str_qustion,"B:","<br>B:")
	str_qustion = replace(str_qustion,"C:","<br>C:")
	
	getQuestionString = str_qustion
end function


'Response.write Session("mh_usersn")
'**************************************************
'參數設定
Dim Qorder	'Qorder : 考題類型順序
Dim Qtype : Qtype = Trim(Request("Qtype")) 'Qtype : 考題項次
Dim QMN : QMN = 7 'QMN : 每階段考題最大值 (選6會考7題)
Dim QN : QN = Trim(Request("QN"))'QN : 考題編號
Dim QL : QL = Trim(Request("QL"))		'QL : 考題LEVEL
Dim shotclock : shotclock = "00:45" 'shotclock : 計數器時間
Dim rclick : rclick = "true" 'rclick : 是否開放右鍵 true為關閉右鍵功能
Dim cqstr : cqstr = Trim(Request("cqstr"))
Dim Qsn
Dim QA
Dim question 		'題目
Dim question_type	'題目上方說明
Dim Qfile			'題目附檔
Dim q
Dim Choices1 		'題目選項1
Dim Choices2 		'題目選項2
Dim Choices3 		'題目選項3
Dim Choices4 		'題目選項4
Dim AR
Dim str_clientQLevel_1
Dim str_clientQLevel_2
Dim Score_G : Score_G=Trim(Request("Score_G"))
Dim Score_V : Score_V=Trim(Request("Score_V"))
Dim Score_P : Score_P=Trim(Request("Score_P"))
Dim Score_S : Score_S=Trim(Request("Score_S"))
Dim Score_L : Score_L=Trim(Request("Score_L"))
Dim Score_R : Score_R=Trim(Request("Score_R"))
'---- hidden -----
Dim agree : agree = Trim(Request("agree"))
Dim Testname
Dim total_question 
Dim qlstr : qlstr = ""
Dim LT : LT = Trim(Request("LT"))
Dim i 

'塞資料的相關設定值
Dim str_sql
Dim var_arr 								'傳excuteSqlStatementRead的陣列
Dim arr_result							'接回來的陣列值

Qorder=array(0,2,5,1,4,3,6)


if (Qtype = "") then Qtype=1 end if 
if (QL = "" ) then QL = 3 end if
if (QN = "" ) then QN = 1 end if  
if (Score_G = "") then Score_G=0 end if
if (Score_V = "") then Score_V=0 end if
if (Score_P = "") then Score_P=0 end if
if (Score_S = "") then Score_S=0 end if
if (Score_L = "") then Score_L=0 end if
if (Score_R = "") then Score_R=0 end if
'------------- 設定題目時間--------------- start ------------
shotclock="00:15"

if Qtype<>4 and Qtype<>5 and QL=1 then shotclock="00:15"
if Qtype<>4 and Qtype<>5 and QL=2 then shotclock="00:20"
if Qtype<>4 and Qtype<>5 and QL=3 then shotclock="00:25"
if Qtype<>4 and Qtype<>5 and QL=4 then shotclock="00:30"
if Qtype<>4 and Qtype<>5 and QL=5 then shotclock="00:35"
if Qtype<>4 and Qtype<>5 and QL=6 then shotclock="00:45"

if Qtype=5 and QL=1 then shotclock="00:25"
if Qtype=5 and QL=2 then shotclock="00:25"
if Qtype=5 and QL=3 then shotclock="00:30"
if Qtype=5 and QL=4 then shotclock="00:35"
if Qtype=5 and QL=5 then shotclock="00:45"
if Qtype=5 and QL=6 then shotclock="00:45"

if Qtype=4 and QL=1 then shotclock="00:20"
if Qtype=4 and QL=2 then shotclock="00:25"
if Qtype=4 and QL=3 then shotclock="00:25"
if Qtype=4 and QL=4 then shotclock="00:30"
if Qtype=4 and QL=5 then shotclock="00:35"
if Qtype=4 and QL=6 then shotclock="00:40"
'------------- 設定題目時間--------------- end ------------
'------------ 開始出題 ------------ start -------------
	if (Qtype="") then Qtype=1 end if 
	
	if Qorder(Qtype)=1 then Testname="process4" '"Grammar - 文 法"
	if Qorder(Qtype)=2 then Testname="process2" '"Vocabulary - 字 彙"
	if Qorder(Qtype)=3 then Testname="process6" '"Pronunciation - 發 音"
	if Qorder(Qtype)=4 then Testname="process5" '"Listening - 聽 力"
	if Qorder(Qtype)=5 then Testname="process3" '"Reading - 閱 讀"
	'if Qorder(Qtype)=6 then Testname="process2" '"Speaking - 會 話"
	
	
	
	if (QL="") then QL=3 end if 
   	
	
	if (QL <= 6 and QL >= 1) then 
		'1--> 1,2 2-->3,4...6-->11,12
		str_clientQLevel_2 = QL * 2 
		str_clientQLevel_1 = str_clientQLevel_2 - 1
		qlstr=" and (client_adaptive_question.QLevel LIKE '%-"&str_clientQLevel_1&"-%' OR client_adaptive_question.QLevel LIKE '%-"&str_clientQLevel_2&"-%')"
		
	end if 
	
	str_sql = " SELECT  sn,answer,question,cntype,secondary_filename,Choices1,Choices2,Choices3,Choices4  "
	str_sql = str_sql & " FROM client_adaptive_question "
	str_sql = str_sql & " LEFT JOIN cfg_adaptive_type ON client_adaptive_question.QType = cfg_adaptive_type.type_no "
	str_sql = str_sql & " WHERE (client_adaptive_question.Valid = @valid)  AND (Secondary_Filename <> @Secondary_Filename ) "
	str_sql = str_sql & " AND (LEFT(client_adaptive_question.QType, 1) = @QType)"&qlstr&""
	
	
	if (QN = "") then QN=1 end if 
	if (LT = "") then LT="tw" end if 
	
	if (cqstr <> "") then 
		if (left(cqstr,1)=",") then cqstr=right(cqstr,len(cqstr)-1) end if 
		if (right(cqstr,1)=",") then cqstr=left(cqstr,len(cqstr)-1) end if 

		str_sql = str_sql & " and sn not in ("&cqstr&") " 

	end if
	
	'response.write "g_str_sql........."&str_sql&"<br>"
	
	var_arr = Array(1,"mp3",Qorder(Qtype))
	arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)

	if (isSelectQuerySuccess(arr_result)) then
    
		'--------- 塞適性測驗 ------------ start ------------
   		if (Ubound(arr_result) > 0) then 
			'有資料
			total_question = Ubound(arr_result) + 1
			if (QN = 1 ) then
				cqstr = ""
			else
				cqstr = Trim(Request("cqstr"))
			end if
			
			Randomize
			Qsn = Int((total_question * Rnd) + 1)
			
			'隨機選題
			if ( Qsn > 1 ) then
				Qsn = Qsn -1 
			end if 
			
			
			cqstr = arr_result(0,Qsn)&","&cqstr '排除已經做過的題目
			
			
			
			QA = arr_result(1,Qsn)  		'答案
			
			question = arr_result(2,Qsn)
			question_type = arr_result(3,Qsn)
			Qfile = arr_result(4,Qsn)
			Choices1 = arr_result(5,Qsn)
			Choices2 = arr_result(6,Qsn)
			Choices3 = arr_result(7,Qsn)
			Choices4 = arr_result(8,Qsn)
		else
			'無資料 (如果已經沒有題目時)
			'response.write "Stop"
			%>
			<form name="fm_evaluation_exe" method="post" action="online_evaluation_exe.asp">
                <input type="hidden" id="cqstr" name="cqstr" value="<%=cqstr%>">
                <input type="hidden" id="Score_G" name="Score_G" value="<%=Score_G%>">  <!--'G文法-->
                <input type="hidden" id="Score_V" name="Score_V" value="<%=Score_V%>">  <!--'V字彙-->
                <input type="hidden" id="Score_P" name="Score_P" value="<%=Score_P%>">  <!--'P發音-->
                <input type="hidden" id="Score_S" name="Score_S" value="<%=Score_S%>">  <!--'S會話-->
                <input type="hidden" id="Score_L" name="Score_L" value="<%=Score_L%>">  <!--'L聽力-->
                <input type="hidden" id="Score_R" name="Score_R" value="<%=Score_R%>">  <!--'R閱讀-->
 			<%
			Response.Write "<script language=""javascript"" type=""text/javascript"">"
			Response.Write "document.fm_evaluation_exe.submit();"
			Response.Write "</script>"
			
			%>
			</form>
			<%
		end if 
	end if 
%>
       <!--內容start-->
        <!--英語能力測試內容start-->
        <div class="main_classcontbox ">
        <!--英語能力測試過程列start-->
        <div class="main_classtest_processbar">
        <%=getWord("you_play_now")%><br/>
        <ul class="<%=Testname%>">
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
        <!--測試時間計時列start-->
        <div class="test_timer">
        <input type="hidden" name="beg2" id="beg2" size="6" value="<%=shotclock%>">
        <h3><%=getWord("surplus")%> <span class="color_1 bold">
		<input type="text" id="disp2" name="disp2" size="6" style="font-family: Times New Roman; font-size: 10 pt" /></span> <%=getWord("sec")%></h3>
        </div>
        <!--測試時間計時列end-->
        <div class="clear"></div>
        <!--測試題內容start-->
        <div class="quizbox">
         
        <%
		'------------------------ 呈現題目 --------------------  start -------------------
		if (question <> "" ) then
		%>
        	<h2><span class="color_1">Q.</span> <%=question_type%></h2>
        <%
			if ( instr(question,"__") > 0 ) then
		%>
        		<p><%=question%></p>
		<%
			else
			'	response.write "bbb"&rs("sn")&"<br>"
				q = getQuestionString(question)
				if ( q<>"jpg" ) and ( q <> "mp3" ) then			
		%>
        			<p><%=q%></p>
		<%
				end if
				
			end if
		
		if  ( Qfile<>"0" )  then
			if right(Qfile,3)="jpg" then
			%>
				<img src="../../audio/q_files/images/<%=Qfile%>" width="120" height="120" />
			<%
			else
			%>
                 <div id="flash_ability_playbtn" style="float:left;"><%=getWord("PLEASE_USED_FLASH_PLAYER_10_AFTER")%></div>
				<script type="text/javascript">
				    var fo = new FlashObject("/flash/ability_playbtn.swf?theFile=http://<%=CONST_VIPABC_WEBHOST_NAME%>/audio/q_files/images/<%=Qfile%>", "playbtn", "40", "40", "7", "#000000"); 
                    fo.addParam("wmode", "transparent");
                    fo.addParam("allowFullscreen","true");
                    fo.write("flash_ability_playbtn");
                </script>    
			<%
	
			end if
	end if
	
		%>
        
        <%end if
		'------------------------ 呈現題目 --------------------  end -------------------
		 %>
        <ul>
        <%
		'-------------------- 第一個選項 ---------------- start --------------
		if Choices1<>"" then
			q = getQuestionString(Choices1)
		%>
        <li><input type="radio" name="Answer" value="<%=Choices1%>" />(a) <%=q%></li>
        <%end if
		'-------------------- 第一個選項  --------------end -------------------
		'-------------------- 第二個選項  --------------start -------------------
		if Choices2<>"" then
			q = getQuestionString(Choices2)
		 %>
        <li><input type="radio" name="Answer"  value="<%=Choices2%>" />(b) <%=q%></li>
        <%end if
		'-------------------- 第二個選項  --------------end -------------------
		'-------------------- 第三個選項  --------------start -------------------
		if Choices3<>"" then
			q = getQuestionString(Choices3)
		 %>
        	<li><input type="radio" name="Answer"  value="<%=Choices3%>" />(c) <%=q%></li>
        <%end if
		'-------------------- 第三個選項  --------------end -------------------
		'-------------------- 第四個選項  --------------start -------------------
		if Choices4<>"" then
			q = getQuestionString(Choices4)
		 %>
         	<li><input type="radio" name="Answer"  value="<%=Choices4%>" />(d) <%=q%></li>
        <%end if
		'-------------------- 第四個選項  --------------end -------------------
		%>
        </ul>
        
        <div class="testgo"><input type="button" onClick="loadOnlineEvaluation();" name="send_answer" id="send_answer" value="Next" class="btn_2" /></div>
        </div>
        <!--測試題內容end-->
        </div>
        <!--英語能力測試內容end-->
        <input type="hidden" id="agree" name="agree" value="<%=agree%>" />
        <input type="hidden" id="Qtype"  name="Qtype" value="<%=Qtype%>" />
        <input type="hidden" id="QL" name="QL" value="<%=QL%>" />
        <input type="hidden" id="QN" name="QN" value="<%=QN%>" />
        <input type="hidden" id="LT" name="LT" value="<%=LT%>" />
        <input type="hidden" id="QA" name="QA" value="<%=QA%>" />
        <input type="hidden" id="cqstr" name="cqstr" value="<%=cqstr%>" />
        <input type="hidden" id="Score_G" name="Score_G" value="<%=Score_G%>" />
        <input type="hidden" id="Score_V" name="Score_V" value="<%=Score_V%>" />
        <input type="hidden" id="Score_P" name="Score_P" value="<%=Score_P%>" />
        <input type="hidden" id="Score_S" name="Score_S" value="<%=Score_S%>" />
        <input type="hidden" id="Score_L" name="Score_L" value="<%=Score_L%>" />
        <input type="hidden" id="Score_R" name="Score_R" value="<%=Score_R%>" />