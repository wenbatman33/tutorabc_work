<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type="text/javascript">
//倒數計時器 start
	var up,down;
	var min1,sec1;
	var cmin1,csec1,cmin2,csec2;
	function Minutes(data) {
		for(var i=0;i<data.length;i++) 
		if(data.substring(i,i+1)==":") 
		break;  
		return(data.substring(0,i)); 
	}
	function Seconds(data) {        
		for(var i=0;i<data.length;i++) 
		if(data.substring(i,i+1)==":") 
		break;  
		return(data.substring(i+1,data.length)); 
	}
	function Display(min,sec) {    
		var disp;       
		if(min<=9) disp=" 0";   
		else disp=" ";  
		disp+=min+":";  
		if(sec<=9) disp+="0"+sec;       
		else disp+=sec; 
		return(disp); 
	}
	function Up() { 
		cmin1=0;        
		csec1=0;        
		min1=0+Minutes(document.timers.beg1.value); 
		sec1=0+Seconds(document.timers.beg1.value); 
		UpRepeat(); 
	}
	function UpRepeat() {   
		csec1++;        
		if(csec1==60) {
			csec1=0; 
			cmin1++; 
		}
		document.timers.disp1.value = Display(cmin1,csec1);   
		if((cmin1==min1)&&(csec1==sec1)) 
			alert("Stopwatch Stopped");     
		else up=setTimeout("UpRepeat()",1000); 
	}
	function Down() {       
		cmin2=1*Minutes($("#beg2").val());        
		csec2=0+Seconds($("#beg2").val());        
		DownRepeat(); 
	}
	function DownRepeat() { 
		csec2--;        
		if(csec2==-1) { 
			csec2=59; 
			cmin2--; 
		}
		$("#disp2").val(Display(cmin2,csec2)); 
		if((cmin2==0)&&(csec2==0))
		{
			//alert($("#disp2").val());
			loadOnlineEvaluation();
		}
		else
		{
			down = setTimeout("DownRepeat()",1000); 
		}
	}
	// End -->
	//讀取題目
	var QL ;
	var QN ;
	var QA ;
	var AR ;
	var LT ;
	var cqstr ;
	var Score_G ;
	var Score_V ;
	var Score_P ;
	var Score_S ;
	var Score_L ;
	var Score_R ;
	var agree ;
	var Qtype ;
	var beg2 ; 
	//計算每次考試過後的分數
	function count_point()
	{
		var Qorder = new Array("0","2","5","1","4","3","6");
		var QMN = 7 ;
		LT = $("#LT").val();
		agree = $("#agree").val();
		Qtype = parseInt($("#Qtype").val());
		QL = parseInt($("#QL").val());
		QN = parseInt($("#QN").val());
		QA = $("#QA").val();
		AR = $("input:radio:checked").val();
		beg2 = $("#beg2").val();
		cqstr = $("#cqstr").val();
		Score_G = parseInt($("#Score_G").val());
		Score_V = parseInt($("#Score_V").val());
		Score_P = parseInt($("#Score_P").val());
		Score_S = parseInt($("#Score_S").val());
		Score_L = parseInt($("#Score_L").val());
		Score_R = parseInt($("#Score_R").val());

		if ( Score_G == "")
		{
			Score_G=0 ; 
		}
		if ( Score_V == "")
		{
			Score_V=0 ; 
		}
		if ( Score_P == "")
		{
			Score_P=0 ; 
		}
		if ( Score_R == "")
		{
			Score_R=0 ; 
		}
		if ( Score_S == "")
		{
			Score_S=0 ; 
		}
		if ( Score_L == "")
		{
			Score_L=0 ; 
		}
		if( QN <= QMN )
		{
			QN=QN+1	 ;
			if( QA == AR )
			{
				if(Qorder[Qtype] == "3" && QL>4 )
				{
					QL=4 ; 
				}
				if(Qorder[Qtype] == "1")
				{
					Score_G=parseInt(Score_G)+QL ; 
				}
				if(Qorder[Qtype] == "2")
				{
					Score_V=parseInt(Score_V)+QL ; 
				}
				if(Qorder[Qtype] == "3")
				{
					Score_P=parseInt(Score_P)+QL ; 
				}
				if(Qorder[Qtype] == "4")
				{
					Score_L=parseInt(Score_L)+QL ; 
				}
				if(Qorder[Qtype] == "5")
				{
					Score_R=parseInt(Score_R)+QL ; 
				}
				if(Qorder[Qtype] == "6")
				{
					Score_S=parseInt(Score_S)+QL ; 
				}
				if(QL<6)
				{
					QL=QL+1 ; 
				}
				//因Pronunciation只有到 Level B
			}
			else
			{
				if(QL>1)
				{
					QL=QL-1 ; 
				}
				else
				{
					QL=0 ; 
				}
				if(Qorder[Qtype] == "1")
				{
					Score_G=parseInt(Score_G)+QL ; 
				}
				if(Qorder[Qtype] == "2")
				{
					Score_V=parseInt(Score_V)+QL ; 
				}
				if(Qorder[Qtype] == "3")
				{
					Score_P=parseInt(Score_P)+QL ; 
				}
				if(Qorder[Qtype] == "4")
				{
					Score_L=parseInt(Score_L)+QL ; 
				}
				if(Qorder[Qtype] == "5")
				{
					Score_R=parseInt(Score_R)+QL ; 
				}
				if(Qorder[Qtype] == "6")
				{
					Score_S=int(Score_S)+QL ; 
				}
				
			}
		}
		else
		{
			//考完7題後進行下一部分
			if (Qtype<6)
			{
				Qtype=Qtype+1 ; 
			}				
			QN=1 ; 
			QL=3 ; 
		}
		if (QN==8) 
		{
			if(Qtype<6)
			{
				Qtype=Qtype+1 ; 
			}
			QN=1 ; 
			QL=3 ; 
		}
		if(Qorder[Qtype]=="3" && QL>4)
		{
			QL=4 ; 
		}
		
	}
	
	function loadOnlineEvaluation()
	{
		count_point();
		$.ajax({    　　　　　　　  
				type: "POST",
				url: "ajax_online_evaluation.asp",
				cache: false,
				async: false,
				data:{beg2:beg2,agree:agree,Qtype:Qtype,QL:QL,QN:QN,LT:LT,cqstr:cqstr,Score_G:Score_G,Score_V:Score_V,Score_P:Score_P,Score_S:Score_S,Score_L:Score_L,Score_R:Score_R },
				error: function(xhr) {
				alert("錯誤"+xhr.responseText)
				//$("#temp_contant").html(xhr);
				},
				
				beforeSend: function(){
					clearTimeout(down);
				},　　　　　　　　　　　　　　　　　　　
				success: function(result) {
					$("#temp_contant").html(result);
					Down();
				}/*, 
				complete: function(){
					//Down();
				}*/
			});
		//Down();
	}
	 /*
	 function TimeView(sec)
	 {     
         $("#disp2").val(sec);
         setTimeout(function () {
			 if (sec > 0) 
			 {         
				 TimeView(sec-1);         
			 }
		  }, 1000);
         
     } */ 
	$(document).ready(function()
	{	
		$("#send_answer").click(function(){
			loadOnlineEvaluation();
		});
		Down();
	});
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

'response.write "Score_G...."&Score_G&"<br>"
'response.write "Score_V...."&Score_V&"<br>"
'response.write "Score_P...."&Score_P&"<br>"
'response.write "Score_S...."&Score_S&"<br>"
'response.write "Score_L...."&Score_L&"<br>"
'response.write "Score_R...."&Score_R&"<br>"
'response.write "Qtype......"&Qtype&"<br>"

'if(session("mh_usersn") <> "") then
'	mh_usersn = Session("mh_usersn")
'	email = session("email")
'	mh_userfullname = session("mh_userfullname")
'else
'	mh_usersn = request("mh_usersn")
'	email = request("email")
'	mh_userfullname = request("mh_userfullname")
'end if

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

'-------------- 答案比對作業 ---------------- start ---------------
if request("step")=2 then
%>
<form name="mf" method="post" action="online_evaluation.asp">

	<input type="hidden" name="Language" value="<%=Request("Language")%>">
<%
	Qtype=request("Qtype")
	QL=request("QL")
	QN=request("QN")
	QA=request("QA")
	AR=request("Answer")
	cqstr=request("cqstr")

	
	if Score_G="" then Score_G=0
	if Score_V="" then Score_V=0
	if Score_P="" then Score_P=0
	if Score_L="" then Score_L=0
	if Score_R="" then Score_R=0
	if Score_S="" then Score_S=0

	if int(QN)<=int(QMN) then
		QN=QN+1
		if QA=AR then
			if Qorder(Qtype)=3 and QL>4 then QL=4
			if Qorder(Qtype)=1 then Score_G=int(Score_G)+QL
			if Qorder(Qtype)=2 then Score_V=int(Score_V)+QL
			if Qorder(Qtype)=3 then Score_P=int(Score_P)+QL
			if Qorder(Qtype)=4 then Score_L=int(Score_L)+QL
			if Qorder(Qtype)=5 then Score_R=int(Score_R)+QL
			if Qorder(Qtype)=6 then Score_S=int(Score_S)+QL
			if QL<6 then QL=QL+1
'			因Pronunciation只有到 Level B
		else
			if QL>1 then QL=QL-1 else QL=0
			if Qorder(Qtype)=1 then Score_G=int(Score_G)+QL
			if Qorder(Qtype)=2 then Score_V=int(Score_V)+QL
			if Qorder(Qtype)=3 then Score_P=int(Score_P)+QL
			if Qorder(Qtype)=4 then Score_L=int(Score_L)+QL
			if Qorder(Qtype)=5 then Score_R=int(Score_R)+QL
			if Qorder(Qtype)=6 then Score_S=int(Score_S)+QL
		end if
	else
'		考完7題後進行下一部分
		if Qtype<6 then Qtype=Qtype+1
		QN=1
		QL=3
	end if

	if QN=8 then
		if Qtype<6 then Qtype=Qtype+1
		QN=1
		QL=3
	end if
	
	if Qorder(Qtype)=3 and QL>4 then QL=4	
	
	
%>
    <input type="hidden" id="agree" name="agree" value="<%=agree%>">
    <input type="hidden" id="Qtype" name="Qtype" value="<%=Qtype%>">
    <input type="hidden" id="QL" name="QL" value="<%=QL%>">
    <input type="hidden" id="QN" name="QN" value="<%=QN%>">
    <input type="hidden" id="LT" name="LT" value="<%=LT%>">
    <input type="hidden" id="cqstr" name="cqstr" value="<%=cqstr%>">
    <input type="hidden" id="Score_G" name="Score_G" value="<%=Score_G%>">  <!--'G文法-->
    <input type="hidden" id="Score_V" name="Score_V" value="<%=Score_V%>">  <!--'V字彙-->
    <input type="hidden" id="Score_P" name="Score_P" value="<%=Score_P%>">  <!--'P發音-->
    <input type="hidden" id="Score_S" name="Score_S" value="<%=Score_S%>">  <!--'S會話-->
    <input type="hidden" id="Score_L" name="Score_L" value="<%=Score_L%>">  <!--'L聽力-->
    <input type="hidden" id="Score_R" name="Score_R" value="<%=Score_R%>">  <!--'R閱讀-->
<!--    <input type="hidden" name="mh_usersn" value="<%'=mh_usersn%>">
    <input type="hidden" name="email" value="<%'=email%>">
    <input type="hidden" name="mh_userfullname" value="<%'=mh_userfullname%>">-->
    <body onLoad="document.mf.submit()" topmargin="0" leftmargin="0">
</form>

<%
	response.end 
'------------ 答案比對作業 ---------------- end --------------

'------------ 開始出題 ------------ start -------------
else
	
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
	'response.write g_str_run_time_err_msg
	'response.end 
	'response.write str_sql

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
			
			'response.write "QA........."&QA&"<br>"
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
			'response.redirect "online_evaluation_finish.asp?mh_usersn="&mh_usersn&"&email="&email&"&mh_userfullname="&mh_userfullname&"&Score_G="&Score_G&"&Score_V="&Score_V&"&Score_P="&Score_P&"&Score_S="&Score_S&"&Score_L="&Score_L&"&Score_R="&Score_R&"&language="&request("language")
			
			%>
			</form>
			<%
		end if 
	end if 

end if 
%>



<div id="temp_contant">
       <!--內容start-->
        <!--英語能力測試內容start-->
        <form name="timers" method="POST" action="online_evaluation.asp?step=2">
        <div class="main_classcontbox ">
        <div class='page_title_3'><h2 class='page_title_h2'>程度分析</h2></div>
		<div class="box_shadow" style="background-position:-80px 0px;"></div> 
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
        <h3><%=getWord("surplus")%> <span class="color_1 bold"><input type="text" id="disp2" name="disp2" size="6" style="font-family: Times New Roman; font-size: 10 pt"></span> <%=getWord("sec")%></h3>
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
                    var fo = new FlashObject("/flash/ability_playbtn.swf?theFile=/audio/q_files/images/<%=Qfile%>", "playbtn", "40", "40", "7", "#000000"); 
                    fo.addParam("wmode", "transparent");
                    fo.addParam("allowFullscreen","true");
                    fo.write("flash_ability_playbtn");
                </script>    
			<%
'				response.write "<OBJECT classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,47,0' WIDTH='40' HEIGHT='40' id='wimpy_button_1' name='wimpy_button01'>"
'				response.write "<PARAM NAME=movie VALUE='wimpy_button.swf?theFile=http://"&webhost_name&"/consultant/asp/q_files/images/"&rs("Secondary_Filename")&"'>"
'				response.write "<PARAM NAME=quality VALUE=high>"
'				response.write "<PARAM NAME=wmode VALUE=transparent>"
'				response.write "<EMBED src='wimpy_button.swf?theFile=http://"&abc_webhost_name&"/consultant/asp/q_files/images/"&rs("Secondary_Filename")&"' quality=high WIDTH='40' HEIGHT='40' NAME='wimpy_button01' TYPE='application/x-shockwave-flash' PLUGINSPAGE='http://www.macromedia.com/go/getflashplayer'></EMBED>"
'				response.write "</OBJECT><br><font color='red'><b>"& chklangstr(1378) &"</b></font>"
'	'			<font color='red'><b> <- 若您無法正常播放考題請點選左邊圖示</b></font>
	
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
        <li><input type="radio" name="Answer" id="radio" value="<%=Choices1%>" />(a) <%=q%></li>
        <%end if
		'-------------------- 第一個選項  --------------end -------------------
		'-------------------- 第二個選項  --------------start -------------------
		if Choices2<>"" then
			q = getQuestionString(Choices2)
		 %>
        <li><input type="radio" name="Answer" id="radio" value="<%=Choices2%>" />(b) <%=q%></li>
        <%end if
		'-------------------- 第二個選項  --------------end -------------------
		'-------------------- 第三個選項  --------------start -------------------
		if Choices3<>"" then
			q = getQuestionString(Choices3)
		 %>
        	<li><input type="radio" name="Answer" id="radio" value="<%=Choices3%>" />(c) <%=q%></li>
        <%end if
		'-------------------- 第三個選項  --------------end -------------------
		'-------------------- 第四個選項  --------------start -------------------
		if Choices4<>"" then
			q = getQuestionString(Choices4)
		 %>
         	<li><input type="radio" name="Answer" id="radio" value="<%=Choices4%>" />(d) <%=q%></li>
        <%end if
		'-------------------- 第四個選項  --------------end -------------------
		%>
        </ul>
        
        <div class="testgo"><input type="button" name="send_answer" id="send_answer" value="Next" class="btn_2"/></div>
        </div>
        <!--測試題內容end-->
        </div>
        <!--英語能力測試內容end-->
        <input type="hidden" id="agree" name="agree" value="<%=agree%>">
        <input type="hidden" id="Qtype"  name="Qtype" value="<%=Qtype%>">
        <input type="hidden" id="QL" name="QL" value="<%=QL%>">
        <input type="hidden" id="QN" name="QN" value="<%=QN%>">
        <input type="hidden" id="LT" name="LT" value="<%=LT%>">
        <input type="hidden" id="QA" name="QA" value="<%=QA%>">
        <input type="hidden" id="cqstr" name="cqstr" value="<%=cqstr%>">
        <input type="hidden" id="Score_G" name="Score_G" value="<%=Score_G%>">
        <input type="hidden" id="Score_V" name="Score_V" value="<%=Score_V%>">
        <input type="hidden" id="Score_P" name="Score_P" value="<%=Score_P%>">
        <input type="hidden" id="Score_S" name="Score_S" value="<%=Score_S%>">
        <input type="hidden" id="Score_L" name="Score_L" value="<%=Score_L%>">
        <input type="hidden" id="Score_R" name="Score_R" value="<%=Score_R%>">
<!--    <input type="hidden" name="mh_usersn" value="<%'=mh_usersn%>">
        <input type="hidden" name="email" value="<%'=email%>">
        <input type="hidden" name="mh_userfullname" value="<%'=mh_userfullname%>"> -->
        </form>
       <!--內容end-->
         <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
    </div>
