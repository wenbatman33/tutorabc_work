<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%dirs="member"%>
<!--#include file="../../include/dir_lang.inc" -->
<!--#include file="../../include/private.inc" -->
<!--#include file="../../function/fn.asp" -->
<!--#include file="../../function/fn_member.asp" -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../../css/style.css" rel="stylesheet" type="text/css" />
<link href="../../css/menu.css" rel="stylesheet" type="text/css" />
<link href="../../css/smenu.css" rel="stylesheet" type="text/css" />
<link href="../../css/mmenu.css" rel="stylesheet" type="text/css" />
<link href="../../css/leftframe.css" rel="stylesheet" type="text/css" />
<link href="../../css/rightframe.css" rel="stylesheet" type="text/css" />
<link href="../../css/footer.css" rel="stylesheet" type="text/css" />
<link href="../../css/topbutton.css" rel="stylesheet" type="text/css" />
<link href="../../css/scrollover.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../../javascript/blur.js"></script>
<script type="text/javascript" src="../../javascript/scrollovers.js"></script>
<style type="text/css">
<!--
.tr_style1 {background-color: #F9F5D9}
.tr_style2 { background-color: #FFF9C6}
.style3 {color: #A0A0A4}
-->
</style>
<title><%=char00392%></title>

</head>
<%
client_sn = request.cookies("jr_csn")
%>
<body>
<div id="container"  >
        <div id="leftframe" >  
  <!--#include file="../../include/logo.asp" -->

			<!--#include file="../../include/member_menu.asp" -->
			 <!--#include file="../../include/word_tutor.asp" -->
			<!--#include file="../../include/word.asp" -->
            <!--#include file="../../include/login.asp" -->
                        <!--#include file="../../include/learning.asp" -->
			 <!--#include file="../../include/memberads.asp" -->
            <!--#include file="../../include/testimonials.asp" -->
            <!--#include file="../../include/puzzle.asp" -->        	

        </div>      
        <div id="rightframe">
        <%p_index = 4%>
            <!--#include file="../../include/top.asp" -->
            <div id="content" >
            <p class="contenttitle"><img src="/images/title01_36.gif" /><%'=char00187 %><%'Placement Test %></p>
			 <div class="tabtab"  >  

<%
'Response.write request.cookies("jr_csn")
'**************************************************
'參數設定

'Qorder : 考題類型順序
'Qtype : 考題項次
'QMN : 每階段考題最大值
'QN : 考題編號
'QL : 考題LEVEL
'shotclock : 計數器時間
'rclick : 是否開放右鍵 true為關閉右鍵功能
Qorder=array(0,2,1,3,4,5,6)
'選3會考4題
QMN=2
Qtype=request("Qtype")
if Qtype="" then Qtype=1
shotclock="00:55"
QL=request("QL")
if QL="" then QL=3

QN=request("QN")
if QN="" then QN=1

if(request.cookies("jr_csn") <> "") then
	mh_usersn = request.cookies("jr_csn")
	email = session("email")
	mh_userfullname = session("mh_userfullname")
else
	mh_usersn = request("mh_usersn")
	email = request("email")
	mh_userfullname = request("mh_userfullname")
end if


shotclock="00:25"
if Qtype<>4 and Qtype<>5 and QL=1 then shotclock="00:25"
if Qtype<>4 and Qtype<>5 and QL=2 then shotclock="00:30"
if Qtype<>4 and Qtype<>5 and QL=3 then shotclock="00:35"
if Qtype<>4 and Qtype<>5 and QL=4 then shotclock="00:40"
if Qtype<>4 and Qtype<>5 and QL=5 then shotclock="00:45"
if Qtype<>4 and Qtype<>5 and QL=6 then shotclock="00:55"

if Qtype=5 and QL=1 then shotclock="00:35"
if Qtype=5 and QL=2 then shotclock="00:35"
if Qtype=5 and QL=3 then shotclock="00:40"
if Qtype=5 and QL=4 then shotclock="00:45"
if Qtype=5 and QL=5 then shotclock="00:55"
if Qtype=5 and QL=6 then shotclock="00:55"


if Qtype=4 and QL=1 then shotclock="00:30"
if Qtype=4 and QL=2 then shotclock="00:35"
if Qtype=4 and QL=3 then shotclock="00:35"
if Qtype=4 and QL=4 then shotclock="00:40"
if Qtype=4 and QL=5 then shotclock="00:45"
if Qtype=4 and QL=6 then shotclock="00:50"

'shotclock="00:55"
rclick="true"
'css="'MS Sans Serif' style='font-size: 12pt'"




'response.write Qtype(2)
'**************************************************


'--------------------------------------------------
'答案比對作業
if request("step")=2 then
%>
<form name="mf" method="post" action="placement_test.asp">
<%
	Qtype=request("Qtype")
	QL=request("QL")
	QN=request("QN")
	QA=request("QA")
	AR=request("Answer")
	cqstr=request("cqstr")
	Score_G=request("Score_G")
	Score_V=request("Score_V")
	Score_P=request("Score_P")
	Score_L=request("Score_L")
	Score_R=request("Score_R")
	Score_S=request("Score_S")
	if Score_G="" then Score_G=0
	if Score_V="" then Score_V=0
	if Score_P="" then Score_P=0
	if Score_L="" then Score_L=0
	if Score_R="" then Score_R=0
	if Score_S="" then Score_S=0

	if int(QN)<=int(QMN) then
		QN=QN+1
		if QA=AR then
			if Qorder(Qtype)=3 and QL>6 then QL=6
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
		if Qorder(Qtype)=1 then Score_G=int(Score_G)+QL
		if Qorder(Qtype)=2 then Score_V=int(Score_V)+QL
		if Qorder(Qtype)=3 then Score_P=int(Score_P)+QL
		if Qorder(Qtype)=4 then Score_L=int(Score_L)+QL
		if Qorder(Qtype)=5 then Score_R=int(Score_R)+QL
		if Qorder(Qtype)=6 then Score_S=int(Score_S)+QL
'		考完3題後進行下一部分
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
<input type="hidden" name="agree" value="<%=request("agree")%>">
<input type="hidden" name="Qtype" value="<%=Qtype%>">
<input type="hidden" name="QL" value="<%=QL%>">
<input type="hidden" name="QN" value="<%=QN%>">
<input type="hidden" name="LT" value="<%=LT%>">
<input type="hidden" name="cqstr" value="<%=cqstr%>">
<input type="hidden" name="Score_G" value="<%=Score_G%>">  <!--'G文法-->
<input type="hidden" name="Score_V" value="<%=Score_V%>">  <!--'V字彙-->
<input type="hidden" name="Score_P" value="<%=Score_P%>">  <!--'P發音-->
<input type="hidden" name="Score_S" value="<%=Score_S%>">  <!--'S會話-->
<input type="hidden" name="Score_L" value="<%=Score_L%>">  <!--'L聽力-->
<input type="hidden" name="Score_R" value="<%=Score_R%>">  <!--'R閱讀-->
<input type="hidden" name="mh_usersn" value="<%=mh_usersn%>">
<input type="hidden" name="email" value="<%=request.cookies("jr_mail")%>">
<input type="hidden" name="mh_userfullname" value="<%=request.cookies("jr_name")%>">
</form>
<script language="JavaScript"><!--
document.mf.submit();
// --></script>

<%	
'--------------------------------------------------
'出題作業
else

%>
<script language="JavaScript"><!--
function soundPlay(sndName)
{
	mySND.src = sndName;
}
// --></script>
<%
Qtype=request("Qtype")
if Qtype="" then Qtype=1
if Qorder(Qtype)=1 then Testname="Grammar - 文 法"
if Qorder(Qtype)=2 then Testname="Vocabulary - 字 彙"
if Qorder(Qtype)=3 then Testname="Pronunciation - 發 音"
if Qorder(Qtype)=4 then Testname="Listening - 聽 力"
if Qorder(Qtype)=5 then Testname="Reading - 閱 讀"
if Qorder(Qtype)=6 then Testname="Speaking - 會 話"

QL=request("QL")
if QL="" then QL=3
if QL=1 then qlstr=" and (client_adaptive_question.QLevel LIKE '%-1-%' OR client_adaptive_question.QLevel LIKE '%-2-%')"
if QL=2 then qlstr=" and (client_adaptive_question.QLevel LIKE '%-3-%' OR client_adaptive_question.QLevel LIKE '%-4-%')"
if QL=3 then qlstr=" and (client_adaptive_question.QLevel LIKE '%-5-%' OR client_adaptive_question.QLevel LIKE '%-6-%')"
if QL=4 then qlstr=" and (client_adaptive_question.QLevel LIKE '%-7-%' OR client_adaptive_question.QLevel LIKE '%-8-%')"
if QL=5 then qlstr=" and (client_adaptive_question.QLevel LIKE '%-9-%' OR client_adaptive_question.QLevel LIKE '%-10-%')"
if QL=6 then qlstr=" and (client_adaptive_question.QLevel LIKE '%-11-%' OR client_adaptive_question.QLevel LIKE '%-12-%')"

QN=request("QN")
if QN="" then QN=1


LT=request("LT")
if LT="" then LT="tw"
%>
<%
'sql="SELECT * FROM client_adaptive_question LEFT OUTER JOIN cfg_adaptive_type ON client_adaptive_question.QType = cfg_adaptive_type.type_no WHERE (client_adaptive_question.Valid = 1) AND (LEFT(client_adaptive_question.QType, 1) = '"&Qorder(Qtype)&"')"&qlstr&"  AND (client_adaptive_question.Approver IS NOT NULL) "
'sql="SELECT * FROM client_adaptive_question LEFT OUTER JOIN cfg_adaptive_type ON client_adaptive_question.QType = cfg_adaptive_type.type_no WHERE (client_adaptive_question.Valid = 1) AND (LEFT(client_adaptive_question.QType, 1) = '"&Qorder(Qtype)&"')"&qlstr&" AND (Secondary_Filename <> 'mp3')"
sql="SELECT TOP (1) client_adaptive_question.sn, client_adaptive_question.secondary_filename, client_adaptive_question.sqsn, client_adaptive_question.qtype, client_adaptive_question.qlevel, client_adaptive_question.question, client_adaptive_question.choices1, client_adaptive_question.choices2, client_adaptive_question.choices3, client_adaptive_question.choices4, client_adaptive_question.choices5, client_adaptive_question.choices6, client_adaptive_question.answer, client_adaptive_question.Flag, client_adaptive_question.Inpdate, client_adaptive_question.Inputer, client_adaptive_question.Approver, client_adaptive_question.Approved_Date, client_adaptive_question.valid, cfg_adaptive_type.type_no, cfg_adaptive_type.etype, cfg_adaptive_type.ctype, cfg_adaptive_type.cntype, cfg_adaptive_type.jptype, cfg_adaptive_type.krtype, cfg_adaptive_type.valid AS Expr1 FROM client_adaptive_question LEFT OUTER JOIN cfg_adaptive_type ON client_adaptive_question.qtype = cfg_adaptive_type.type_no WHERE (client_adaptive_question.valid = 1) AND (LEFT(client_adaptive_question.qtype, 1) = '"&Qorder(Qtype)&"')"&qlstr&" AND (client_adaptive_question.secondary_filename <> 'mp3' OR client_adaptive_question.secondary_filename IS NULL) "
'response.write sql
'if qtype=5 then
'response.end
'end if
cqstr=request("cqstr")
if cqstr<>"" then sql=sql&cqstr
'response.write sql&"<br>"
'response.end

sql=sql&" ORDER BY NEWID() "
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

Score_G=request("Score_G")
Score_V=request("Score_V")
Score_P=request("Score_P")
Score_S=request("Score_S")
Score_L=request("Score_L")
Score_R=request("Score_R")
if Score_G="" then Score_G=0
if Score_V="" then Score_V=0
if Score_P="" then Score_P=0
if Score_S="" then Score_S=0
if Score_L="" then Score_L=0
if Score_R="" then Score_R=0

'如果已經沒有題目時
'RESPONSE.WRITE sql
if rs.eof then
	response.write "Stop"
	response.redirect "evaluation_result.asp?mh_usersn="&request.cookies("jr_csn")&"&email="&request.cookies("jr_mail")&"&mh_userfullname="&request.cookies("jr_name")&"&Score_G="&Score_G&"&Score_V="&Score_V&"&Score_P="&Score_P&"&Score_S="&Score_S&"&Score_L="&Score_L&"&Score_R="&Score_R
else
	total_question=rs.recordcount

	if QN=1 then
		cqstr=""
	else
		cqstr=request("cqstr")
	end if
	Randomize
	Qsn = Int((total_question * Rnd) + 1)
	if Qsn>1 then
		for i=1 to Qsn-1
			rs.movenext
		next
	end if
	cqstr=cqstr&" and sn<>"&rs("sn")
	QA=rs("answer")
%>

<%
if rclick="true" and 2=1 then
%>

<bg sound src="silent.aif" id="mySND">

<%'防止使用右鍵%>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" onContextmenu="return false"  style="background-image: url('/images/bg.jpg')">
<%
end if
%>
		<form name="timers" method="POST" action="placement_test.asp?step=2&language=<%=request("language")%>">			        
            <table border="0" align="right" cellpadding="0" cellspacing="0" >
              <tr>
                <td width="10"><img src="/images/tab<%if qtype=1 then response.write "in" else response.write "out"%>_left.gif" /></td>
                <td width="50" background="/images/tab<%if qtype=1 then response.write "in" else response.write "out"%>_bg.gif"><div align="center" class="tabtable"><a href="profile.asp"><%=char00144 %></a></div></td>
                <%
                if qtype=1 then 
                    pic =  "inout" 
                elseif qtype=2 then
                    pic =  "outin"
                    
                else
                    pic = "outout"
                end if
                 %>
                <td width="10"><img src="/images/tab<%=pic%>_right.gif" /></td>
                <td width="50" background="/images/tab<%if qtype=2 then response.write "in" else response.write "out"%>_bg.gif"><div align="center" class="tabtable"><%=char00143 %></div></td>
                <%
                if qtype=2 then 
                    pic =  "inout" 
                elseif qtype=3 then
                    pic =  "outin"
                    
                else
                    pic = "outout"
                end if
                 %>
                <td width="10"><img src="/images/tab<%=pic%>_right.gif" /></td>             
                <td width="50" background="/images/tab<%if qtype=3 then response.write "in" else response.write "out"%>_bg.gif"><div align="center" class="tabtable"><%=char00145 %></div></td>
                 <%
                if qtype=3 then 
                    pic =  "inout" 
                elseif qtype=4 then
                    pic =  "outin"
                    
                else
                    pic = "outout"
                end if
                 %>
                <td width="10"><img src="/images/tab<%=pic%>_right.gif" /></td>
                <td width="50" background="/images/tab<%if qtype=4 then response.write "in" else response.write "out"%>_bg.gif"><div align="center" class="tabtable"><%=char00148 %></div></td>
                <%
                if qtype=4 then 
                    pic =  "inout" 
                elseif qtype=5 then
                    pic =  "outin"
                    
                else
                    pic = "outout"
                end if
                 %>
                <td width="10"><img src="/images/tab<%=pic%>_right.gif" /></td>
                <td width="50" background="/images/tab<%if qtype=5 then response.write "in" else response.write "out"%>_bg.gif"><div align="center" class="tabtable"><%=char00147 %></div></td>
                <td width="10"><img src="/images/tab<%if qtype=5 then response.write "in" else response.write "out"%>_right.gif" /></td>
              </tr>
            </table>
            </div>
			<br clear="all" />
			<table width="700" border="0" align="center" cellpadding="0" cellspacing="2" class="mtable" >
					
					
				<tr><td style="text-align: left;">
				
				<%
if instr(rs("question"),"|__|")>0 then
	Qfile=rs("secondary_filename")

	q=replace(rs("question"),"<","@")
	q=replace(q,">","*")
	q=replace(q,"@","<u>")
	q=replace(q,"*","</u>")
	q=replace(q,"[","<br><font color=red><b>")
	q=replace(q,"]"," </b></font>")
	q=replace(q,"B:","<br>B:")
	q=replace(q,"C:","<br>C:")
	response.write q&"<br>"
else
	q=replace(rs("question"),"<","@")
	q=replace(q,">","*")	
	q=replace(q,"@","<u>")
	q=replace(q,"*","</u>")
	q=replace(q,"[","<br><font color=red><b>")
	q=replace(q,"]"," </b></font>")
	q=replace(q,"B:","<br>B:")
	q=replace(q,"C:","<br>C:")
	if q<>"jpg" and q<>"mp3" then
		response.write q & "<br>"
	end if
	Qfile=rs("secondary_filename")

end if
        %><!--Images--><%
if not isnull(rs("secondary_filename")) and rs("secondary_filename")<>"" then

	if right(rs("secondary_filename"),3)="jpg" then
		response.write "<img src=http://www.muchtalk.com/consultant/asp/q_files/images/" & rs("secondary_filename") & " width=120 height=120>"
	else
	
		response.write "<OBJECT classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,47,0' WIDTH='40' HEIGHT='40' id='wimpy_button_1' name='wimpy_button01'>"
		response.write "<PARAM NAME=movie VALUE='/flash/wimpy_button.swf?theFile=../../sound/"&rs("Secondary_Filename")&"'>"
		response.write "<PARAM NAME=quality VALUE=high>"
		response.write "<PARAM NAME=wmode VALUE=transparent>"
		response.write "<EMBED src='/flash/wimpy_button.swf?theFile=../../sound/"&rs("Secondary_Filename")&"' quality=high WIDTH='40' HEIGHT='40' NAME='wimpy_button01' TYPE='application/x-shockwave-flash' PLUGINSPAGE='http://www.macromedia.com/go/getflashplayer'></EMBED>"
		response.write "</OBJECT>"& char00293
'		<font color='red'><b> <- 若您無法正常播放考題請點選左邊圖示</b></font>

	end if
end if

%>
<!--Images
請為題目中的註解找出同義的字彙.<br><br>請用滑鼠點一下喇叭圖示, 並依照您所聽到的訊息來做答. 請點選出您所聽到的字或句子.<br><br><img border="0" src="../../images/speaker.gif" width="25" height="25" style="cursor:hand">
-->
	                
            </td>
				</tr>
					
						<table width="700" border="0" align="center" cellspacing="0" class="mtable" style="text-align: left;"\>
              <tbody>
<%
if not isnull(rs("Choices1")) AND rs("Choices1")<>"" then
%>
								<tr>
									<td width="5%"><b>(a)</b></td>
									<td align="center" width="5%"><input type="radio" value="<%=rs("Choices1")%>" name="answer"></td>
									<td style="text-align: left">
<%
	q=replace(rs("Choices1"),"<","@")
	q=replace(q,">","*")
	q=replace(q,"@","<u>")
	q=replace(q,"*","</u>")
	q=replace(q,"[","<font color=red><b>")
	q=replace(q,"]"," </b></font>")
	q=replace(q,"B:","<br>B:")
	q=replace(q,"C:","<br>C:")
	response.write q
%>
									</td>
								</tr>
<%
end if
%>
<%
if not isnull(rs("Choices2")) AND rs("Choices2")<>"" then
%>
								<tr>
									<td  width="5%"><b>(b)</b></td>
									<td align="center" width="5%"><input type="radio" value="<%=rs("Choices2")%>" name="answer"></td>
									<td style="text-align: left">
<%
	q=replace(rs("Choices2"),"<","@")
	q=replace(q,">","*")
	q=replace(q,"@","<u>")
	q=replace(q,"*","</u>")
	q=replace(q,"[","<font color=red><b>")
	q=replace(q,"]"," </b></font>")
	q=replace(q,"B:","<br>B:")
	q=replace(q,"C:","<br>C:")
	response.write q
%>
									</td>
								</tr>
<%
end if
%>
<%
if not isnull(rs("Choices3")) AND rs("Choices3")<>"" then
%>
								<tr>
									<td  width="5%"><b>(c)</b></td>
									<td align="center" width="5%"><input type="radio" value="<%=rs("Choices3")%>" name="answer"></td>
									<td  style="text-align: left">
<%
	q=replace(rs("Choices3"),"<","@")
	q=replace(q,">","*")
	q=replace(q,"@","<u>")
	q=replace(q,"*","</u>")
	q=replace(q,"[","<font color=red><b>")
	q=replace(q,"]"," </b></font>")
	q=replace(q,"B:","<br>B:")
	q=replace(q,"C:","<br>C:")
	response.write q
%>
									</td>
								</tr>
<%
end if
%>
<%
if not isnull(rs("Choices4")) AND rs("Choices4")<>"" then
%>
								<tr>
									<td  width="5%"><b>(d)</b></td>
									<td align="center" width="5%"><input type="radio" value="<%=rs("Choices4")%>" name="answer"></td>
									<td  style="text-align: left">
<%
	q=replace(rs("Choices4"),"<","@")
	q=replace(q,">","*")
	q=replace(q,"@","<u>")
	q=replace(q,"*","</u>")
	q=replace(q,"[","<font color=red><b>")
	q=replace(q,"]"," </b></font>")
	q=replace(q,"B:","<br>B:")
	q=replace(q,"C:","<br>C:")
	response.write q
%>
									</td>
								</tr>
<%
end if
%>
								
								<tr>
									<td  colspan="3">
									<input type="submit" value="<%=char00055%>" name="B1" style="font-family: Arial; font-size: 9pt"></td>
								</tr>
								 </tbody>
							</table>
					
<!--答案區塊-->
<%
'response.write "<font color='FFFFFF'>level : "&QL&"</font>"
%>
</body>
<input type="hidden" name="agree" value="<%=agree%>">
<input type="hidden" name="Qtype" value="<%=Qtype%>">
<input type="hidden" name="QL" value="<%=QL%>">
<input type="hidden" name="QN" value="<%=QN%>">
<input type="hidden" name="LT" value="<%=LT%>">
<input type="hidden" name="QA" value="<%=QA%>">
<input type="hidden" name="cqstr" value="<%=cqstr%>">
<input type="hidden" name="Score_G" value="<%=Score_G%>">
<input type="hidden" name="Score_V" value="<%=Score_V%>">
<input type="hidden" name="Score_P" value="<%=Score_P%>">
<input type="hidden" name="Score_S" value="<%=Score_S%>">
<input type="hidden" name="Score_L" value="<%=Score_L%>">
<input type="hidden" name="Score_R" value="<%=Score_R%>">
<input type="hidden" name="mh_usersn" value="<%=mh_usersn%>">
<input type="hidden" name="email" value="<%=request.cookies("jr_mail")%>">
<input type="hidden" name="mh_userfullname" value="<%=request.cookies("jr_name")%>">
</form>
<%
end if

end if
conn.close
%>							
		</td></tr>			

	</table>
            </div>                                                				                        
        </div>
    </div>               
<!--#include file="../../include/untop.asp" -->
</body>
<script type="text/javascript" src="../../javascript/evenrows.js"></script>

</html>