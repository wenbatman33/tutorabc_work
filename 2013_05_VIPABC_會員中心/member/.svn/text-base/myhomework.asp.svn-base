<!--#include virtual="/include/global.inc"-->
<!--#include virtual="/include/StringReplace.inc"-->


<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--<meta http-equiv="Content-Type" content="text/html; charset=utf-8">-->
<!--#include file="../../include/setcookie.inc" -->

	
<% 
Server.ScriptTimeout=60000

islock = false
if(Session("mh_usersn")="") then islock = true
if(request("client_sn")<>"" or request("client")<>"") then islock = false



IF islock Then

   Response.write "<script language='JavaScript'>alert('親愛的會員:\n\n由於您太久未操作網頁.\n\n系統將無法為您開啟Homework頁面\n\n請重新登入謝謝!!');window.close()</script>"
Else


	%>
	<!--#include file="../../tools/cdic/fn.inc"-->
	
	<!--#include file="../../include/conn_temp.inc" -->
	<!--#include file="../../tools/webconfig/functions_chkwebconfiginifile.asp"-->
	<!--#include file="../../tools/lang/functions_chklangstrfortutorabc.asp"-->
	
	<%
	mhour=request("stime")
	mdate=request("sdate")
	session_sn=request("session_sn")
	'session_sn="2005040416001"
	if(Session("client_sn") <> "") then client_sn=session("client_sn")
	if(request("client_sn") <> "") then 
		client_sn=request("client_sn")
	end if
    
    if(not isnumeric(client_sn)) then
    client_sn=trim(Session("client_sn"))
    end if

'	if(request("client") <> "") then 
'		set obj = server.createobject("baseControls.security")
'		client_sn=obj.decode(request("client"))
'		set obj = nothing
'	end if
	
	if(not isnumeric(client_sn)) then
		response.write "<script language='javascript'>"
		response.write "window.close();"
		response.write "</script>"
		response.end
	end if	
	
	sql="SELECT client_attend_list.Session_sn, dbo.client_attend_list.client_sn, dbo.client_attend_list.Attend_Date, dbo.client_attend_list.Attend_Sestime, dbo.client_attend_list.Attend_consultant, dbo.client_attend_list.Attend_mtl_1, dbo.client_attend_list.Attend_Level, dbo.con_basic.basic_fname + ' ' + dbo.con_basic.basic_lname AS Expr1, dbo.material.LTitle, dbo.material.LD, dbo.client_homework.Score, dbo.client_homework.anser,dbo.client_homework.rdn, dbo.client_homework.Sn, dbo.client_homework.voc_sn FROM dbo.client_attend_list LEFT OUTER JOIN dbo.client_homework ON dbo.client_attend_list.client_sn = dbo.client_homework.Client_Sn AND dbo.client_attend_list.Session_sn = dbo.client_homework.Session_Sn LEFT OUTER JOIN dbo.material ON dbo.client_attend_list.Attend_mtl_1 = dbo.material.Course LEFT OUTER JOIN dbo.con_basic ON dbo.client_attend_list.Attend_consultant = dbo.con_basic.Con_Sn WHERE (dbo.client_attend_list.Session_sn = '"&session_sn&"') AND (dbo.client_attend_list.client_sn = "&client_sn&")"
	'暫時先不判斷有沒有上課
	' AND (dbo.client_attend_list.Class_yesno = '9')
	set rs=Server.CreateObject("Adodb.RecordSet")
	'if session("client_sn")="145390" then
	'response.write sql & "<br>"
	'response.end
	'end if
	rs.open sql,conn,1,1
	
	if not rs.eof then
		arr_tmp = array("0","0","0","0","0","0","0","0","0","0")
		isanser = ""
		
		if(not isNull(rs("anser"))) then
			isanser = "1"
			arr_tmp = split(rs("anser"),",")
			rdn = rs("rdn")
			vocsn = rs("voc_sn")
		else
			Randomize
			rdn=int(rnd*5)+1
			vocsn = ""
		end if
	
	course=rs("attend_mtl_1")
	title = rs("ltitle")
	'20091229 Lucy start
	If (title<>"" And Not isnull(title)) Then
		title = StringReplace(title)
	End If
	'20091229 Lucy end
	sessionTime=rs("attend_date")&" - "&rs("attend_sestime")&":30"
	consultant=rs("expr1")
	level=rs("attend_level")
	client_level=rs("attend_level")
	description = rs("ld")
	'20091229 Lucy start
	If (description<>"" And Not isnull(description)) Then
		description = StringReplace(description)
	End If
	'20091229 Lucy end
	hsn=rs("sn")
	client_sn=rs("client_sn")
	Session_Sn=rs("Session_Sn")
	
	'問題的陣列
	dim Q()
	redim Q(10)
	'答案的陣列
	dim A()
	redim A(10)
	'選項的陣列
	dim C()
	redim C(10)
	'亂數的陣列
	dim R()
	redim R(10)
	'真實答案的陣列
	dim T()
	redim T(10)
	'單字的陣列
	Dim V
	dim vtr()
	redim vtr(5)
	
	sql="SELECT * FROM lesson_question WHERE (Course = '"&course&"') and (Question <> '')"
	set rs=Server.CreateObject("Adodb.RecordSet")
	'response.write sql & "<br>"
	rs.open sql,conn,1,1
	
	for i=1 to 5
		if not rs.eof then
            if (rs("question") <> "" and Not IsNull(rs("question"))) then
			    Q(i) = StringReplace(rs("question"))
            else
			    Q(i) = rs("question")
            end if
			mcstr="<font class='answ'>"
			rastr="<font class='answ'>"
			for j=1 to 5
                str_choices = rs("choices"&j)
				
				'20091228 Lucy start (template變數)
				str_answer = rs("answer")
				if str_choices<>"" and not isnull(str_choices) then
                    str_choices = StringReplace(str_choices)
					if(trim(arr_tmp(i-1)) <> "") then
						if(j=cint(arr_tmp(i-1))) then
							issel = "checked"
						else
							issel = ""
						end if
					end if
					mcstr=mcstr&"<input type='radio' value='"&j&"' name='MQ"&i&"' "&issel&" >&nbsp;"&str_choices&"&nbsp;"
					
					'20091228 Lucy start
					If (str_answer <> "" And Not IsNull(str_answer)) Then
						str_answer = StringReplace(str_answer)
					End If
					
					If str_choices=str_answer Then
						A(i)=j
						rastr=rastr&"<input type='radio' name='X"&i&"' disabled=true> <font color='blue'>"&str_choices&"</font> "
					Else
						rastr=rastr&"<input type='radio' name='X"&i&"' disabled=true> "&str_choices&" "
					End If
					'20091228 Lucy end
				end if
			next
			mcstr=mcstr&"</font>"
			rastr=rastr&"</font>"
			C(i)=mcstr
			T(i)=rastr
			if rs.recordcount>i then rs.movenext
		end if
	next
	
	'亂數的首值為
	if rdn=1 then
		V=Array("","2","1","5","4","3")
		AA=Array("","B","A","E","D","C")
	end if
	if rdn=2 then
		V=Array("","1","4","3","2","5")
		AA=Array("","A","D","C","B","E")
	end if
	if rdn=3 then
		V=Array("","3","2","1","4","5")
		AA=Array("","C","B","A","D","E")
	end if
	if rdn=4 then
		V=Array("","4","5","2","3","1")
		AA=Array("","E","C","D","A","B")
	end if
	if rdn=5 then
		V=Array("","5","1","3","2","4")
		AA=Array("","B","D","C","E","A")
	end if

'	response.write rdn
	
	'=====不管舊有單字=========================================================
	'sql="SELECT * FROM Material_vocabulary WHERE (Course = '"&course&"') "
	'20050404新增判斷需要中文釋義和英文釋義都要有的情況下才可以執行.
	'sql=sql & "AND (c_dictionary_state <> '0') AND (e_dictionary_state <> '0')"
	
	'set rs=Server.CreateObject("Adodb.RecordSet")
	'response.write sql & "<br>"
	'rs.open sql,conn,3,3
	
	
	'i=6
	'while not rs.eof
	'if 6+int(rs.recordcount)>10 then maxno=10 else maxno=6+int(rs.recordcount)
	'for i=6 to maxno
	'	Q(i)=rs("vocabulary")
	'	keyword=rs("vocabulary")
	
	'	vtr(i-5)=v_content
	'	i=i+1
	'	if i<maxno then rs.movenext
	'next
	'wend
	'=====不管舊有單字=========================================================
	
	if rs.recordcount <5 then lossno=5-rs.recordcount else lossno=0
	lossno=5
	isvoc = false
	
	if(vocsn = "") then
		'sql="SELECT TOP  "& lossno  &" sn,  vocabulary_keyword FROM cfg_vocabulary_consult WHERE (vocabulary_level = "&level&") ORDER BY  NEWID()"
		'sql="SELECT TOP "& lossno  &" sn, vocabulary AS vocabulary_keyword FROM material_vocabulary WHERE (course = "&course&") AND (c_dictionary_state = '1') OR (course = "&course&") AND (e_dictionary_state = '1')"
		'工單PM09070203 DENEY修改 ADD -START
		sql="SELECT TOP "& lossno  &" sn, vocabulary AS vocabulary_keyword FROM material_vocabulary WHERE (course = "&course&" AND c_dictionary_state = '1'  AND valid = '1') OR (course = "&course&" AND e_dictionary_state = '1' AND valid = '1') or (valid='1' and is_checked='1')"
		'工單PM09070203 DENEY修改 ADD -END
	else
		'sql = "SELECT TOP  "& lossno  &" sn,  vocabulary_keyword FROM cfg_vocabulary_consult WHERE sn in(" & vocsn & ")"	
		sql = "SELECT TOP  "& lossno  &" sn, vocabulary as vocabulary_keyword FROM material_vocabulary WHERE sn in(" & vocsn & ")"	
		arr_voc = split(vocsn,",")
		tmpvoc = ""
		for ii=0 to ubound(arr_voc)
			tmpvoc = tmpvoc & " when " & arr_voc(ii) & " then " & ii
		next
		tmpvoc = "order by (case sn " & tmpvoc & " end) "
		sql = sql & tmpvoc
		isvoc = true
	end if

	if level=88 then sql="SELECT Sn, Vocabulary as vocabulary_keyword FROM Material_vocabulary WHERE (Course = 100000)"

	set rs=Server.CreateObject("Adodb.RecordSet")
	'if(session("client_sn") = 101539) then
		'response.write sql & "<br>"
	'end if
	rs.open sql,conn,1,1
	'response.write rdn&"<br>"
'	x=1
	while not rs.eof
		Q(i)=rs("vocabulary_keyword")
		keyword=rs("vocabulary_keyword")
'		if x=1 then keyword="satisfy"
'		if x=2 then keyword="surround"
'		if x=3 then keyword="accompany"
'		if x=4 then keyword="telephone"
'		if x=5 then keyword="through"
		if(not isvoc) then vocsn = vocsn & rs("sn") & ","
	%>
	<!--#include file="../../tools/cdic/functions_chkdictionaryforv.asp"-->
	<%
'		response.write v_content&"<br><br>"
		vtr(i-5)=v_content
'		x=x+1
		i=i+1
		rs.movenext
	wend
'	題目出現的順序
	'response.write vocsn
	'response.write rdn&"<br>"
	anw=v(1)&","&v(2)&","&v(3)&","&v(4)&","&v(5)
	if(not isvoc and not isnull(vocsn) and vocsn <> "") then vocsn = left(vocsn, len(vocsn) - 1)
'	if(not isvoc) and len(vocsn)>1 then vocsn = left(vocsn, len(vocsn) - 1)


'	單字選項的順序	
	for i=6 to 10
		if V(i-5)="1" then C(i)=vtr(1)
		if V(i-5)="2" then C(i)=vtr(2)
		if V(i-5)="3" then C(i)=vtr(3)
		if V(i-5)="4" then C(i)=vtr(4)
		if V(i-5)="5" then C(i)=vtr(5)
	next
	
'	response.write "vtr(1):"&vtr(1)&"<br>C:"&C(6)&"<br>"&"<br>"
'	response.write "vtr(2):"&vtr(2)&"<br>C:"&C(7)&"<br>"&"<br>"
'	response.write "vtr(3):"&vtr(3)&"<br>C:"&C(8)&"<br>"&"<br>"
'	response.write "vtr(4):"&vtr(4)&"<br>C:"&C(9)&"<br>"&"<br>"
'	response.write "vtr(5):"&vtr(5)&"<br>C:"&C(10)&"<br>"&"<br>"
	
	
	R(6)="A"
	R(7)="B"
	R(8)="C"
	R(9)="D"
	R(10)="E"
	
'	point=request("point")
	%>
	<html>
	<head>
	<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
    <meta name="ProgId" content="FrontPage.Editor.Document">
    
	<title>TutorABC Homework</title>
	<style type="text/css">
	<!--
		.title	 { font-size : 18pt; font-family : Verdana; font-weight : bold; }
		.type	 { font-size : 11pt; font-family : Verdana; font-weight : bold; color : #333333; }
		.ques	 { font-size : 9pt; font-family : Verdana; font-weight : bold; }
		.answ    { font-size : 9pt; font-family : Verdana; color : #333333; }
	-->
	
	</STYLE>
	</head>
	
	<body>
	
	<div align="center">
	  <center>
	  <form name='WebForm' method="POST" action="myhomework1.asp">
	  
	  <table border="0" cellspacing="1" width="550" id="AutoNumber1">
	    <tr>
	      <td width="100%" style="padding-right: 10" align="right"><font face="Verdana" style="font-size: 9pt">本課後作業內容由 <a style="color: #4166af" href="http://<%=abc_webhost_name%>/">TutorABC</a> 提供</font></td>
	    </tr>
	    <tr>
	      <td width="100%">
	      <table border="0" cellspacing="0" width="100%" cellpadding="0" id="AutoNumber2">
	        <tr>
	          <td width="100%">
	      <table border="0" cellspacing="0" width="100%" cellpadding="0" id="AutoNumber3">
	        <tr>
	          <td width="100%" style="border: 2px ridge #B5B9AB">
	      	  <img border="0" src="<%=getinifilestr("webrootpath")%>images/tutorabc_mail_top.gif" width="550" height="70"></td>
	        </tr>
	        <tr>
	          <td width="100%" style="border-left: 2px ridge #B5B9AB; border-right: 2px ridge #B5B9AB; padding: 10">
	          <table border="0" cellspacing="1" width="100%" id="AutoNumber4" bgcolor="#9F0505">
	            <tr>
	              <td width="100%" bgcolor="#333333" colspan="2" align="center">
	              <font class="title" color="#FFFFFF">Session Information</font></td>
	            </tr>
	            <tr>
	              <td width="30%" bgcolor="#FFFFFF" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3" align="right">
	              <font face="Verdana" size="2">Course #</font></td>
	              <td width="70%" bgcolor="#FFFFFF" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3">
	              <font face="Verdana" size="2"><%=course%></font>&nbsp;&nbsp;<!--<a target="_blank" title="課程複習" href="http://<%=webhost_name%>/"><img border="0" src="<%=getinifilestr("webrootpath")%>images/camera.gif"></a>--></td>
	            </tr>
	            <tr>
	              <td width="30%" bgcolor="#CCCCCC" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3" align="right">
	              <font face="Verdana" size="2">Title</font></td>
	              <td width="70%" bgcolor="#CCCCCC" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3">
	              <font face="Verdana" size="2"><%=title%></font></td>
	            </tr>
	            <tr>
	              <td width="30%" bgcolor="#FFFFFF" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3" align="right">
	              <font face="Verdana" size="2">Session Time</font></td>
	              <td width="70%" bgcolor="#FFFFFF" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3">
	              <font face="Verdana" size="2"><%=sessiontime%></font></td>
	            </tr>
	            <tr>
	              <td width="30%" bgcolor="#CCCCCC" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3" align="right">
	              <font face="Verdana" size="2">Consultant</font></td>
	              <td width="70%" bgcolor="#CCCCCC" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3">
	              <font face="Verdana" size="2"><%=Consultant%></font></td>
	            </tr>
	            <tr>
	              <td width="30%" bgcolor="#FFFFFF" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3" align="right">
	              <font face="Verdana" size="2">Level</font></td>
	              <td width="70%" bgcolor="#FFFFFF" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3">
	              <font face="Times New Roman">
	<%
	for cl=1 to 12
		if cl=level then imgstr="c" else imgstr="g"
	%>
					<!--<img src="/images/strength/Old/<%=imgstr&cl%>.gif" border="0" width="18" height="18">-->
					<img src="<%=getinifilestr("webrootpath")%>images/strength/<%=imgstr&cl%>.gif" border="0" width="18" height="18">
	<%
	next
	%></font></td>
	            </tr>
	            <tr>
	              <td width="30%" bgcolor="#CCCCCC" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3" align="right">
	              <font face="Verdana" size="2">Description</font></td>
	              <td width="70%" bgcolor="#CCCCCC" style="padding-left: 10; padding-right: 10; padding-top: 3; padding-bottom: 3">
	              <font face="Verdana" size="2"><%=description%></font></td>
	            </tr>
	          </table>
	          </td>
	        </tr>
	        <tr>
	          <td width="100%" style="border-bottom:2px ridge #B5B9AB; border-left:2px ridge #B5B9AB; border-right:2px ridge #B5B9AB; padding:10; ">          
	          <table border="0" cellspacing="1" width="100%" id="AutoNumber5" height="150">
	<%
	if point<>"" then
	%>
	            <tr>
	              <td width="100%" colspan="2" height="29">
	              <font class="title">Your score is </font><u><font class="title" color="#ff0000"><%=point%></font></u></td>
	            </tr>
	            <tr>
	              <td width="100%" height="1" colspan="2" background="/images/bw_line.gif"></td>
	            </tr>
	<%
	end if
	%>
	            <tr>
	              <td width="100%" colspan="2" height="22">
	              <font class="type">Please answer these questions:</font></td>
	            </tr>
	            <tr>
	              <td width="100%" height="1" colspan="2" background="/images/bw_line.gif"></td>
	            </tr>
	<%
	for i=1 to 5
		if i=1 or i=3 or i=5 then mcstr=" bgcolor='#EEEFEF'" else mcstr=""
	%>
	            <tr>
	              <td width="5%" height="22" align="right"<%'=mcstr%>><font class="ques"><%=i%>.</font></td>
	              <td width="95%" height="22"<%'=mcstr%>>
	              <font class="ques">&nbsp;<%=Q(i)%></font></td>
	            </tr>
	            <tr>
	              <td width="5%" height="22"<%'=mcstr%>>　</td>
	              <td width="95%" height="22"<%'=mcstr%>><%=C(i)%></td>
	            </tr>
	            <tr>
	              <td width="100%" height="1" colspan="2" background="/images/bw_line.gif"></td>
	            </tr>
	<%
	next
	%>
	            <tr>
	              <td width="100%" height="22" colspan="2"></td>
	            </tr>
	            <tr>
	              <td width="100%" height="22" colspan="2"><font class="type">Please find the correct definition:</font></td>
	            </tr>
	            <tr>
	              <td width="100%" height="1" colspan="2" background="/images/bw_line.gif"></td>
	            </tr>
	<%
	for i=6 to 10
	%>
	            <tr>
	              <td width="5%" height="22" align="right">
	              <font class="ques"><%=i&"."%>
	              </font></td>
	              <td width="95%" height="22">
	              <table border="0" cellspacing="0" width="100%" cellpadding="0" id="AutoNumber6">
	                <tr>
					  <td width="6%" style="padding-left: 5; padding-right: 5">
	                  <select size="1" name="MQ<%=i%>" style="font-family: Verdana; font-size: 10 pt">
	                  <option value="">--</option>
					  <option value="A" <%if(arr_tmp(i-1)="A") then response.write "selected"%>>A</option>
	                  <option value="B" <%if(arr_tmp(i-1)="B") then response.write "selected"%>>B</option>
	                  <option value="C" <%if(arr_tmp(i-1)="C") then response.write "selected"%>>C</option>
	                  <option value="D" <%if(arr_tmp(i-1)="D") then response.write "selected"%>>D</option>
	                  <option value="E" <%if(arr_tmp(i-1)="E") then response.write "selected"%>>E</option>
	                  </select></font></td>
	                  <td width="22%" style="padding-left: 5; padding-right: 5"><font class="ques"><%=replace(Q(i),chr(34),"'")%><%'題目%></font></td>
	                  <td width="6%" style="padding-left: 5; padding-right: 5;" valign="middle">
	                  <img border="0" src="/images/<%=R(i)%>.gif" width="16" height="16"></td>
	                  <td width="66%" style="padding-left: 5; padding-right: 5">
	                  <font class="answ"><%=replace(C(i),chr(34),"'")%><%'答案%></font></td>
	                </tr>
					<tr>
					<td colspan="2" width="28%"></td>
					<td colspan="2" background="/images/bw_line.gif" width="72%"></td>
					</tr>
	              </table>
				  </td>
	            </tr>

	<%
	next
	%>
	            <tr>
	              <td width="100%" height="1" colspan="2"></td>
	            </tr>
	            <tr>
	              <td width="100%" height="3" colspan="2" align="center"></td>
	            </tr>
	            <tr>
	              <td width="100%" height="1" colspan="2" align="center"><input type="submit" value="Submit" name="B1" style="font-family: Verdana; font-size: 9 pt"></td>
	            </tr>
	          </table>
	          </td>
	        </tr>
	        <!--- <tr>
	          <td width="100%" style="border-bottom:2px ridge #B5B9AB; border-left:2px ridge #B5B9AB; border-right:2px ridge #B5B9AB; ">&nbsp;</td>
	        </tr> --->
	      </table>
	          </td>
	        </tr>
	        </table>
	      </td>
	    </tr>
	    <tr>
	      <td width="100%">
	      　</td>
	    </tr>
	    </table>
	<%
	for i=1 to 10	
	%>
		<input type="hidden" name="Q<%=i%>" value="<%=replace(replace(replace(replace(Q(i),chr(34),"'"),"’","#@"),"""","##"),"…","%^")%>"><%=chr(13)%>
	<%
	next
	%>
	<%
	for i=1 to 10	
	%>
		<input type="hidden" name="C<%=i%>" value="<%=replace(replace(C(i),chr(34),"'"),"""","|")%>"><%=chr(13)%>
	<%
	next
	%>
	<%
	for i=1 to 5	
	%>
		<input type="hidden" name="A<%=i%>" value="<%=A(i)%>"><%=chr(13)%>
	<%
	next
	%>
	<%
	for i=1 to 5	
	%>
		<input type="hidden" name="AA<%=i%>" value="<%=AA(i)%>"><%=chr(13)%>
	<%
	next
	%>
	<%
	for i=1 to 5	
	%>
		<input type="hidden" name="T<%=i%>" value="<%=replace(T(i),"’","#@")%>"><%=chr(13)%>
	<%
	next
	%>
	<%
	for i=1 to 5	
	%>
		<input type="hidden" name="V<%=i%>" value="<%=V(i)%>"><%=chr(13)%>
	<%
	next
	%>
		<input type="hidden" name="rdn" value="<%=rdn%>">
		<input type="hidden" name="anw" value="<%=anw%>">
		<input type="hidden" name="course" value="<%=course%>">
		<input type="hidden" name="mdate" value="<%=mdate%>">
		<input type="hidden" name="mhour" value="<%=mhour%>">
		<input type="hidden" name="title" value="<%=title%>">
		<input type="hidden" name="sessiontime" value="<%=sessiontime%>">
		<input type="hidden" name="consultant" value="<%=consultant%>">
		<input type="hidden" name="level" value="<%=level%>">
		<input type="hidden" name="description" value="<%=description%>">
		<input type="hidden" name="hsn" value="<%=hsn%>">
		<input type="hidden" name="client_sn" value="<%if request("client") <> "" then response.write request("client") else response.write Session("mh_usersn") end if%>">
		<input type="hidden" name="Session_Sn" value="<%=Session_Sn%>">
		<input type="hidden" name="isanser" value="<%=isanser%>">
		<input type="hidden" name="vocsn" value="<%=vocsn%>">
		</form>
	  </center>
	</div>
	<%
		if(isanser <> "") then
			response.write "<script language='javascript'>"
			response.write "WebForm.submit();"
			response.write "</script>"
		end if
	%>
	</body>
	</html>
	<%
	conn.close
	else
		response.write "<script language='javascript'>"
		response.write "alert('您已經完成本次的家庭作業!');"
		response.write "window.close();"
		response.write "</script>"
	end if
	rs.close
	conn.close
End IF
%>
