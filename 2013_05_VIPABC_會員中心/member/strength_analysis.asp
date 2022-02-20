<% 
client_sn = request.cookies("ming_csn")

'========取得適性測驗的LEVEL==================
GC = 0
VC = 0
PC = 0
LC = 0
RC = 0
SC = 0

sql10="select top 6  * from placement_test_record where client_sn ="&client_sn&" order by test_sn "
set rs10=server.createobject("adodb.recordset")
rs10.open sql10,conn,1,1
if rs10.EOF=false then
do until rs10.eof

if rs10("evaluation_type")=1 then 
    GC=rs10("level_mean") '文法
end if

if rs10("evaluation_type")=2 then 
    VC=rs10("level_mean") '字彙
end if
if rs10("evaluation_type")=3 then 
    PC=rs10("level_mean") '發音
end if

if rs10("evaluation_type")=4 then 
    LC=rs10("level_mean") '聽力
end if
if rs10("evaluation_type")=5 then 
    RC=rs10("level_mean") '閱讀
end if
    
if rs10("evaluation_type")=6 then 
    SC=rs10("level_mean") '會話
end if    
rs10.Movenext
loop
end if


'=====取得客戶上課後的Level
ClientAllLevel client_sn,GL,VL,PL,LL,RL,SL

sql_booking="select top 1  * from customer_booking_record where client_sn ="&client_sn&" AND stype<>'3' order by sdate desc "
set rs_booking=server.createobject("adodb.recordset")
rs_booking.open sql_booking,conn,1,1
if not rs_booking.eof then

Cbooking="YES"
else
Cbooking="NO"
end if
rs_booking.close
conn.close
%>
<div>

		<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table5">
	    <tr>
			<td height="20"></td>
		</tr>		
		  <tr>
			<td rowspan="11" align="left" valign="top">
			  <table width="560" border="0" align="center" cellpadding="0" cellspacing="0" class="system_table">
				<tr>
				  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="45%"><img src="../../images/tital/Strengths_Analysis_<%= lang %>.gif" /></td>
                      <td width="55%" style="background-image:url(../../images/title_bg.gif); background-position:center; background-repeat:repeat-x;">&nbsp;</td>
                    </tr>
                  </table></td>
				</tr>
				<tr>
				  <td ></td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
				<tr>
				  <td class="small_font">
					<table border="0" cellspacing="1" id="table2"  align="center">
						<% If Cbooking="YES" Then %>
                        
                        <tr>
							<td width="48%" colspan="2" style="padding-bottom: 5px" align="center" class="middle_font"><b><%=char00140 %><%'Begin Strengths Analysis %></b></td>
							<td width="4" style="padding-bottom: 5px"></td>
							<td width="48%" colspan="2" style="padding-bottom: 5px" align="center" class="middle_font"><b><%=char00142 %><%'Current Session Focus %></b></td>
						</tr>
						<tr>
							<td width="15%" class="small_font">
							<font color="#444444"><%=char00143 %><%'Grammar %></font></td>
							<td width="33%" align="center"><img border="0" src="../../images/strength/focus<%=GC %>.gif" width="160" height="20"></td>
							<td width="4">&nbsp;</td>
							<td width="15%" class="small_font">
							<font color="#444444"><%=char00143 %><%'Grammar %></font></td>
							<td width="33%" align="center">
							<img border="0" src="../../images/strength/-focus<%=GL %>.gif" width="160" height="20"></td>
						</tr>
						<tr>
							<td width="15%" style="padding-top: 3px; padding-bottom: 5px" class="small_font">
							<font color="#444444"><%=char00144 %><%'Vocabulary %></font></td>
							<td width="33%" align="center">
							<img border="0" src="../../images/strength/focus<%=VC %>.gif" width="160" height="20"></td>
							<td width="4">&nbsp;</td>
							<td width="15%" class="small_font">
							<font color="#444444"><%=char00144 %><%'Vocabulary %></font></td>
							<td width="33%" align="center">
							<img border="0" src="../../images/strength/-focus<%=VL %>.gif" width="160" height="20"></td>
						</tr>
						<tr>
							<td width="15%" style="padding-top: 3px; padding-bottom: 5px" class="small_font">
							<font color="#444444"><%=char00145 %><%'Pronunciation %></font></td>
							<td width="30%" align="center">
							<img border="0" src="../../images/strength/focus<%=PC %>.gif" width="160" height="20"></td>
							<td width="4">&nbsp;</td>
							<td width="15%" class="small_font">
							<font color="#444444"><%=char00145 %><%'Pronunciation %></font></td>
							<td width="30%" align="center">
							<img border="0" src="../../images/strength/-focus<%=PL %>.gif" width="160" height="20"></td>
						</tr>
						<tr>
							<td width="15%" style="padding-top: 3px; padding-bottom: 5px" class="small_font">
							<font color="#444444"><%=char00146 %><%'Speaking %></font></td>
							<td width="30%" align="center">
							<img border="0" src="../../images/strength/focus<%=SC %>.gif" width="160" height="20"></td>
							<td width="4">&nbsp;</td>
							<td width="15%" class="small_font">
							<font color="#444444"><%=char00146 %><%'Speaking %></font></td>
							<td width="30%" align="center">
							<img border="0" src="../../images/strength/-focus<%=SL %>.gif" width="160" height="20"></td>
						</tr>
						<tr>
							<td width="15%" style="padding-top: 3px; padding-bottom: 5px" class="small_font">
							<font color="#444444"><%=char00147 %><%'Reading %></font></td>
							<td width="30%" align="center">
							<img border="0" src="../../images/strength/focus<%=RC %>.gif" width="160" height="20"></td>
							<td width="4">&nbsp;</td>
							<td width="15%" class="small_font">
							<font color="#444444"><%=char00147 %><%'Reading %></font></td>
							<td width="30%" align="center">
							<img border="0" src="../../images/strength/-focus<%=RL %>.gif" width="160" height="20"></td>
						</tr>
						<tr>
							<td width="15%" style="padding-top: 3px; padding-bottom: 5px" class="small_font">
							<font color="#444444"><%=char00148 %><%'Listening %></font></td>
							<td width="30%" align="center"><img border="0" src="../../images/strength/focus<%=LC %>.gif" width="160" height="20"></td>
							<td width="4">&nbsp;</td>
							<td width="15%" class="small_font">
							<font color="#444444"><%=char00148 %><%'Listening %></font></td>
							<td width="30%" align="center"><img border="0" src="../../images/strength/-focus<%=LL %>.gif" width="160" height="20"></td>
						</tr>
                        <tr>
						    <td  colspan="5"  align="center">
						        
						        <br />
						        <input type="button"  value="<%=strMember00045%>"  style="font-size: larger; font-weight:bolder" onClick="return window.location.href='strength_analysis_exe.asp';"/>
						    </td>
						</tr>
						<% Else %>
                        <tr>
							<td colspan="2" style="padding-bottom: 5px" align="center" class="middle_font"><b><%=char00140 %><%'Begin Strengths Analysis %></b></td>
							
						</tr>
						<tr>
							<td width="30%" align="right" class="small_font" >
							<font color="#444444"><%=char00143 %><%'Grammar %></font></td>
							<td width="70%" align="center"><img border="0" src="../../images/strength/focus<%=GC %>.gif" width="160" height="20"></td>
						</tr>
						<tr>
							<td width="30%" align="right" style="padding-top: 3px; padding-bottom: 5px;" class="small_font">
							<font color="#444444"><%=char00144 %><%'Vocabulary %></font></td>
							<td width="70%" align="center">
							<img border="0" src="../../images/strength/focus<%=VC %>.gif" width="160" height="20"></td>
							
						</tr>
						<tr>
							<td width="30%" align="right" style="padding-top: 3px; padding-bottom: 5px;" class="small_font">
							<font color="#444444"><%=char00145 %><%'Pronunciation %></font></td>
							<td width="70%" align="center">
							<img border="0" src="../../images/strength/focus<%=PC %>.gif" width="160" height="20"></td>
							
						</tr>
						<tr>
							<td width="30%" align="right" style="padding-top: 3px; padding-bottom: 5px;" class="small_font">
							<font color="#444444"><%=char00146 %><%'Speaking %></font></td>
							<td width="70%" align="center">
							<img border="0" src="../../images/strength/focus<%=SC %>.gif" width="160" height="20"></td>
							
						</tr>
						<tr>
							<td width="30%" align="right" style="padding-top: 3px; padding-bottom: 5px;" class="small_font">
							<font color="#444444"><%=char00147 %><%'Reading %></font></td>
							<td width="70%" align="center">
							<img border="0" src="../../images/strength/focus<%=RC %>.gif" width="160" height="20"></td>
							
						</tr>
						<tr>
							<td width="30%" align="right" style="padding-top: 3px; padding-bottom: 5px;" class="small_font">
							<font color="#444444"><%=char00148 %><%'Listening %></font></td>
							<td width="70%" align="center"><img border="0" src="../../images/strength/focus<%=LC %>.gif" width="160" height="20"></td>
						</tr>
                       <tr>
						    <td  colspan="2"  align="center">
						        <br />
						        <input type="button"  value="<%=strMember00045%>"  style="font-size: larger; font-weight:bolder" onClick="return window.location.href='strength_analysis_exe.asp';"/>
						    </td>
						</tr> 
						<% End If %>
						
					</table>
					</td>
				</tr>
                </table>
			</td>
				</tr>					
			  </table>
<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>