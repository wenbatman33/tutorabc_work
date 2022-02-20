<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
if isRegisterMember() = false then
            %>
            <script type="text/javascript">
                alert("尚未登入");
                location.href = 'member_login.asp';	
            </script>

            <%    
end if



dim str_check_purchase_sql
dim now_date
        
now_date=dateValue(now())
var_arr = Array(g_var_client_sn,now_date)
str_check_purchase_sql = "SELECT * FROM client_purchase where  client_sn=@g_var_client_sn and product_edate>@now_date and valid=1 " 
arr_result = excuteSqlStatementRead(str_check_purchase_sql, var_arr, CONST_VIPABC_RW_CONN)       
if (isSelectQuerySuccess(arr_result) = false ) then
        
            '有資料
            %>
            <script type="text/javascript">
                alert("你无法使用此页");
                location.href = 'member.asp';	
            </script>

            <%      

end if        


dim str_check_client_datetime_sql
var_arr = Array(g_var_client_sn)
str_check_client_datetime_sql = "select * from add_client_datetime where client_sn=@g_var_client_sn and flag=4 and testflag=1" 
arr_result = excuteSqlStatementRead(str_check_client_datetime_sql, var_arr, CONST_VIPABC_RW_CONN)       
'response.write g_str_sql_statement_for_debug
if (isSelectQuerySuccess(arr_result)) then
    if (Ubound(arr_result)>=0) then
          Call alertGo("你已经预约过了请勿重复预约,如有问题请洽客服 : "& CONST_SERVICE_PHONE2 &"，谢谢", "/program/member/reservation_class/normal_and_vip.asp", CONST_ALERT_THEN_REDIRECT)
    end if
      

end if        



dim nowtime
dim nowweek                
dim nowhour
dim nowMinute
nowweek=Weekday(now())
nowhour=(Hour(now()) + 1)
nowMinute=Minute(now())
nowtime=nowhour&":"&nowMinute


dim str_client_basic_sql
dim var_arr
dim arr_result

var_arr = Array(g_var_client_sn)
str_client_basic_sql = "SELECT fname,lname FROM client_basic where sn=@g_var_client_sn" 
arr_result = excuteSqlStatementRead(str_client_basic_sql, var_arr, CONST_VIPABC_RW_CONN)       
if (isSelectQuerySuccess(arr_result)) then
    if (Ubound(arr_result) >= 0) then
        '有資料    
            
        cust_name=arr_result(0,0)
        cust_lame=arr_result(1,0)
    end if           
end if        



dim weekname(6)

weekname(0)="星期日"  '星期日
weekname(1)="星期一"  '星期一
weekname(2)="星期二"  '星期二
weekname(3)="星期三"  '星期三
weekname(4)="星期四"  '星期四
weekname(5)="星期五"  '星期五
weekname(6)="星期六"  '星期六

dim opentime(28)
opentime(0) = "09:00"
opentime(1) = "09:30"
opentime(2) = "10:00"
opentime(3) = "10:30"
opentime(4) = "11:00"
opentime(5) = "11:30"
opentime(6) = "12:00"
opentime(7) = "12:30"
opentime(8) = "13:00"
opentime(9) = "13:30"
opentime(10) = "14:00"
opentime(11) = "14:30"
opentime(12) = "15:00"
opentime(13) = "15:30"
opentime(14) = "16:00"
opentime(15) = "16:30"
opentime(16) = "17:00"
opentime(17) = "17:30"
opentime(18) = "18:00"
opentime(19) = "18:30"
opentime(20) = "19:00"
opentime(21) = "19:30"
opentime(22) = "20:00"
opentime(23) = "20:30"
opentime(24) = "21:00"
opentime(25) = "21:30"
opentime(26) = "22:00"
opentime(27) = "22:30"
opentime(28) = "23:00"

dim af1
dim af2
dim af3
dim af4
dim af5
dim af6
dim af7
dim g_var_first_sesion
dim str_first_sesion_sql



        var_arr = Array(g_var_first_sesion)
        str_first_sesion_sql = "SELECT * FROM customer_first_sesion where mod_type=1" 
        arr_result = excuteSqlStatementRead(str_first_sesion_sql, var_arr, CONST_VIPABC_RW_CONN)       
        if (isSelectQuerySuccess(arr_result)) then
            '有資料

          
            
            af1=int(arr_result(1,0))
            af2=int(arr_result(2,0))
            af3=int(arr_result(3,0))
            af4=int(arr_result(4,0))
            af5=int(arr_result(5,0))
            af6=int(arr_result(6,0))
            af7=int(arr_result(7,0))
           
        end if          


%>
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery.timers.js"></script>
<script type="text/javascript" language="javascript">

    //chooseOne()函式，參數為觸發該函式的元素本身   
    function chooseOne(cb) {
        //先取得同name的chekcBox的集合物件   
        var obj = document.getElementsByName("chk");
        for (i = 0; i < obj.length; i++) {
            //判斷obj集合中的i元素是否為cb，若否則表示未被點選   
            if (obj[i] != cb) obj[i].checked = false;
            //若是 但原先未被勾選 則變成勾選；反之 則變為未勾選   
            //else  obj[i].checked = cb.checked;   
            //若要至少勾選一個的話，則把上面那行else拿掉，換用下面那行   
            else obj[i].checked = true;
        }
    }
    function chkdata() {
        var len = document.fm1.chk.length;
        var class_time = "";
        for (i = 0; i < len; i++) {

            if (document.fm1.chk[i].checked == true) {
				
                class_time = document.fm1.chk[i].value;

            }
        }

        if (confirm("<%' 提醒您！建议您完成预约测试后再体验第一堂课，所以您预约订位的时间需于测试时间 %>提醒您！建议您完成预约测试后再体验第一堂课，所以您预约订位的时间需于测试时间" + class_time + "<% '之后。如需修改请按「取消」或按「确定」完成First Session测试预约。 %>之后。如需修改请按「取消」或按「确定」完成First Session测试预约。")) {
            document.fm1.Submit.disabled = true;
            document.fm1.submit();
        }
        else {
            //alert('asdf');

            //return false;

        }
    }

</script>

<!--內容start-->
       <div id="menber_membox">
                
                <p align="left"><font face="新細明體" size="2" color="#000080">
                
				「</font><b><font face="Arial Unicode MS" color="#000080" size="2">First 
				Session</font></b><font face="新細明體" size="2" color="#000080">」<b><% '通訊品質測試 %>通讯质量测试</b></font>
				<p align="left"><font face="新細明體" color="#000080" size="2"> 
				<% '親愛的 %>亲爱的</font><font face="Arial Unicode MS" color="#000080" size="2"><%=cust_name %></font><font face="新細明體" size="2" color="#000080">，  
			<% '歡迎您加入 %>欢迎您加入</font><font face="Arial Unicode MS" color="#000080" size="2">VIPABC</font><font size="2" face="新細明體" color="#000080"> 
				! 
				<% '請勾選預約測試的時間，以確保您上課的品質。 %>请勾选预约测试的时间，以确保您上课的质量。</font></p></p>

      <table width="570" border="0" align="center" cellpadding="5" cellspacing="1" bordercolor="#000000" bgcolor="#000000" style="border-right: #000000 1px solid; border-top: #000000 1px solid; border-left: #000000 1px solid; border-bottom: #000000 1px solid;" >
                <tbody>
                <form name=fm1 method="post" action="first_session_exe.asp">
                  <tr>
                    <td width="65"  class="style0" bgcolor="#8E8E8E" class="fonteleven"><div align="center"></div></td>
                  <%
                   dim i
                   dim orderdate
                   dim syear
                   dim smonth
                   dim sday
                   dim weeks
                   dim symd
                   dim j
                   dim holidayD
                   dim findH
                   dim YMD
                   dim date_hour
                   dim date_min
                  for i = 0 to 6
                   orderdate = dateadd("d",i,dateadd("d",0,date))
				   syear = left(year(orderdate),4)		'22090604 joey mod by RD09052709			   
                   smonth = right("0"&month(orderdate),2)
                   sday    = right("0"&day(orderdate),2)
                   weeks     = weekday(orderdate)
				   symd=syear+smonth+sday				'22090604 joey mod by RD09052709	
                  %>
                    <td width="65"  class="style0" bgcolor="#8E8E8E" class="fonteleven"><div align="center"><font size=2 color="ffffff"><%=smonth%>/<%=sday%><br />
                    (<%=weekname(weeks-1)%>)</font></div></td>
                    <%next%>
                    
                  </tr>
                  
                  <tr>
                  <%for j = 0 to 26
                  if j=0 or j=1 or j=4 or j=5 or j=8 or j=9 or j=12 or j=13 or j=16 or j=17 or j=20 or j=21 or j=24 or j=25 or j=28 then
                  %>
                    <td height="30" class="style2" align="middle" bgcolor="#87CEFA" ><div align="center"><font size=2><%=opentime(j)%></font></div></td> 
                  <%else%>
                     <td height="30" class="style5" align="middle" bgcolor="#D4D4D4" ><div align="center"><font size=2><%=opentime(j)%></font></div></td> 
                    <!--tstr是上課時段-->
                  <%
                   end if
                                    
                   for i =0 to 6
                   orderdate = dateadd("d",i,dateadd("d",0,date))
				   syear = left(year(orderdate),4)				    '22090604 joey mod by RD09052709	
                   smonth = right("0"&month(orderdate),2)
                   sday    = right("0"&day(orderdate),2)
                   weeks     = weekday(orderdate)
				   symd=syear+smonth+sday							'22090604 joey mod by RD09052709	
					'response.write syear & "	" & smonth & "	" & sday & "<BR>"
					
                   findH=InStr(holidayD,symd)    '22090610 joey mod by RD09052709
                   
                   '判斷此堂課預約人數
                   YMD=year(orderdate)&"/"&smonth&"/"&sday
                   date_hour=Mid(opentime(j),1,2)  
                   date_min=Mid(opentime(j),4,5)
				   'response.write YMD & "	" & date_hour & "	" & date_min & "<BR>"

                   dim str_first_sesion_time_sql
                   dim count_people
                   dim open
                   dim language
                   dim cust_name
                   dim cust_lame
     
                    var_arr = Array(YMD,date_hour,date_min)
                    str_first_sesion_time_sql = "select COUNT(*) as count_peo from add_client_datetime where YMD=@YMD and hour1=@date_hour and min1=@date_min AND (flag =4)" 
                   'response.write(str_first_sesion_time_sql)
                    arr_result = excuteSqlStatementRead(str_first_sesion_time_sql, var_arr, CONST_VIPABC_RW_CONN)       
                    if (isSelectQuerySuccess(arr_result)) then
                    
						count_people=arr_result(0,0)

                    end if
        
					'========== START ==========
					'CS11070604 修改預約測試管理,調整測試人力的排班時間  Freeman 20121126
					'測試人力皆由cs_manpower_schedule這張表紀錄，抓出預定人數與測試人力比對Dim int_count_manpower
					'有餘額則開放客戶預訂
					hour_min = date_hour &":"& date_min
					var_arr = Array(syear,smonth,hour_min)
					sql = "SELECT d_"& day(orderdate) &" AS count_manpower FROM cs_manpower_schedule WHERE year=@year AND month=@month AND time=@time"					
					Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")    '建立連線
					intSqlWriteResult = objPCInfoInsert.excuteSqlStatementEach(sql,var_arr,CONST_VIPABC_RW_CONN) 	'無傳入參數用
					if(not objPCInfoInsert.eof) then 'SQL成功
						if objPCInfoInsert.recordcount=1 then	
							if isnull(objPCInfoInsert("count_manpower")) then
								int_count_manpower = 0
							else
								int_count_manpower = objPCInfoInsert("count_manpower")
							end if
						else
							int_count_manpower=0
						end if 
					else
						int_count_manpower=0
					end if
					objPCInfoInsert.close
					Set objPCInfoInsert = nothing
					
					int_count_manpower = CInt(int_count_manpower)
					count_people = CInt(count_people)
					if (int_count_manpower=0) or (int_count_manpower <= count_people) then
						open="N"
					elseif int_count_manpower > count_people then
						open="Y"
					end if
					'=========== END ===========
					
                   
                   
                    '判斷時段幾個人就額滿
                    'if weeks=7 and j<9 then
                    ' open="N"
					' elseif  weeks=1 and j<9 then
					' open="N"
					' elseif  weeks=2 and count_people >= af1 and j=<9 then
					' open="O"
					' elseif  weeks=2 and count_people >= af2 and j>=10 and j<=11 then
					' open="O"
					' elseif  weeks=2 and count_people >= af3 and j>=12 and j<=16 then
					' open="O"
					' elseif  weeks=2 and count_people >= af4 and j>=17 and j<=18 then
					' open="O"
					' elseif  weeks=2 and count_people >= af5 and j>=19 and j<=26 then
					' open="O"
					' elseif  weeks=2 and count_people >= af6 and j>=27 then
					' open="O"
					' elseif  weeks=3 and count_people >= af1 and j=<9 then
					' open="O"
					' elseif  weeks=3 and count_people >= af2 and j>=10 and j<=11 then
					' open="O"
					' elseif  weeks=3 and count_people >= af3 and j>=12 and j<=16 then
					' open="O"
					' elseif  weeks=3 and count_people >= af4 and j>=17 and j<=18 then
					' open="O"
					' elseif  weeks=3 and count_people >= af5 and j>=19 and j<=26 then
					' open="O"
					' elseif  weeks=3 and count_people >= af6 and j>=27 then
					' open="O"
					' elseif  weeks=4 and count_people >= af1 and j=<9 then
					' open="O"
					' elseif  weeks=4 and count_people >= af2 and j>=10 and j<=11 then
					' open="O"
					' elseif  weeks=4 and count_people >= af3 and j>=12 and j<=16 then
					' open="O"
					' elseif  weeks=4 and count_people >= af4 and j>=17 and j<=18 then
					' open="O"
					' elseif  weeks=4 and count_people >= af5 and j>=19 and j<=26 then
					' open="O"
					' elseif  weeks=4 and count_people >= af6 and j>=27 then
					' open="O"
					' elseif  weeks=5 and count_people >= af1 and j=<9 then
					' open="O"
					' elseif  weeks=5 and count_people >= af2 and j>=10 and j<=11 then
					' open="O"
					' elseif  weeks=5 and count_people >= af3 and j>=12 and j<=16 then
					' open="O"
					' elseif  weeks=5 and count_people >= af4 and j>=17 and j<=18 then
					' open="O"
					' elseif  weeks=5 and count_people >= af5 and j>=19 and j<=26 then
					' open="O"
					' elseif  weeks=5 and count_people >= af6 and j>=27 then
					' open="O"
					' elseif  weeks=6 and count_people >= af1 and j=<9 then
					' open="O"
					' elseif  weeks=6 and count_people >= af2 and j>=10 and j<=11 then
					' open="O"
					' elseif  weeks=6 and count_people >= af3 and j>=12 and j<=16 then
					' open="O"
					' elseif  weeks=6 and count_people >= af4 and j>=17 and j<=18 then
					' open="O"
					' elseif  weeks=6 and count_people >= af5 and j>=19 and j<=26 then
					' open="O"
					' elseif  weeks=6 and count_people >= af6 and j>=27 then
					' open="O"
					' elseif  weeks=7 and count_people >= af7 and j>=9 then
					' open="O"
					' elseif  weeks=1 and count_people >= af7 and j>=9 then
					' open="O"
					' elseif  weeks=7 and  j>26 then
					' open="N"
					' elseif  weeks=1 and  j>26 then
					' open="N"		
					' elseif YMD="2008/06/08"  or YMD="2008/09/29" or YMD="2008/10/10" or YMD="2008/12/25" or YMD="2009/05/28"  or YMD="2009/08/07" or YMD="2009/08/08" then
					' open="N"	
					' elseif  opentime(j) < nowtime and weeks=nowweek then	 
					' 					 
                    ' open="N" 
					'else
                    ' open="Y"
					'end if
                   
				   
'22090604 joey mod by RD09052709 start		
                   	 findH=InStr(holidayD,symd)  
					 if findH<>0 then
						findH1=findH+9
						findHa=InStr(findH1,holidayD,"*")
						findHb=Mid(holidayD, findH1, findHa-findH1)
						if left(opentime(j),1)=0 then
							gh=", "& Mid(opentime(j),2,1) &","
						else
							gh=", "&Mid(opentime(j),1,2) &","
						end if
						findHc=InStr(findHb,gh)
						if findHc>0 then
							open="O"
						end if
					end if
'end
				   
				   
				   
            if  open="Y" then   
                if j=0 or j=1 or j=4 or j=5 or j=8 or j=9 or j=12 or j=13 or j=16 or j=17 or j=20 or j=21 or j=24 or j=25 or j=28 then
                %>
					<td align="middle" class="style2" title="" bgcolor="#87CEFA"><div align="center" >
                <% else %>
					<td align="middle" class="style5" title="" bgcolor="#D4D4D4"><div align="center" >
                <% end if  %>
            <% else %>
               <td align="middle" class="style10" title="" bgcolor="#D4D4D4"><div align="center" >
              <%     
              end if         
			  
			  
                   
                    if open="Y" then %>
						<input type="checkbox" value="<%=year(orderdate) %>/<%=smonth %>/<%=sday %>-<%=opentime(j)%>" name="chk" onClick="chooseOne(this)" >
						<br>
					<%elseif open="O" then%>
						<font size=2 color="#ff0000" class="block" ></font>
					<%end if %>
                   </div></td>
                   <%next %>
                    
                   
                  </tr>
                  <%next%>
                  
                </tbody>
           
                 <input type="hidden" name="language" value=<%=language%>>
                  <input type="hidden" name="aurl" value=<%=request("aurl")%>>
                  <input type="hidden" name="cust_name" value=<%=cust_name%>>
                  <input type="hidden" name="cust_lame" value=<%=cust_lame%>>
                  </table>
                     <p align="center"><input type="button" name="Submit" value="確定預約" onclick='chkdata();'></p></form>   
         <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
       </div>
       <!--內容end-->
       
       
       
  
       
