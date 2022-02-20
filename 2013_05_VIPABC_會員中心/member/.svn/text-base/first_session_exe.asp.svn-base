<!--#include virtual="/lib/include/global.inc" -->
<% 
dim class_time

'如果前面傳空的時間到這程式將導回
class_time=trim(request("chk"))

if class_time = "" then
%>
     <script type="text/javascript">
         alert("请选择时间");			
				    history.go(-1)  				
	 </script>
<%
response.end
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
                alert("你已经预约过了请勿重复预约,如有问题请洽客服");
                location.href = 'member.asp';	
            </script>

            <%
        

end if        






cust_name=request("cust_name")&" "&request("cust_lame")
dim cust_name
dim date_year
dim date_month
dim date_day
dim date_hour
dim date_min
dim class_date
dim YMD

date_year=Mid(class_time,1,4)
date_month=Mid(class_time,6,2)
date_day=Mid(class_time,9,2)
date_hour=Mid(class_time,12,2)
date_min=Mid(class_time,15,2)

class_date=date_year&"/"&date_month&"/"&date_day

YMD=class_date



         dim  var_arr
         dim  str_client_datetime_sql
         dim  arr_result
        var_arr = Array(g_var_client_sn)
        str_client_datetime_sql = "select * from add_client_datetime where client_sn=@g_var_client_sn and flag=4 and testflag=1 " 
        arr_result = excuteSqlStatementRead(str_client_datetime_sql, var_arr, CONST_VIPABC_RW_CONN)       
        if (isSelectQuerySuccess(arr_result)) then
         if (Ubound(arr_result) > 0) then
            '有資料
            %>
            <script type="text/javascript">
                alert("你已经预约过了请勿重复预约,如有问题请洽客服");
                location.href = 'member.asp';	
            </script>

            <%
          end if  

        end if        



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


       dim count_pople
       dim str_first_sesion_time_sql
       dim min1
       dim anowtime
       dim count_people
       var_arr = Array(YMD,date_hour,date_min)
       str_first_sesion_time_sql = "select COUNT(*) as count_peo from add_client_datetime where YMD=@YMD and hour1=@date_hour and min1=@date_min AND (flag =4)" 
       arr_result = excuteSqlStatementRead(str_first_sesion_time_sql, var_arr, CONST_VIPABC_RW_CONN)       
       if (isSelectQuerySuccess(arr_result)) then
                    
           count_pople=arr_result(0,0)

       end if
		   
		'========== START ==========
		'CS11070604 修改預約測試管理,調整測試人力的排班時間  Freeman 20121126
		'測試人力皆由cs_manpower_schedule這張表紀錄，抓出預定人數與測試人力比對
		'有餘額則開放客戶預訂
		Dim int_count_manpower
		hour_min = date_hour &":"& date_min
		var_arr = Array(date_year,date_month,hour_min)
		sql = "SELECT d_"& date_day &" AS count_manpower FROM cs_manpower_schedule WHERE year=@year AND month=@month AND time=@time"					
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
		end if
		objPCInfoInsert.close
		Set objPCInfoInsert = nothing
		
		if int_count_manpower=0 or (int(int_count_manpower)<=int(count_people)) then
			response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');location.href=('first_session.asp');</script>"
			response.End
		elseif int_count_manpower>count_people then

		end if
		'=========== END ===========
					
     

       'if  trim(Weekday(YMD))<>1 and trim(Weekday(YMD))<>7 then   
       '           min1= date_min
       '
       '           if min1="0" then
       '           min1="00"
       '           end if
       '           anowtime=(date_hour&""&min1)                              
       '
       '        if  anowtime < 1400 then
       '        
       '            if count_pople >= af1 then
       '                response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');history.go(-1);</script>"
       '                response.End
       '            end if
       '        
       '        end if
       '        
       '        if anowtime >= 1400 and anowtime < 1500 then
       '        
       '            if count_pople >= af2 then
       '                response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');history.go(-1);</script>"
       '                response.End
       '            end if
       '        
       '        end if
       '        
       '        if anowtime >= 1500 and anowtime =< 1700 then
       '        
       '            if count_pople >= af3 then
       '                response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');history.go(-1);</script>"
       '                response.End
       '            end if
       '        
       '        end if
       '
       '        if anowtime >= 1730 and anowtime =< 1800 then
       '        
       '            if count_pople >= af4 then
       '                response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');history.go(-1);</script>"
       '                response.End
       '            end if
       '        
       '        end if
       '        
       '        if anowtime >= 1830 and anowtime =< 2200 then
       '        
       '            if count_pople >= af5 then
       '                response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');history.go(-1);</script>"
       '                response.End
       '            end if
       '        
       '        end if
       '
       '        if anowtime >= 2230  then
       '        
       '            if count_pople >= af6 then
       '                response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');history.go(-1);</script>"
       '                response.End
       '            end if
       '        
       '        end if
       '
       '
       '
       'else
       '        '禮拜六~禮拜日
       '        if Cint(date_hour)> 13 and Cint(date_hour)< 22 then
       '       
       '            if count_pople >= af7 then
       '                response.write "<script language='javascript1.2'>alert('该时段测试人数超过上限!');history.go(-1);</script>"
       '                response.End
       '            end if
       '        
       '        end if
       '
       '
       'end if
        





       dim str_client_basic_sql
       dim email
       dim call_phone
       dim client_name
       dim service_staff_sn
       var_arr = Array(g_var_client_sn)
       str_client_basic_sql = "SELECT a.email,a.cphone_code,a.cphone,a.cname,b.staff_sn FROM  client_basic a,client_service_staff b where a.sn=@g_var_client_sn and a.service_code=b.service_code" 
       'response.write(str_first_sesion_time_sql)
       arr_result = excuteSqlStatementRead(str_client_basic_sql, var_arr, CONST_VIPABC_RW_CONN)       
       if (isSelectQuerySuccess(arr_result)) then
                    
        Email=arr_result(0,0)
        call_phone=arr_result(1,0)&arr_result(2,0)
        client_name=arr_result(3,0)
        service_staff_sn=arr_result(4,0)

       end if


dim class_time1
class_time1=class_date&" "&date_hour&":"&date_min

str_description = "[客戶自約首次測試時間]預約日期:" & class_time1 & "。" &_
                  "<br/><font color=red>" & client_name & "</font>&nbsp;預約於 " & getFormatDateTime(now, 2) & "<hr/>"

 dim str_tablecol
 dim str_tablevalue
 str_tablecol = "client_sn,YMD,hour1,min1,staff_sn,flag,testflag,description"
 str_tablevalue = "@client_sn,@class_date,@date_hour,@date_min,@service_staff_sn,@flag,@testflag,@description"
 var_arr = Array(g_var_client_sn,class_date,date_hour,date_min,service_staff_sn,4,1,str_description)
 arr_result = excuteSqlStatementScalar("Insert into add_client_datetime ("&str_tablecol&") values("&str_tablevalue&") Select SCOPE_IDENTITY() " , var_arr, CONST_VIPABC_RW_CONN)

      



str_insert_client_customer_care_item = "client_sn, agent_sn, agent, comment, types, acd_sn"
str_insert_client_customer_care_value = "@client_sn, @agent_sn, @agent, @comment, @types, @acd_sn"
arr_client_customer_care_value = Array(g_var_client_sn, _
                                        service_staff_sn, _
                                        "C," & service_staff_sn, _
                                        str_cont1 & str_description, _
                                        "22", _
                                        arr_result(0))

str_sql = " INSERT INTO client_customer_care (" & str_insert_client_customer_care_item & ") " &_
          " VALUES (" & str_insert_client_customer_care_value &") "
            'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
int_write_result = excuteSqlStatementWrite(str_sql, arr_client_customer_care_value, CONST_VIPABC_RW_CONN)


'抓add_client_datetime最大值
       dim  str_client_datetime_sn_sql
       dim  Getsn
       str_client_datetime_sn_sql = "SELECT TOP (1) sn FROM  add_client_datetime ORDER BY sn DESC" 
       'response.write(str_first_sesion_time_sql)
       arr_result = excuteSqlStatementRead(str_client_datetime_sn_sql, "", CONST_VIPABC_RW_CONN)       
       if (isSelectQuerySuccess(arr_result)) then
            
        Getsn=arr_result(0,0)   

       end if


     dim str_send_email_body_url
     dim str_class_email_subject

	str_send_email_body_url = "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/mail_template/first_session/first_session_email.asp?cust_name="&escape(cust_name)&"&class_time="&escape(class_time1)&""							
    str_class_email_subject = "感谢您预约VIPABC 「First Session」测试"
    dim int_send_email_result
    
    int_send_email_result = sendEmailFromUrl(Email, cust_name, str_class_email_subject, str_send_email_body_url, CONST_VIPABC_WEBSITE)

    ' -----開始寄信給CS告知有客戶要做第一次預約測試-----START
    
    dim str_staff_sn_cs_sql
    dim p_var_arr_opt_paras
	dim exec_result		     
    dim str_send_email_staff_body_url    
    dim arr_set_paras
    dim arr_opt_paras
    dim str_mail_subjcet
    dim arr_email_to_paras
    dim arr_email_from_paras
    dim int_count

    str_staff_sn_cs_sql = "select fname,fname+lname+'@tutorabc.com' as email from staff_basic where dept='6' and status='1' AND position<'52'" 
    'str_staff_sn_cs_sql = "select fname,fname+lname+'@tutorabc.com' as email from staff_basic where sn=481 or sn=460" 
    arr_result = excuteSqlStatementRead(str_staff_sn_cs_sql, "", CONST_VIPABC_RW_CONN)       
    if (isSelectQuerySuccess(arr_result)) then

          '有資料

      if (Ubound(arr_result) >= 0) then
      
        for int_count = 0  to Ubound(arr_result,2)
		  str_send_email_staff_body_url = "http://"&CONST_VIPABC_WEBHOST_NAME&"/program/mail_template/first_session/first_session_staff_email.asp?staff_name="&escape(arr_result(0,int_count))&"&client_name="&escape(client_name)&"&client_phone="&escape(call_phone)&"&client_email="&escape(Email)&"&class_time="&escape(class_time1)&""							
		  str_mail_subjcet = "客戶("&client_name&")預約測試時間："&class_time1
		  arr_email_from_paras 	= Array(CONST_MAIL_ADDRESS_SERVICES_ABC,"",str_mail_subjcet,"","",str_send_email_staff_body_url)
		  arr_email_to_paras 	= Array(arr_result(1,int_count))
		  arr_set_paras 		= Array(CONST_MAIL_SERVER_CLIENT, CONST_EMAIL_CHARSET)
		  arr_opt_paras	        = Array(false)
            				
		  exec_result = sendEmail(arr_email_from_paras, arr_email_to_paras, arr_set_paras, arr_opt_paras, CONST_TUTORABC_WEBSITE)

         next
          
      end if
   end if   
   
   ' -----開始寄信給CS告知有客戶要做第一次預約測試-----END
		'後來就變改成 要預約完才能訂課 Johnny
		session("bolFirstTestQuestionaryhaveTestNOTOkOrder") = "Y"  '與  E:\web\www.vipabc.com\program\member\reservation_class\normal_and_vip.asp 和 E:\web\www.vipabc.com\program\member\reservation_class\include\include_prepare_and_check_data.inc 連動相關
		Call alertGo("感谢您的预约！ \n客服人员将于【"&class_time1&"】致电给您进行电脑设备测试，有任何问题亦可洽询客服专线："& CONST_SERVICE_PHONE2 &"\n按下「确定」钮后，系统将导至「DCGS学习偏好设定」。", "/program/member/learning_dcgs.asp", CONST_ALERT_THEN_REDIRECT)
        'Call alertGo("客服人员将于该时段致电，如欲更改测试时间\n请洽客服专线："& CONST_SERVICE_PHONE2 &"，谢谢", "/program/member/reservation_class/normal_and_vip.asp", CONST_ALERT_THEN_REDIRECT)
%>


