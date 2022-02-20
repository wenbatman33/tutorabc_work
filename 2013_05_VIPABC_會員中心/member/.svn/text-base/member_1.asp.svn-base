<!--#include virtual="/lib/include/html/template/template_change.asp"-->

<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery.timers.js"></script>
<script type="text/javascript" language="javascript">

//function test(){
//   //do something...
//   //alert('aaa');
//}
//
//$('body').everyTime('1s',test);
//
////每1秒執行，並命名計時器名稱為A
//$('body').everyTime('1s','A',function(){
//	$('body').text();
//});





$(document).ready(function(){
	//先取得系統時間
	//特别注明：北京时间是格林尼治标准时加八小时，因此我用的起始时间也是从八点起算的  
   // var   secondServer   =   <%=DateDiff("s","2009-12-02 00:00:00", Now())%>;  
//	alert(secondServer); 
//	secondServer = secondServer / 60 ;
//	var a_sec = secondServer / 60 /60;
//	alert(secondServer); 
//	alert(a_sec);
	
	var server_hour = <%=hour(now())%>
	var server_minutes = <%=minute(now())%>
	var server_second = <%=second(now())%>
	//alert(server_hour); 
	//alert(server_minutes); 
	//alert(server_second); 
	
	
	
	
	function TimeView(sec){
		 //var showSecon = Math.floor(sec / 60 /60)  ;
		 
         var showTime = Math.floor(sec / 60) ; //分
		 var passTime = 180 - sec  //已經過幾秒
		 var dSecs = 0 ;
		 
		 if ( passTime > 60){
		 	passTime = passTime - 60;
		
		 }

		if (passTime < 60 ) {
			dSecs = 60 - passTime;
		 }else{
			dSecs = sec;
		}
  
		 
		$("input[name='t1']").val(showTime+":"+dSecs);
		setTimeout(function () {
     
			if (sec > 0) { 
				TimeView(sec-1);
			}
		}, 1000);
	}   
	TimeView(180)//設定秒數
})

</script>

<!--內容start-->
       <div id="menber_membox">
         <!--收藏課程start-->
         <div id="aaatest"><input type="button" name="t1" id="t1" value="" class="btn_1 m-left5"/></div>
           <%
			Dim int_client_sn : int_client_sn = session("client_sn")
			Dim str_client_attend_sql 
			Dim str_attend_sestime : str_attend_sestime = "" 
			Dim int_i 
			Dim str_today_session : str_today_session = "" '今日下堂課時間
			Dim str_other_session : str_other_session = "" '其他咨訊時段
			Dim arr_other_session : arr_other_session = "" '其他咨訊時段陣列
			
			Dim str_show_today	'顯示今日上課狀態
			Dim int_now			'現在的時間+3.5小時
			Dim int_session		'實際上課時間
			Dim bol_room_bout : bol_room_bout = 0 'false (預設沒有"進教室"鈕)
			'101632
			var_arr = Array("101632",getFormatDateTime(date(),5),hour(now()))
			str_client_attend_sql = "select Top 4 attend_date,attend_sestime from client_attend_list " 
			str_client_attend_sql = str_client_attend_sql & " where client_sn=@client_sn and attend_date>=@attend_date and attend_sestime >= @attend_sestime"   
			str_client_attend_sql = str_client_attend_sql & " order by attend_date,attend_sestime "
			
			
			arr_result = excuteSqlStatementRead(str_client_attend_sql, var_arr, CONST_VIPABC_RW_CONN)    
		   
		    'Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
			'Response.Write("錯誤訊息" & g_int_run_time_err_code & "<br>")
			'Response.Write("錯誤訊息sql....:" & g_str_sql_statement_for_debug & "<br>")
			  
          if (isSelectQuerySuccess(arr_result)) then
                       
            if (Ubound(arr_result) >= 0) then
				for int_i = 0  to Ubound(arr_result,2)
					if (CDate(arr_result(0,int_i)) = CDate(date()) and str_today_session = "" ) then 
						'找出今日下一堂課程時間
						str_today_session = arr_result(1,int_i)&":30"
					else
						'找出其他咨訊時間
						str_other_session = str_other_session & "@" &getFormatDateTime(arr_result(0,int_i),6) &"&nbsp;/&nbsp;"&arr_result(1,int_i)&":30"
					end if 
				next
				'整理其他咨訊時間的值
				if (str_other_session <> "") then 
					if left(str_other_session,1)="@" then str_other_session=right(str_other_session,len(str_other_session)-1)
					if right(str_other_session,1)="@" then str_other_session=left(str_other_session,len(str_other_session)-1)
					str_other_session = str_other_session & "@"
				end if 
				
				'response.write str_other_session & "<br>"
				'response.write str_today_session & "<br>"
           else
		   		response.write "ccccccccccccccc"

                '錯誤訊息
            '	Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
            ''	Response.End()
	            'call sub alertGo
	           'Call alertGo("資料錯誤，請洽客服人員", "", CONST_ALERT_THEN_GO_BACK)
	           'Response.End
            end if
        end if 
            %>
        <!--標題start-->
        <div class="main_mylist">
        
        <ul>
        <%
		
		
		if ( str_today_session <> "") then 
			'表示今日有課
			str_show_today = getWord("NEXT_CLASS_TODAY")&"："
			'--------------- 課程為3.5小時內，才開放"進入教室"按鈕 -------------- start -----------
			int_now = DateAdd("n" , 210 , now()) '現在的時間 加上210分(3.5小時)
			int_now = CDbl(Right("0"&hour(int_now),2)&Right("0"&minute(int_now),2))
			int_session = CDbl(Right("0"&hour(str_today_session),2)&Right("0"&minute(str_today_session),2)) '實際上課的時間
			if int_now >= int_session then '是否在開課前3.5小時內，即顯示進入教室按鈕
				bol_room_bout = 1
			end if 
			'--------------- 課程為3.5小時內，才開放"進入教室"按鈕 -------------- end -----------
		else
			str_show_today = getWord("NO_CLASS_TODAY")
		end if 
		%>
        
        <li><%=str_show_today%><span class="black"><%=str_today_session%></span>
        <%if ( bol_room_bout = 1) then %>
        	<span class="right_02"><input type="submit" name="button" id="button" value="+ <%=getWord("CLASS_3")%>" class="btn_1 m-left5"/></span>
        <% end if %>
        <span class="right_02"><input type="submit" name="button2" id="button2" value="+ <%=getWord("SHCEDULE_CLASS")%>" class="btn_1 m-left5"/></span>
        </li>
         
        <li><%=getWord("OTHER_CLASS")%>：
        <%
		if (str_other_session <> "") then 
			arr_other_session = split(str_other_session,"@")
			for int_i = 0 to ubound(arr_other_session)
				if (arr_other_session(int_i) <> "") then 
			%>
					<span class="black small"><%=arr_other_session(int_i)%></span><span class="interval small">|</span>
			<%
				end if 
			next
		end if 
		%>
       <a href="#"><%=getWord("OTHER_CLASS_TIMETABLE")%></a><a href="#"><%=getWord("MY_CLASS_RECORD")%></a></li>
        <li><%=getWord("YOU_HAVE")%><span class="red">10</span><%=getWord("YOU_HAVE_POINT_2")%><a href="#"><%=getWord("DETAIL_DESCRIPTION")%></a></li>
        </ul>
        </div>
        <!--標題end-->
        <!--NEWS標頭-->
        <div class="con_news">
        <div class="con_header_left"><%=getWord("news")%></div>
        <div class="con_header_right small_12"><%=getWord("NEWS_1")%></div>
        <div class="clear"></div>
        </div>
        <!--NEWS標頭end-->
        
        <!--個人資料標頭-->
        <div class="con_close">
        <div class="con_header_left"><%=getWord("APPLY_FOR_NEW_LEVEL")%></div>
        <div class="con_header_right"><%=getWord("DATA_COMPLETE_RATE")%>：80% <a href="profile.asp"><%=getWord("EDIT_PERSONAL_DATA")%></a><a href="learning_dcgs.asp"><%=getWord("SET_DCGS")%></a><!--<a href="#">打印证书</a>--></div>
        <div class="con_header_right_02"><a href="#"><%=getWord("close")%> ─</a></div>
        <div class="clear"></div>
        </div>
         <!--個人資料標頭end-->
         <%
        dim str_client_email
        dim int_company_tel
        dim str_client_member_sql


        dim arr_result
        dim var_arr
      
        var_arr = Array(int_client_sn)
        str_client_member_sql = "select email,cname,fname,lname,sex,nlevel,birth,client_address,cphone_code+cphone as cphone" 
        str_client_member_sql = str_client_member_sql & ",htel_code,htel,otel_code,otel,clb_cphone_2 from client_basic where sn=@g_int_client_sn" 

        arr_result = excuteSqlStatementRead(str_client_member_sql, var_arr, CONST_VIPABC_RW_CONN)       
          if (isSelectQuerySuccess(arr_result)) then
            '有資料

            if (Ubound(arr_result) >= 0) then

            '   Response.Write("測試 SQL 的SELECT 語法(有值)<br>")
            '	Response.Write("欄:" & Ubound(arr_result) & "<br>")
            '	Response.Write("列:" & Ubound(arr_result,2) & "<br>")
            '	For int_row = 0 To Ubound(arr_result, 2)
            '        For int_col = 0 To Ubound(arr_result)
            '            Response.Write ("g_arr_result("&int_col&","&int_row&")=")&(arr_result(int_col, int_row) & "<br>")    ' col代表筆數, row代表欄位
            '        Next
            '    Next
            '沒有資料
               %>
                <div class="profile">
                   <ul>
                     <li><%=getWord("EMAIL_ACCOUNT")%></li>
                     <li class="li_3"><%=arr_result(0,0) %></li>
                     <li><%=getWord("PASSWORD")%></li>
                     <li class="li_2">******<a href="#"><%=getWord("MODIFY_PASSWORD")%></a></li>
                     
                     <li><%=getWord("CONTRACT_35")%></li>
                     <li class="li_2"><%=arr_result(1,0) %></li>
                     <li><%=getWord("ENAME")%></li>
                     <li class="li_2"><%=arr_result(2,0) %></li>
                     <li><%=getWord("fNAME")%></li>
                     <li class="li_2"><%=arr_result(3,0) %></li>
                     <%
                     dim int_client_sex
                     int_client_sex = arr_result(4,0)
                     if int_client_sex = 1 then
                       int_client_sex = getWord("MALE")
                     else
                       int_client_sex = getWord("feMALE")
                     end if 
                     dim int_client_level
                     int_client_level = arr_result(5,0)
             
					
					if (int_client_level <> "") then 
					    int_client_level = convertNumber(int_client_level, CONST_TW_NOR_NUMBER)
						
					 End if 
                     
                     %>
                     <li><%=getWord("grade_kind")%></li>
                     <li class="li_2"><%=getWord("ordinal_numer")%><%=int_client_level %><%=getWord("a_class")%><a href="#"><%getWord("APPLY_FOR_NEW_LEVEL")%></a></li>
                     <li><%=getWord("sex")%></li>
                     <li class="li_2"><%=int_client_sex %></li>
                     <li><%=getWord("BIRTHDAY")%></li>
                     <li class="li_2"><%=arr_result(6,0) %></li>
                     <li><%=getWord("COMMUNICATION_ADDRESS")%></li>
                     <li class="li_4"><%=arr_result(7,0) %></li>             
                     
                     <li><%=getWord("CONTACT_US_MOBIL_PHONE")%>1</li>
                     <li class="li_2"><%=arr_result(8,0) %></li>
                     <li><%=getWord("CONTACT_US_MOBIL_PHONE")%>2</li>
                     <li class="li_3"><%=arr_result(13,0) %></li>
                     <li><%=getWord("home_phone")%></li>
                     <li class="li_2"><%=arr_result(9,0) %>-<%=arr_result(10,0) %></li>
                     <li><%=getWord("OFFICE_PHONE")%></li>
                     <%
                         if  isnull(arr_result(11,0)) or isnull(arr_result(12,0))  then
                            int_company_tel=arr_result(11,0) &"-"& arr_result(12,0)
                         else
                            int_company_tel=getWord("NO_DATA")
                         end if
                         
                     
                      %>
                     <li class="li_2"><%=int_company_tel %></li>
                     <li><%=getWord("ROLE")%> </li>
                     <li class="li_2"></li>
                     <li><%=getWord("DCGS_11")%></li>
                     <li class="li_2"></li>
                     <li><%=getWord("DCGS_12")%></li>
                     <li class="li_2"></li>
                     <li><%=getWord("DCGS_5")%></li>
                     <li class="li_2"></li>
                   </ul>
                 </div>
               <%     
            else


                '錯誤訊息
            '	Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
            ''	Response.End()
	            'call sub alertGo
	           Call alertGo(getWord("DATA_ERROR")&" "&getWord("EMAIL_ERROR"), "", CONST_ALERT_THEN_GO_BACK)
	            Response.End
            end if
        end if       
                 
          %>
        
         <!---個人資料---->
         <!---個人資料end---->
         <!--我的收藏標頭-->
<!--         <div class="con_close">
           <div class="con_header_left">我的收藏</div>
           <div class="con_header_right"><a href="#">编辑我的收藏</a></div>
           <div class="con_header_right_02"><a href="#">关闭 ─</a></div>
           <div class="clear"></div>
         </div>-->
         <!--我的收藏標頭end-->
         <!--課程回顧-->
         <!--選單-->
<!--         <div class="main_memtitle">
           <div class="mune_01">课程回顾</div>
           <div ><a href="member_2.html" class="mune_02">最爱顾问名单</a></div>
           <div ><a href="member_3.html" class="mune_03">精选电子报</a></div>
         </div>
         <div class="clear"></div>-->
         <!--選單end-->
         <!--列表內容-->
<!--         <div class="menber_content_list">
           <ul>
             <li><span class="no">&nbsp;</span><span class="date">日期</span><span class="time">时间</span><span class="topic">主题</span><span class="consultant">顾问</span><span class="remove">&nbsp;</span></li>
             <li><span class="no">1</span><span class="date">2009 / 10 / 22</span><span class="time">14:30</span><span class="topic">英文初学者的单字小百科 - 客厅英文初学者的单字小百科 - 客厅</span><span class="consultant">Janelle Harrington</span>
                 <input type="submit" name="button2" id="button" value="+ 观看录像文件" class="btn_2"/>
             </li>
             <li><span class="no">2</span><span class="date">2009 / 10 / 22</span><span class="time">14:30</span><span class="topic">英文初学者的单字小百科 - 客厅</span><span class="consultant">Janelle Harrington</span>
                 <input type="submit" name="button2" id="button" value="+ 观看录像文件" class="btn_2"/>
             </li>
             <li><span class="no">3</span><span class="date">2009 / 10 / 22</span><span class="time">14:30</span><span class="topic">英文初学者的单字小百科 - 客厅</span><span class="consultant">Janelle Harrington</span>
                 <input type="submit" name="button2" id="button" value="+ 观看录像文件" class="btn_2"/>
             </li>
             <li><span class="no">4</span><span class="date">2009 / 10 / 22</span><span class="time">14:30</span><span class="topic">英文初学者的单字小百科 - 客厅</span><span class="consultant">Janelle Harrington</span>
                 <input type="submit" name="button2" id="button" value="+ 观看录像文件" class="btn_2"/>
             </li>
             <li><span class="no">5</span><span class="date">2009 / 10 / 22</span><span class="time">14:30</span><span class="topic">英文初学者的单字小百科 - 客厅</span><span class="consultant">Janelle Harrington</span>
                 <input type="submit" name="button2" id="button" value="+ 观看录像文件" class="btn_2"/>
             </li>
           </ul>
           <div class="clear"></div>
         </div>-->
         <!--列表內容end-->
         <!--儲存選單-->
<!--         <div class="main_storage"> <span class="left_member">您收藏的课程回顾共 4 堂</span> <span class="right_page"><a href="#">上一頁</a><span class="right_text">1</span><a href="#">2</a><a href="#">3</a><a href="#">4</a><a href="#">5</a><a href="#">6</a><a href="#">7</a><a href="#">8</a><a href="#">下一頁</a></span>
             <div class="clear"></div>
         </div>-->
         <!--儲存選單emd-->
         <!--列表內容end-->
         <!--課程回顧end-->
         <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
       </div>
       <!--內容end-->
       
       
       
  
       
