<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<link href="/css/new_guide.css" rel="stylesheet" type="text/css" />
<link href="/css/style.css" rel="stylesheet" type="text/css" />
<link href="/css/first_time/guide.css" rel="stylesheet" type="text/css" />
<%
Dim strExternalLink                 : strExternalLink               = getRequest("externalLink", CONST_DECODE_NO)       '是不是後來外部連結 是=1 不是=空
Dim strTestflag_computer            : strTestflag_computer          = getRequest("testflag_computer", CONST_DECODE_NO)  '重新測試的原應為 換電腦重測?電腦重灌?雜音....
Dim intclient_sn                    : intclient_sn                  = getSession("client_sn", CONST_DECODE_NO)          '客戶編號      'SQL資料參數
Dim strClientEname                  : strClientEname                = getSession("mh_userfullname", CONST_DECODE_NO)    '客戶名  
Dim bolFirstTestQuestionaryPassed2  : bolFirstTestQuestionaryPassed2= getSession("bolFirstTestQuestionaryPassed", CONST_DECODE_NO)      '是否有通過首冊測試=1
'沒有客戶SN先導回首頁--開始--
Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")     

%>
<%
'''''''''''''這段從 ABC官網 E:\web\www.tutorabc.com\asp\member\first_time\guide-2firstsession.asp copy來 有些還不清楚流程定義
	'檢查客戶目前是否有買產品(如果有買產品則無法登入Demo Room) 2006/7/4 Sam Chiang.....開始	
	sqlstr="SELECT client_sn FROM dbo.client_purchase WHERE (client_sn = "& Session("mh_usersn") &")"
	intResult = objPCInfo.excuteSqlStatementEach( sqlstr , Array() , CONST_TUTORABC_RW_CONN)
	purchase_str=0
	IF not objPCInfo.eof Then
		 purchase_str=0
	Else
		 purchase_str=1
	End IF
	'檢查客戶目前是否有買產品(如果有買產品則無法登入Demo Room) 2006/7/4 Sam Chiang.....結束
	'mydate=cdate("2006/10/5")
	mydate=date
	'定義初始值......開始
		total_attend_num=0
		allattendsr=""
		x_allattendsr=""
		'可預約的checkbox數量
		session_select_num_1=0
		'已預約的數量
		session_select_num_2=0
		attend_session_str=""
	'定義初始值......結束

	sql="SELECT Sn, CONVERT(varchar, HDate, 112) AS hdate, Block AS block, Holiday_E AS holiday FROM company_calendar WHERE (HDate >= '"&date&"' AND HDate <= '"&dateadd("d",mydate,6)&"')"
	intResult = objPCInfo.excuteSqlStatementEach( sql , Array() , CONST_TUTORABC_RW_CONN)
	while not objPCInfo.eof
		holidaystr=holidaystr&objPCInfo("hdate")&"+"&objPCInfo("block")&"*"&objPCInfo("holiday")&"/"
		holidayD=holidayD&objPCInfo("hdate")&"+"&objPCInfo("block")&"*"           '20090604 joey add by RD09052709
		objPCInfo.movenext
	wend
	
	aclient_sn=session("client_sn")  '94 line
	nowweek=Weekday(now())
	nowhour=(Hour(now()) + 1)
	nowMinute=Minute(now())
	nowtime=nowhour&":"&nowMinute
	
	dclient_sn=session("client_sn")  'guide-2firstsession.asp  103 line
	dnow_date=dateValue(now())
	
	'如果產品不在該時間點內  'guide-2firstsession.asp  105 line
	sqlT="SELECT * FROM client_purchase where  client_sn='"&dclient_sn&"' and product_edate>'"&dnow_date&"' and valid=1 "
	intResult = objPCInfo.excuteSqlStatementEach( sqlT , Array() , CONST_TUTORABC_RW_CONN)
	if objPCInfo.RecordCount < 1 then
	%>
	<script type="text/javascript">
		alert("你无法使用此网页");
		//////window.close();
	</script>
	<%
		'''response.End
	end if
	
	baurl=request("aurl")	'guide-2firstsession.asp 119 line
	if baurl=1 then
	   baurl="reservation_lobby.asp"
	elseif baurl=2 then
	   baurl="reservation.asp"
	elseif aurl=3 then
	   baurl="vip_requirements.asp"
	end if
	
	sqld="SELECT fname,lname FROM client_basic where sn='"&aclient_sn&"'"
	intResult = objPCInfo.excuteSqlStatementEach( sqld , Array() , CONST_TUTORABC_RW_CONN)
	if not objPCInfo.EOF then
		cust_name=objPCInfo("fname")
		cust_lame=objPCInfo("lname")
	end if
	
	dim weekname(6)
	weekname(0)="日"  '星期日
	weekname(1)="一"  '星期一
	weekname(2)="二"  '星期二
	weekname(3)="三"  '星期三
	weekname(4)="四"  '星期四
	weekname(5)="五"  '星期五
	weekname(6)="六"  '星期六
	
	dim intTimeQuantity : intTimeQuantity = 28  '20120220 有幾個測試時間
	'dim opentime(1)
	ReDim opentime(intTimeQuantity)     ' 宣告陣列大小
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
	
	sqlE="SELECT * FROM customer_first_sesion where mod_type=1"   'guide-2firstsession.asp 195 line
	intResult = objPCInfo.excuteSqlStatementEach( sqlE , Array() , CONST_TUTORABC_RW_CONN)
	if not objPCInfo.EOF then
		af1=int(objPCInfo("count_people_1"))
		af2=int(objPCInfo("count_people_2"))
		af3=int(objPCInfo("count_people_3"))
		af4=int(objPCInfo("count_people_4"))
		af5=int(objPCInfo("count_people_5"))
		af6=int(objPCInfo("count_people_6"))
		af7=int(objPCInfo("count_people_7"))
	end if

	
	'Tree  223 line

	' 到 258 line 好像沒什麼作用先不弄
	
	'--判斷客戶合約開始日--開始--
	Dim intArrSize : intArrSize = 0             'SQL算陣列的數量 0~N
	Dim dateClientPurchaseContractEdate : dateClientPurchaseContractEdate = ""  '客戶合約開始日
	'/***********SQL資料陣列用法**************/
	strSql =    getSqlNotes() &_  
				"   SELECT     TOP 1                        "&_
				"       convert(CHAR,contract_sdate,111) as contract_sdate   "&_
				"   FROM                                    "&_
				"       client_purchase WITH (nolock)       "&_
				"   WHERE                                   "&_
				"       (client_sn = @intclient_sn)         "                                                           '建立連線
	intResult = objPCInfo.excuteSqlStatementEach( strSql , Array(intclient_sn) , CONST_TUTORABC_RW_CONN)
	intArrSize = objPCInfo.recordcount-1          '取的陣列大小 0~N
	
	if ( 0 <= intArrSize ) then
		dateClientPurchaseContractEdate = objPCInfo("contract_sdate")
	end if
	'--判斷客戶合約開始日--結束--
	
	
	'新增的
	'判斷是否做過首測
	Dim haveDoFirstTest : haveDoFirstTest = 0
	strSql = "SELECT TOP 1 * FROM add_client_datetime WITH(NOLOCK)  WHERE client_sn =@add_client_datetime_sn  and flag=5  and status=6  ORDER BY sn desc "  'and staff_sn =275 and testflag=1 考慮到 ims建立 故不考慮設限自動化測試
	intResult = objPCInfo.excuteSqlStatementEach( strSql , Array( intclient_sn  ) , CONST_TUTORABC_RW_CONN)
	flag_is_first_time_session = "F"
	if not objPCInfo.EOF then
		flag_is_first_time_session = "R"
	end if
%>
<script type="text/javascript">   
  //chooseOne()函式，參數為觸發該函式的元素本身   
    function chooseOne(cb){   
        //先取得同name的chekcBox的集合物件   
        var obj = document.getElementsByName("chk");   
         for (i=0; i<obj.length; i++){   
             //判斷obj集合中的i元素是否為cb，若否則表示未被點選   
             if (obj[i]!=cb) obj[i].checked = false;   
             //若是 但原先未被勾選 則變成勾選；反之 則變為未勾選   
             //else  obj[i].checked = cb.checked;   
             //若要至少勾選一個的話，則把上面那行else拿掉，換用下面那行   
             else obj[i].checked = true;   
         }   
     }

	//將測試項目，寫入hidden值
	function writeComputerProblems()
	{
	   // 測試項目描述
		var str_testflag_description ="(";
		for (var i = 0; i < $(":checkbox[name='testflag_computer']:checked").size(); i++)
		{
			if (i > 0)
				str_testflag_description += "," + $($(":checkbox[name='testflag_computer']:checked")[i]).attr("title");
			else
				str_testflag_description += $($(":checkbox[name='testflag_computer']:checked")[i]).attr("title");
		}
		str_testflag_description += ")";
		$("#str_computer_problems").val(str_testflag_description);
	}
	function chkdata()
	{
		writeComputerProblems();
		var len = document.fm1.chk.length;
		var class_time ="";
		var str_session_type = "<%=flag_is_first_time_session%>"
		//20100303 首測重測 lucy
		var obj = document.getElementById("div_computer_retest_detail");
		var checked = true;
		if(obj.style.display!="none") {	//div_computer_retest_detail顯示與否來判斷
			obj = document.getElementsByName("testflag_computer");
			checked = false;
			for (i=0; i<obj.length; i++) {
				if (obj[i].checked==true) {
					checked = true;
				}
			}
		}
		if (checked==false) {
			alert("请点选欲重新测试电脑的原因");//請點選欲重新測試電腦的原因  <%=strLangMemGuide2f00002%>
			obj[0].focus();
			return false;
		}
		//20100303 首測重測 lucy
		for (i=0;i<len;i++){

			if (document.fm1.chk[i].checked==true){

				class_time = document.fm1.chk[i].value;
			}
		}


		if (class_time == "")
		{
			alert("请勾选预约时间");//請勾選預約時間  <%=strLangMemGuide2f00003%>
			return false;
		}

		if (str_session_type=="F")
		{
			if (!confirm("<%' 提醒您！建議您完成預約測試後再體驗第一堂課，所以您預約訂位的時間需於測試時間 %><%="提醒您！建议您完成预约测试后再体验第一堂课，所以您预约订位的时间需于测试时间"%>"+class_time+"<% '之後。如需修改請按「取消」或按「確定」完成First Session測試預約。 %><%="之后。如需修改请按「取消」或按「确定」完成First Session测试预约。"%>") )			
			{
				//alert('asdf');

				return false;
			}
		}

		document.fm1.Submit.disabled=true;
		return true;
	}
	function chkdata2()
	{
        
        //var len = document.myform.chk.length;
        //class_time = document.myform.chk.value;
        //if (class_time == "")
        //{
        //    window.alert(len)
        //    return false;
        //}
    }
	function chkAlreadyOrderTestTime(str_check_time,obj)
	{
        //檢查是否沒通過首冊 沒通過要檢查試合約開始時間--開始--
        var str_FirstTestQuestionaryPassed = "<%=bolFirstTestQuestionaryPassed2%>"
        var date_ClientPurchaseContractEdate = "<%=dateClientPurchaseContractEdate%>"
        if ("1" != str_FirstTestQuestionaryPassed)
        {
            if (str_check_time > date_ClientPurchaseContractEdate)
            {
                window.alert("<%="建议您于合约开始前完成计电脑试预约，谢谢!"%>")//建議您於合約開始前完成電腦測試預約，謝謝!
            }
        }
        //檢查是否沒通過首冊 沒通過要檢查試合約開始時間--結束--


		$.ajax({ 
				url: 'ajax_chk_order_session.asp',
				data:{client_sn:<%=session("client_sn")%>},
				cache: false, 
				error: function(xhr) {alert("错误"+xhr.responseText)},
				success: function(result) {
					if(result != "")
					{
						//您在2010/3/6 10:30 (原預約時段)已預約測試，您確定更改預約時間為 2010/3/6 14:30(更改預約時段)嗎?  ※提醒您更改預約時間尚未完成,點擊下方確認預約鍵後才完成更改預約時間!
						var answer = confirm("<%="您在"%>"+result+"<%="(原预约时段)已预约测试，您确定更改预约时间为"%>"+str_check_time+"<%="(更改预约时段)吗?  ※提醒您更改预约时间尚未完成,点击下方确认预约键后才完成更改预约时间!"%>");
						if (answer)
						{
							$(obj).attr("checked",true);
							//$("#submit_order_session_time").attr("disabled",true);
							$("#myform").submit();
						}
						else
						{
							$(obj).attr("checked",false);
							//$("#submit_order_session_time").attr("disabled",true);
						}
					}
				}
			}); 
	}
	$(document).ready(function(){
//		$.ajax({ 
//				url: '../ajax_chk_order_session.asp',
//				data:{client_sn:<%=session("client_sn")%>},
//				cache: false, 
//				async: false,
//				error: function(xhr) {
//				alert("错误"+xhr.responseText)
//				//$("body").html(xhr.responseText);
//				},
//				success: function(result) {
//					if(result != "")
//					{
//						//您在2010/3/6 10:30 (原預約時段)已預約測試，您確定更改預約時間為 2010/3/6 14:30(更改預約時段)嗎?
//						var answer = confirm("您在"+result+"(原预约时段)已预约测试，您确定更改预约时间吗?");
//						if (answer)
//						{
//							$("#submit_order_session_time").attr("disabled",false);
//						}
//						else
//						{
//							window.location = "/program/member/member.asp" ;
//						}
//					}
//				}
//			}); 
		$("input[name='chk']").click(function(){
			chkAlreadyOrderTestTime($(this).val(),this);
		});
		//檢查所有input checkbox中的選項是不是可以選
		var int_numbers_of_input_checkbox = $("input[name='chk']").not($("input : disabled")).size();
		if (int_numbers_of_input_checkbox==0)
		{
			alert("<%="所有时段都已额满"%>");//所有時段都已額滿
			//window.location = "/program/member/member.asp" ;
		}
	});
</script>   
<!--內容start-->
<div id="menber_membox">
			<div align="center">
				<!-- ------------>
				<form method="POST" name="myform" action="guide-2firstsession_exe.asp?testflag_computer=<%=strTestflag_computer%>&externalLink=<%=strExternalLink%>" onsubmit="return chkdata2();">
				<div class="guide" align="left">
					<%if ( "1" = strExternalLink ) then %>
						<img src="/images/first_time/title_reset<%=strImagesLanguage%>.gif">																		
					<%else%>
						<img src="/images/first_time/title_guide<%=strImagesLanguage%>.jpg">
					<%end if%>
					<div class="guide-area" style="font-size:12px;">
						<div class="reset-desc"><img src="/images/environment_test/new_test_step4.png" /></div>

						<table class="firstsession-table01" style="font-size:12px;">
							<input type="hidden" name="test_session_type" value="<%=flag_is_first_time_session%>">
								  <tr>
									<th class="fs-time">
										<%=char00305%><!--顯示上排左邊第一個-->
									</th>
										<%for i = 0 to 6
										   orderdate = dateadd("d",i,dateadd("d",0,date))
										   syear = left(year(orderdate),4)		'22090604 joey mod by RD09052709			   
										   smonth = right("0"&month(orderdate),2)
										   sday    = right("0"&day(orderdate),2)
										   weeks     = weekday(orderdate)
										   symd=syear+smonth+sday				'22090604 joey mod by RD09052709	
									  %>
											<th class="fs-time">
												<%=smonth%>/<%=sday%>(<%=weekname(weeks-1)%>)<!--顯示上排日期-->
											</th>
										<%next%>
								  </tr>

								  <tr>
								  <%	
								  '==========判斷這個人是否已經在這個時段預約過，如果預約過且首測還沒通過 不能讓他選==start========
									Dim arr_result_of_order_date : arr_result_of_order_date = getAlreadyOrderTestTime(session("client_sn"),CONST_TUTORABC_RW_CONN) '判斷是否訂過課
									Dim bol_is_order_time : bol_is_order_time = false '是否有訂過
									Dim date_already_order : date_already_order = "" '已經訂過的時間
									if isArray(arr_result_of_order_date) then
										bol_is_order_time = true
										date_already_order = arr_result_of_order_date(0)

						'tree   response.write arr_result_of_order_date(0) & "<br>" 
						'tree   response.write arr_result_of_order_date(1) & "<br>"
						'tree   response.write arr_result_of_order_date(2) & "<br>"
						'tree   response.write arr_result_of_order_date(3) & "<br>"
						'tree   response.write arr_result_of_order_date(4) & "<br>"
						'tree   response.write arr_result_of_order_date(5) & "<br>"
						'tree   response.write arr_result_of_order_date(6) & "<br>"
						'tree   response.write arr_result_of_order_date(7) & "<br>"
						'tree   response.end

									else
										bol_is_order_time = false
									end if
									'==========判斷這個人是否已經在這個時段預約過，如果預約過且首測還沒通過 不能讓他選==end==========
								  for j = 1 to Cint(intTimeQuantity)
								  %>
									<td height="30"><%=opentime(j)%></td>   <!--左邊時間-->
									<!--tstr是上課時段-->
								  <%
								   for i =0 to 6
								   orderdate = dateadd("d",i,dateadd("d",0,date))
								   syear = left(year(orderdate),4)				    '22090604 joey mod by RD09052709	
								   smonth = right("0"&month(orderdate),2)
								   sday    = right("0"&day(orderdate),2)
								   weeks     = weekday(orderdate)
								   symd=syear+smonth+sday							'22090604 joey mod by RD09052709	

								   findH=InStr(holidayD,symd)    '22090610 joey mod by RD09052709

								   '判斷此堂課預約人數
								   YMD=year(orderdate)&"/"&smonth&"/"&sday
								   date_hour=Mid(opentime(j),1,2)  
								   date_min=Mid(opentime(j),4,5)
								   sqlE="select COUNT(*) as count_peo from add_client_datetime where YMD='"&YMD&"' and hour1='"&date_hour&"' and min1='"&date_min&"' AND (flag =4) "
								   intResult = objPCInfo.excuteSqlStatementEach( sqlE , Array() , CONST_TUTORABC_RW_CONN)
								   if not objPCInfo.EOF then
									count_people=objPCInfo("count_peo")
								   end if

								   '-------------------2010/02/24 Ryanlin 新增 GM09091813 首測重測登入頁面判斷導入 start-------------------
								   Dim cs_manpower_sql '用來查詢"當天"，"該時段"，的測試人力加總
								   cs_manpower_sql = "select sum(isnull(d_"&day(orderdate)&",0)) as manpower_numbers  from cs_manpower_schedule where year='"&year(orderdate)&"' and month='"&month(orderdate)&"' and time='"&opentime(j)&"'"
								   Dim g_arr_result '回傳的加總值
								   Dim int_test_manpower : int_test_manpower = 0
								   g_arr_result = excuteSqlStatementReadQuick(cs_manpower_sql, CONST_TUTORABC_RW_CONN)
								   '回傳二維陣列
									if (isSelectQuerySuccess(g_arr_result)) then
										'如果沒有值 則用0取代
										if Ubound(g_arr_result)< 0 or isnull(g_arr_result(0,0)) then
											g_arr_result(0,0) = 0
											int_test_manpower = g_arr_result(0,0)
										end if
										'如果當天的測試人力大於預約的人
											int_test_manpower = g_arr_result(0,0)
										'''response.write("<br>g_arr_result:" & g_arr_result(0,0) &"/count_people:"&count_people )
										if 	g_arr_result(0,0)*1 > count_people*1	 then
											'check_box設成開啟的狀態
											open="Y"
										else
											open="O"
										end if
										'''response.write("open:" & open )
									else

										'錯誤訊息
											'''response.write("錯誤訊息" & g_arr_result & "<br>")
									end if
									'判斷時段幾個人就額滿

									Dim t_day : t_day = Day(Date()) '今天
									Dim t_month : t_month = month(Date()) '這個月
									Dim t_year : t_year = year(Date()) '今年
									Dim check_status : check_status = "" 'Check_BOX狀態
									'判斷目前之前的CHECK_BOX都要disabled
									if DateDiff("n",Cdate(year(orderdate)&"-"&month(orderdate)&"-"&day(orderdate)&" "&opentime(j)),Cdate(t_year&"-"&t_month&"-"&t_day&" "&Time())) > -30 then
										check_status = "disabled"
									end if
								   '-------------------2010/02/24 Ryanlin 新增 GM09091813 首測重測登入頁面判斷導入 end-------------------


							'============首測重測新增判斷已經選過得日期=====start=======
							Dim str_show_order_session_image : str_show_order_session_image = ""
							if bol_is_order_time then
								if (Year(orderdate) = Year(date_already_order) and Month(orderdate) = Month(date_already_order) and Day(orderdate) = Day(date_already_order) ) then
									Dim format_Hour : format_Hour = ""
									if Len(Hour(date_already_order))<2 then
										format_Hour = "0"&Hour(date_already_order)
									else
										format_Hour = Hour(date_already_order)
									end if
									Dim format_min : format_min = ""
									if Len(Minute(date_already_order))<2 then
										format_min = "0"&Minute(date_already_order)
									else
										format_min = Minute(date_already_order)
									end if
									if opentime(j) = format_Hour &":"&format_min then
										open="O"
										str_show_order_session_image = " <img src='/images/first_time/reserved_b.gif'/>"
									end if
								end if
							end if
							'''response.write("open2:" & open )
							'============首測重測新增判斷已經選過得日期=====end=======
							
							
							int_test_manpower = CInt(int_test_manpower)
							count_people = CInt(count_people)
					
							if (int_test_manpower = 0) or (int_test_manpower <= count_people) then
								open="N"
							elseif (int_test_manpower > count_people) then
								open="Y"
							end if
							
							
							
							'-----------計算當天有的班表人力--------------start-----------------
							'Dim test_manpower '有的測試人力
							'Dim cs_work_schedule_manpower '班表人力
							'Dim work_time_sum,g_int_row
							'Dim g_arr_result_test_manpower 'sql接收陣列
							'Redim work_time_sum(29)
							'Dim s '計數
							''選出選取年份月份所有的員工排班表
							'sql="SELECT d_"&day(orderdate)&" FROM cs_work_schedule where year='"&year(orderdate)&"' and month='"&month(orderdate)&"'"
							''response.write sql
							''資料庫新寫法
							'g_arr_result_test_manpower = excuteSqlStatementReadQuick(sql, CONST_TUTORABC_RW_CONN)			
							'If (g_int_run_time_err_code = CONST_FUNC_EXE_SUCCESS) then
							'	'如果沒有值 則用0取代
							'	if (Ubound(g_arr_result_test_manpower)<0 or isnull(g_arr_result_test_manpower(0,0))) then
							'		g_arr_result_test_manpower(0,0) = 0
							'	end if
							'	For g_int_row = 0 To Ubound(g_arr_result_test_manpower, 2)
							'		if (isnull(g_arr_result_test_manpower(0,g_int_row))) then
							'			g_arr_result_test_manpower(0,g_int_row) = 0
							'		end if
							'		'判斷長時段類型 轉存計算後結果至 work_time陣列
							'		If (g_arr_result_test_manpower(0,g_int_row)="C") then
							'			for s=0 to 17
							'				work_time_sum(s)=work_time_sum(s)+1
							'			next
							'		elseIf((g_arr_result_test_manpower(0,g_int_row)="E") ) then
							'			for s=2 to 21
							'				work_time_sum(s)=work_time_sum(s)+1
							'			next
							'		elseIf((g_arr_result_test_manpower(0,g_int_row)="G") ) then
							'			for s=8 to 25
							'				work_time_sum(s)=work_time_sum(s)+1
							'			next
							'		elseIf((g_arr_result_test_manpower(0,g_int_row)="H") ) then
							'			for s=9 to 26
							'				work_time_sum(s)=work_time_sum(s)+1
							'			next	
							'		elseIf((g_arr_result_test_manpower(0,g_int_row)="I") ) then
							'			for s=10 to 27
							'				work_time_sum(s)=work_time_sum(s)+1
							'			next
							'		elseIf((g_arr_result_test_manpower(0,g_int_row)="K") ) then
							'			for s=11 to 28
							'				work_time_sum(s)=work_time_sum(s)+1
							'			next						
							'		End If	
							'	Next
							'end if
							'對應index把該時間對應班表人力找出來
							'Dim m
							'for m= 0 to Cint(intTimeQuantity)
							'	if work_time_sum(m)="" then
							'		work_time_sum(m)= 0
							'	end if
							'	if date_hour&":"&date_min = opentime(m) then
							'		cs_work_schedule_manpower = work_time_sum(m)
							'	end if 
							'next
							'如果當天的測試人力大於班表人力 block住
							'if cint(count_people) >= cint(cs_work_schedule_manpower) then
							'	open = "O"
							'end if
							''如果班表人力為零 直接blocked住
							'if cint(cs_work_schedule_manpower) = 0 or cint(int_test_manpower) = 0 then
							'	open = "O"
							'end if 
							'''response.write("open3:" & open  &"/"& count_people &"/"& cs_work_schedule_manpower &"/"& cs_work_schedule_manpower &"/"& int_test_manpower)
							'-----------計算當天有的班表人力--------------end-----------------
								if  open="Y" then   
								  %>
										<td>                                    <!--可預約的樣式-->
										<div align="center" >                           
							  <%else %>
									<td class="no-time">                        <!--不可預約的樣式-->
										<%=str_show_order_session_image %>      <!-- 顯示已預約的圖案-->
									<div align="center" >
							  <%     
								end if         
								'response.write "****"&open
									if open="Y" then %>
										<input type="checkbox" <%=check_status%> value="<%=year(orderdate) %>/<%=smonth %>/<%=sday&" " %> <%=opentime(j)%>" name="chk" onClick="chooseOne(this)" >
										<%elseif open="O" then%>
											<font size=2 color="#ff0000" class="block" ></font>
									<%end if %>
										</div></td>
								   <%next %>
								  </tr>
							<%next%>

					   

							<input type="hidden" name="language" value=<%=language%>>
							<!--------2010/02/24 Ryanlin 新增button GM09091813 首測重測登入頁面判斷導入 start------------------->
							<input type="hidden" name="aurl" value="4">
							<!--------2010/02/24 Ryanlin 新增button GM09091813 首測重測登入頁面判斷導入 end------------------->
							<input type="hidden" name="cust_name" value=<%=cust_name%>>
							<input type="hidden" name="cust_lame" value=<%=cust_lame%>>
						</table>
						
					   <table width=100%>
							<tr>
								 <td align="left">
									<span class="style1" ><%=strLangMemGuide2f00015%></span><%'客服專線：02-3365-9999%>
								 </td>
								 <td align="right">
									<!--<input type="button" name="Submit2" id="Submit2" value="上一步" class="btn-gray" onClick="location.href='XXXX.asp?language=<%=strlanguage%>'">-->
									<input type="submit" name="button" id="submit_order_session_time" value="确定预约" class="btn-org"><%'確定預約%>
								 </td>
							</tr>
						</table>
							
							
						
						<!--First Session END-->
					</div>
				</div>
				<%
				'    '先將問卷資料存起來等預約電腦測試後再寫入 應為以前邏輯問卷需要預約的資料(sn) 資料來自於E:\web\www.tutorabc.com\asp\questionary\guide-2environment_View.asp
				'    Dim strAnswerValue : strAnswerValue = ""  '存問卷的答案選項編號
				'    Dim strProblemName : strProblemName = ""  '存問卷的題目選項編號
				'    for i = 1 to Request.Form.Count
				'        strAnswerValue = strAnswerValue & "%#%" & Request.Form.Item(i)    '範例 %#%856    %#%858  %#%864, 123   %#%869    %#%878, 456 %#%880  %#%穩定－您目前的頻寬為2M以上   %#%883  %#%885  %#%887  %#%889  %#%填好送出
				'        strProblemName = strProblemName & "%#%" & Request.Form.Key(i)     '範例 %#%284    %#%285  %#%286        %#%287    %#%288      %#%289  %#%290                          %#%291  %#%292  %#%293  %#%294  %#%button
				'    next
				%>
				<!--<input type="hidden" name="AnswerValue" id="AnswerValue" value="<%=strAnswerValue%>" />--><!--資料..-->
				<!--<input type="hidden" name="ProblemName" id="ProblemName" value="<%=strProblemName%>" />--><!--資料..-->
				<input type="hidden" name="date_already_order" id="date_already_order" value="<%=date_already_order%>" /><!--原來預約時間　當沒有修改預約時間時將要改為原來的時間..-->
			</form>
				<!-- ------------>                                                                            
			</div>
 <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<%
objPCInfo.close                               '關閉物件
Set objPCInfo = nothing                       '設定物件為空
%>
<!--內容end-->