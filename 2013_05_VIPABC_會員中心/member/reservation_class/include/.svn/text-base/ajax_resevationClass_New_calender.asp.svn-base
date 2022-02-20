<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include_prepare_and_check_data.inc"-->
<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class.asp"-->

<!--#include virtual="/program/member/reservation_class/functions/functions_reservation_class_order_class.asp"-->
<%
    '萬年曆的製作原則是找出判斷平、閏年的法則，然後定出所求當年或當月的第一天是星期幾。
    Dim i, int_day_num
    
    Dim int_dat_year, int_dat_month, int_dat_day	'選到的月份要用到的
	Dim int_ajax_res_class_type						'訂位課程的型類
		
    Dim int_arr_month_day(12)						'宣告每個月份的天數，而預設二月為平年的天數28天
	Dim str_arr_week_name(8)						'ASP Weekday 傳回的是 1(星期日),2(星期一),3(星期二),4(星期三),5(星期四),6(星期五),7(星期六)
	
    '為判斷元旦而宣告一變數 int_dat_year_pass 從西元紀元始至去年共過了幾年
    '變數int_dat_year_first 表示元旦是屬於星期幾  ps:在vb中 \ 是求除後整數，而 / 是求除後商數可含小數
    Dim int_dat_year_pass, int_dat_year_first
    
    Dim int_day_pass								'變數 intdaypass 表示今年元旦起至上個月底已過了多少天

    Dim int_dat_month_of_week_first					'變數 int_dat_month_of_week_first 表示要顯示月份的當月第一天是星期幾
    
	Dim str_week_day : str_week_day = ""			'週曆標題中的日期 (月/日)
	Dim str_week_head : str_week_head = ""			'週曆的標題 (月/日 星期幾)

	Dim d_start_date : d_start_date = Date()		'週曆要顯示的第一天
	Dim arr_selectable_day							'可以訂課的時間array(每一天一個array)
	Dim arr_arr_client_booking_time_and_type		'以訂位的時間與型類(每一天一個array)
													'有三層, 第一層為日期, 第二層為一個個訂位資料, 第三層為{時間, 型類}
	Dim int_hour_index								'hour(0~23)的index
	Dim int_day_index								'day(0~6)的index
	Dim int_data_index								'可訂課資料的index
	Dim str_reserved								'是否已訂位了，並且要顯示的圖的src
	Dim b_check_reserved							'若已訂位，是否要顯示check box已取消訂位
	Dim str_reserved_sn								'已訂位課程的sn
	Dim	b_reservatale
	
	Dim str_time_block_color : str_time_block_color = ""	'一對多課程需要特別標示24小時內的訂課時間
	
    '變數int_dat_year為萬年曆運算所需的年，變數int_dat_month為萬年曆運算所需的月
    '自動抓取電腦的系統時間做出當月的月曆
    int_dat_year = Trim(Request("dat_year"))
    int_dat_month = Trim(Request("dat_month"))
    int_dat_day = Trim(Request("dat_day"))
	int_ajax_res_class_type = Cint(Trim(Request("class_type")))

	
    '*************** 設定預設值 START ***************
    '目前的年度
    if (isEmptyOrNull(int_dat_year)) then
        int_dat_year = Year(now) '年度的格式為 2003 四位數
    else
        int_dat_year = Int(int_dat_year)
    end if

    '目前的月份
    if (isEmptyOrNull(int_dat_month)) then
        int_dat_month = Month(now) '月份的格式為 6 一位數
    else
        int_dat_month = Int(int_dat_month)
    end if
	
    '目前的日期
    if (isEmptyOrNull(int_dat_day)) then
        int_dat_day = Day(now)
    else
        int_dat_day = Int(int_dat_day )
    end if

	
    '宣告每個月份的天數，而預設二月為平年的天數28天
    int_arr_month_day(0) = 0
    int_arr_month_day(1) = 31
    int_arr_month_day(2) = 28
    int_arr_month_day(3) = 31
    int_arr_month_day(4) = 30
    int_arr_month_day(5) = 31
    int_arr_month_day(6) = 30
    int_arr_month_day(7) = 31
    int_arr_month_day(8) = 31
    int_arr_month_day(9) = 30
    int_arr_month_day(10) = 31
    int_arr_month_day(11) = 30
    int_arr_month_day(12) = 31
   
    'ASP Weekday 傳回的是 1(星期日),2(星期一),3(星期二),4(星期三),5(星期四),6(星期五),7(星期六)
    str_arr_week_name(0) = str_msg_font_saturday
    str_arr_week_name(1) = str_msg_font_sunday
    str_arr_week_name(2) = str_msg_font_monday
    str_arr_week_name(3) = str_msg_font_tuesday
    str_arr_week_name(4) = str_msg_font_wednesday
    str_arr_week_name(5) = str_msg_font_thursday
    str_arr_week_name(6) = str_msg_font_friday

    '判斷是否為閏年或平年。閏年的話，二月則為29天；平年二月則為28天(預設值)。
    if ((int_dat_year mod 4 = 0) and not(int_dat_year mod 100 = 0)) or (int_dat_year mod 400 = 0) then
        int_arr_month_day(2) = 29
    end if

    '為判斷元旦而宣告一變數 int_dat_year_pass 從西元紀元始至去年共過了幾年
    '變數int_dat_year_first 表示元旦是屬於星期幾  ps:在vb中 \ 是求除後整數，而 / 是求除後商數可含小數
    int_dat_year_pass = int_dat_year - 1
    int_dat_year_first = (int_dat_year + int_dat_year_pass \ 4 - int_dat_year_pass \ 100 + int_dat_year_pass \ 400) mod 7 

        
    '變數 intdaypass 表示今年元旦起至上個月底已過了多少天
    int_day_pass = 0
    for i = 0 to (int_dat_month - 1)
        int_day_pass = int_day_pass + int_arr_month_day(i)
    next

    
    '變數 int_dat_month_of_week_first 表示要顯示月份的當月第一天是星期幾
    int_dat_month_of_week_first = (int_dat_year_first + (int_day_pass mod 7)) mod 7
    
    '*************** 設定預設值 END ***************

	'取得要查詢的那一天
	str_week_day = Right("0" & int_dat_month, 2) & "/" & (int_dat_day-1) & "/" & int_dat_year
	'取得int_dat_year/int_dat_month/int_dat_day這一天是星期幾
	int_dat_month_of_week_first = (int_dat_month_of_week_first + int_dat_day) mod 7
%>
<!--課程列表start-->

    <table cellpadding="0" cellspacing="0" class="tab" width="778">
        <!-- 星期幾 標題 START -->
		<tr>
			<th class="t1" width="68">时间/日期</th>
			
			<%
				'週曆標題
				For i = 0 To 6
					str_week_day = DateAdd("d", 1, str_week_day)
					
					If (str_week_day = Date()) Then	'如果是今天
						str_week_head = str_week_head & "<th class=""in"">"
					Else
						str_week_head = str_week_head & "<th>"
					End If
					
					str_week_head = str_week_head & Right("0" & Month(str_week_day), 2) & "/" & Right("0" & Day(str_week_day), 2)
					str_week_head = str_week_head & str_arr_week_name((int_dat_month_of_week_first + i) mod 7)
					str_week_head = str_week_head & "</th>"
				Next
				
				Response.Write str_week_head
			%>
		</tr>
        <!-- 星期幾 標題 END -->

		<!--
			每天每個時段是否有定位
		-->
		<%
			'取得要查詢的前一天的Date物件
			str_week_day = dateadd("yyyy", (int_dat_year - Year(str_week_day)), str_week_day )
			str_week_day = dateadd("m", (int_dat_month - Month(str_week_day)), str_week_day )
			str_week_day = dateadd("d", (int_dat_day - Day(str_week_day)-1), str_week_day )
			
			
			ReDim arr_selectable_day(6)
			ReDim arr_arr_client_booking_time_and_type(6)
			'Response.Write(UBound(arr_arr_client_booking_time_and_type) & " a<br />")
			
			For int_day_index = 0 To 6
				str_week_day = dateadd("d", str_week_day, 1)
				
				if (DateDiff("d", str_service_start_date, str_week_day)>0) then
					'在服務開始日期之前 (不能訂位)
					arr_selectable_day(int_day_index) = ""
					
				elseif (DateDiff("d", str_week_day, str_service_end_date)>0) then
					'在服務結束日期之後 (不能訂位)
					arr_selectable_day(int_day_index) = ""
				
				elseif (DateDiff("d", str_last_first_session_datetime, str_week_day)>0) then
					'在最近一次預約首測測試時間之後 (不能訂位)
					arr_selectable_day(int_day_index) = ""
					
				elseif (DateDiff("d", DateAdd("m", 1, Now()), str_week_day)>0) then
					'今日起一個月後的時間 (不能訂位)
					arr_selectable_day(int_day_index) = ""
				else
					'取得可訂位的時間
					arr_selectable_day(int_day_index) = getSelectableClassTime (Year(str_week_day) & "/" & Right("0" & Month(str_week_day), 2) & "/" & Right("0" & Day(str_week_day), 2), int_ajax_res_class_type, g_var_client_sn, "", CONST_VIPABC_RW_CONN)
				end if
				'TODO: test 需要去掉
				'
				arr_selectable_day(int_day_index) = getSelectableClassTime (Year(str_week_day) & "/" & Right("0" & Month(str_week_day), 2) & "/" & Right("0" & Day(str_week_day), 2), int_ajax_res_class_type, g_var_client_sn, "", CONST_VIPABC_RW_CONN)
				
				
				'TODO: test
				'arr_arr_client_booking_time_and_type = getCustomerBookingTimeAndType(Year(str_week_day) & "/" & Right("0" & Month(str_week_day), 2) & "/" & Right("0" & Day(str_week_day), 2), g_var_client_sn, CONST_VIPABC_RW_CONN)
				arr_arr_client_booking_time_and_type(int_day_index) = getCustomerBookingTimeAndType(Year(str_week_day) & "/" & Right("0" & Month(str_week_day), 2) & "/" & Right("0" & Day(str_week_day), 2), 205304, CONST_VIPABC_RW_CONN)
				'Response.Write(UBound(arr_arr_client_booking_time_and_type(int_day_index)) & " 2<br />")
				if (isArray(arr_arr_client_booking_time_and_type(int_day_index))) Then
					'有值
					'Response.Write(arr_arr_client_booking_time_and_type(int_day_index)(0)(0)) & "<br />"
					'Response.Write(UBound(arr_arr_client_booking_time_and_type(int_day_index)) & " b<br />")
				End If
				
			Next
			'Response.Write(UBound(arr_arr_client_booking_time_and_type) & " aa<br />")
			
			
			'Response.Write("int_ajax_res_class_type: " & int_ajax_res_class_type & "<br />")
			'For int_day_index = 0 To 6
			'	Response.Write(arr_selectable_day(int_day_index) & "<br />")
			'Next
			For int_day_index = 0 To 6
				arr_selectable_day(int_day_index) = Split(arr_selectable_day(int_day_index), ",")
			Next
			
			
			For int_hour_index = 0 To 23	'每一個時段
		%>
				<tr>
					<!--時段的title-->
					<td class="t1"><%=Right("0" & int_hour_index, 2)%>:30</td>
		<%

				For int_day_index = 0 To 6	'每一天
					'取得要查詢的前一天的Date物件
					str_week_day = dateadd("yyyy", (int_dat_year - Year(str_week_day)), str_week_day )
					str_week_day = dateadd("m", (int_dat_month - Month(str_week_day)), str_week_day )
					str_week_day = dateadd("d", (int_dat_day - Day(str_week_day)-1), str_week_day )
					
					'調整日期, (查詢的這一天)
					str_week_day = dateadd("d", str_week_day, int_day_index)	
					'調整小時
					str_week_day = dateadd("h", int_hour_index, str_week_day )
					
					
					'TODO: 取得已訂課資料
					'--	special_sn is not null 大會堂
					'--	attend_livesession_types	7 VIP (一對一)
					'--	attend_livesession_types	4 (一對多)
					str_reserved = ""
					b_check_reserved = false
					str_reserved_sn = ""
					if (isArray(arr_arr_client_booking_time_and_type(int_day_index))) Then
						'有訂位資料
						For int_data_index = 0 To UBound(arr_arr_client_booking_time_and_type(int_day_index))
							'如果訂位資料的上課時間是這個時間
							If (CInt(arr_arr_client_booking_time_and_type(int_day_index)(int_data_index)(0)) = int_hour_index) Then
								'如果訂位為大會堂
								If (arr_arr_client_booking_time_and_type(int_day_index)(int_data_index)(1) = CONST_CLASS_TYPE_LOBBY) Then
									'顯示的圖
									str_reserved = "/images/reservation_class/ss_icon.gif"
									If (int_ajax_res_class_type=CONST_CLASS_TYPE_LOBBY) Then
										b_check_reserved = true
										str_reserved_sn = arr_arr_client_booking_time_and_type(int_day_index)(int_data_index)(2)
									End If
									
								'如果訂位為一對一
								ElseIf (arr_arr_client_booking_time_and_type(int_day_index)(int_data_index)(1) = CONST_CLASS_TYPE_ONE_ON_ONE) Then
									'顯示的圖
									str_reserved = "/images/reservation_class/reserved_vip.gif"
									If (int_ajax_res_class_type=CONST_CLASS_TYPE_ONE_ON_ONE) Then
										b_check_reserved = true
										str_reserved_sn = arr_arr_client_booking_time_and_type(int_day_index)(int_data_index)(2)
									End If
									
								'如果訂位為一對多
								ElseIf (arr_arr_client_booking_time_and_type(int_day_index)(int_data_index)(1) = CONST_CLASS_TYPE_ONE_ON_THREE) Then
									'顯示的圖
									str_reserved = "/images/reservation_class/reserved.gif"
									If (int_ajax_res_class_type=CONST_CLASS_TYPE_ONE_ON_THREE) Then
										b_check_reserved = true
										str_reserved_sn = arr_arr_client_booking_time_and_type(int_day_index)(int_data_index)(2)
									End If
									
								End If
								Exit For	'找到資料，離開Loop
								
							End If
						Next
					End If
					
					
					'如果這個時間還沒訂課:
					
					'大會堂
					If (int_ajax_res_class_type=CONST_CLASS_TYPE_LOBBY) Then
						'取得lobby session資料
						'str_reserved_sn (若有已訂位，此為其sn)
						
						
					
					'一對一、一對多等等
					Else
						'從已取得的可訂位資料(arr_selectable_day) 判斷/找出這時是否可訂位
						b_reservatale = False
						IF(IsArray(arr_selectable_day(int_day_index))) Then	'如果可訂課資料是Array
							For int_data_index = 1 To ( Ubound(arr_selectable_day(int_day_index))-1 )
								'這個時間=可訂位的時間
								If (int_hour_index = CInt(arr_selectable_day(int_day_index)(int_data_index))) Then
									b_reservatale = True
								End If
							Next
							
						Else	'取得資料有問題
							b_reservatale = False
						End If
						
						'1.如果是一對一 just show
						'2.如果是一對多 (24外 -> 白色背景), (24-12, 12-4, 4-2 -> 紅色背景)
						'3.TODO 如果是大會堂
						str_time_block_color = ""
						If (CInt(int_ajax_res_class_type)=CInt(CONST_CLASS_TYPE_ONE_ON_THREE)) Then
							If (CInt(DateDiff("h", Now(), str_week_day))<0) Then
								str_time_block_color = "bgcolor=""#FFA6A6"""
							End If
						End If
					End If
					
					'輸出結果：
					'用str_reserved顯示已訂課圖案
					'用b_check_reserved判斷是否顯示check box讓使用者取消訂位
					If (str_reserved<>"") Then
		%>
					<td align="center" background="<%=str_reserved%>">
		<%
						If (b_check_reserved) Then
		%>
						<input type="checkbox" id="reserved_<%=str_reserved_sn%>" name="chk_reserved_class" checked="true" />
		<%
						End If
		%>
					</td>
		<%
					'用b_reservatale判斷可否訂課
					elseIf (b_reservatale) Then
					'可以訂課
						'一對一，大會堂：沒差
					
						'一對多：特別標示一天之內 (by str_time_block_color)
		%>
					<td <%=str_time_block_color%>><input type="checkbox" id="<%=Year(str_week_day)%><%=Right("0" & Month(str_week_day), 2)%><%=Right("0" & Day(str_week_day), 2)%><%=Right("0" & int_hour_index, 2)%>" name="chk_res_class" /></td>
		<%
					Else
					'不能訂課
						'如果是一對一: 灰色
						'如果是一對多: 灰色
						If (CInt(int_ajax_res_class_type)=CInt(CONST_CLASS_TYPE_ONE_ON_ONE)) Then
							str_time_block_color = "bgcolor=""#CCCCCC"""
						ElseIf (CInt(int_ajax_res_class_type)=CInt(CONST_CLASS_TYPE_ONE_ON_THREE)) Then
							str_time_block_color = "bgcolor=""#CCCCCC"""
						Else
							str_time_block_color = ""
						End If
		%>
					<td <%=str_time_block_color%>>&nbsp;</td>
		<%
					End If
				Next
		%>
				</tr>
		<%
			Next
		%>
    </table>
<!--課程列表end-->