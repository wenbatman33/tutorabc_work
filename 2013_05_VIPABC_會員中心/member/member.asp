<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include file="reservation_class/include/include_prepare_and_check_data.inc"-->
<%
'response.write getSessionRoomTime
Dim sys_now_datetime : sys_now_datetime = getFormatDateTime(now(), 8)	'取得系統時間
Dim str_get_class_time : str_get_class_time = ""						'取得上課時間
Dim int_i 
Dim str_today_session : str_today_session = "" 	'今日下堂課時間
Dim str_special_sn : str_special_sn = ""		'大會堂sn
Dim str_other_session : str_other_session = "" 	'其他咨訊時段
Dim arr_other_session : arr_other_session = "" 	'其他咨訊時段陣列
Dim str_get_session_sn : str_get_session_sn = getSessionRoomTime()	  '欲找的session_sn
Dim str_session_sn : str_session_sn = ""
'response.write str_session_hour
Dim str_show_today	'顯示今日上課狀態
Dim int_dcgs_now	'現在的時間+3.5小時
Dim int_session		'實際上課時間
Dim int_now_time	'現在時間
Dim bol_room_bout : bol_room_bout = 0 'false (預設沒有"進教室"鈕)
	
Dim int_attend_level : int_attend_level="" '取得客戶的等級
Dim int_client_data_rate : int_client_data_rate = 0 '客戶數據的完成度
	
'檢查客戶是否有登入
if (isEmptyOrNull(g_var_client_sn)) then
	Call alertGo(getWord("PLEASE_LOGIN"), "/program/member/member_login.asp", CONST_ALERT_THEN_REDIRECT)
	Response.End
end if

dim bolDebugMode : bolDebugMode = false
dim intStartTime : intStartTime = ""
dim bolShortSession : bolShortSession = false

arr_result = getClientAttendList(g_var_client_sn, str_get_session_sn)
if ( isSelectQuerySuccess(arr_result) ) then
    if ( true = bolDebugMode ) then
        response.Write "Ubound(arr_result) : " & Ubound(arr_result) & "<br/>"
    end if	   
	if (Ubound(arr_result) >= 0) then
		for int_i = 0  to Ubound(arr_result,2)
            if ( true = bolDebugMode ) then
                response.Write "arr_result(2,int_i) : " & arr_result(2,int_i) & "<br/>"
                response.Write "arr_result(6,int_i) : " & arr_result(6,int_i) & "<br/>"
            end if
            intStartTime = isShortSession(arr_result(2,int_i), arr_result(6,int_i), 0, CONST_VIPABC_RW_CONN)
			if ( int_i = 0 and left(str_get_session_sn,8) = left(arr_result(2,int_i),8) ) then  '捉第一筆
                if ( false = intStartTime ) then
					str_today_session = getFormatDateTime(arr_result(0,int_i),6) &"&nbsp;/&nbsp;"&arr_result(1,int_i)&":30"
                else
                    intStartTime = right("00" & intStartTime, 2)
                    str_today_session = getFormatDateTime(arr_result(0,int_i),6) &"&nbsp;/&nbsp;"&arr_result(1,int_i)&":"&intStartTime
                    bolShortSession = true
                end if
                str_get_class_time = arr_result(1,int_i)
				str_special_sn = arr_result(4,int_i) 		'大會堂
				str_session_sn = arr_result(2,int_i) 		'session_sn
			else
				'找出其他咨訊時間
				if ( false = intStartTime ) then
                    str_other_session = str_other_session & "@" &getFormatDateTime(arr_result(0,int_i),6) &"&nbsp;/&nbsp;"&arr_result(1,int_i)&":30"
                else
                    intStartTime = right("00" & intStartTime, 2)
                    str_other_session = str_other_session & "@" &getFormatDateTime(arr_result(0,int_i),6) &"&nbsp;/&nbsp;"&arr_result(1,int_i)&":"&intStartTime
                end if
			end if 
		next
		'整理其他咨訊時間的值
		if ( str_other_session <> "") then 
			if left(str_other_session,1)="@" then str_other_session=right(str_other_session,len(str_other_session)-1)
			if right(str_other_session,1)="@" then str_other_session=left(str_other_session,len(str_other_session)-1)
			str_other_session = str_other_session & "@"
		end if 
		'response.write str_other_session & "<br>"
		'response.write str_today_session & "<br>"
	else
		'Call alertGo("資料錯誤，請洽客服人員", "", CONST_ALERT_THEN_GO_BACK)
		'Response.End
	end if
end if 
'str_get_class_time = "11"
'response.write str_today_session&"........."
%>
<!-- 2012/07/13 Hunter：TS 通報 Jing 已修正，因此移除。 -->
<!-- 2012/06/13 Hunter：Flash 11.3 升級影響 TutorMeet 上課加入判斷程式碼 Start -->
<!--<script src="/lib/javascript/detectFlash.js" type="text/javascript"></script>-->
<!--
<script type="text/javascript">
	var hasRequestedVersion = DetectFlashVer();

	if (!hasRequestedVersion) {
		//var alternateContent = 'TutorMeet content requires the Adobe Flash Player 10.3 ' + '<a href=flash/getflash.html>Get Flash Player 10.3</a>';
		//document.write(alternateContent);  // insert non-flash content
		//alert('Flash 版本不相符，需重新安裝！');
		var w = window.open("getflash.htm", "_blank", 'width=800,height=600,location=0,resizable=0,scrollbars=1');
		w.focus()
	}
</script>
-->
<!-- 2012/06/13 Hunter：Flash 11.3 升級影響 TutorMeet 上課加入判斷程式碼 End -->
<script type="text/javascript" src="/lib/javascript/JQuery/Scripts/jquery.timers.js"></script>
<script type="text/javascript">
<!--
// 系統時間
var str_now_sys_datetime = '<%=sys_now_datetime%>';  
var str_session = "<%=str_get_class_time%>"; //今日上課時間
var str_specianl = "<%=str_special_sn%>";	//大會堂
var CLASS_2 = "<%=getWord("CLASS_2")%>";
var CLASS_3 = "<%=getWord("CLASS_3")%>";
var close = "<%=getWord("close")%>";
var FAQ_2 = "<%=getWord("FAQ_2")%>";
<% if ( true = bolShortSession ) then %>
var int_room_min = "<%=intStartTime%>" ; //前三分鐘進教室
var int_end_min = "<%=intStartTime+10%>"; //結束倒數時間
var int_query_start_room_min = "<%=intStartTime%>" ;
<% else %>
//var int_room_min = 24 ; //前三分鐘進教室
var int_end_min = 27; //結束倒數時間
if ( str_specianl ! = "" ) 
{
	var int_room_min = 17 ;
}
else
{
	var int_room_min = 24 ;	
}
var int_query_start_room_min = 30;
<% end if %>
//-->
</script>
<script type='text/javascript' src="/lib/javascript/framepage/js/member.js"></script>
<!--內容start-->
       <div id="menber_membox">
	     <%if session("ShareMGM") = "1" then%>
	     <div align="center">
		    <a href="http://www.vipabc.com/program/mail_template/MgmEDM/ShareToFriend/index.asp?stype=Banner" target="_blank"><img src="/program/mail_template/MgmEDM/ShareToFriend/images/banner.jpg" alt="" border="0"></a>
         </div>
		 <%end if%>
         <!--收藏課程start-->
         <div id="aaatest">
         <div id="div_clock"></div>
        </div>
        <!--標題start-->
        <div class="main_mylist">
        <ul>
        <%
        if ( true = bolDebugMode ) then
            response.Write "str_today_session : " & str_today_session & "<br/>"
        end if
		if ( str_today_session <> "") then 
			'表示今日有課
			str_show_today = getWord("now_class")
			
			'--------------- 課程為3.5小時內，才開放"進入教室"按鈕 -------------- start -----------
            str_session_sn = getSessionWholeTime(str_session_sn, 3, 0, CONST_VIPABC_RW_CONN) & ":00"
			'str_session_sn = left(str_session_sn,4)&"/"&right(left(str_session_sn,6),2)&"/"&right(left(str_session_sn,8),2)&"  "&right(left(str_session_sn,10),2)&":30:00"
			int_dcgs_now = DateAdd("n", 210, now()) '現在的時間 加上210分(3.5小時)
			int_dcgs_now = sCDbl(getFormatDateTime(int_dcgs_now,10))
			int_session = sCDbl(getFormatDateTime(str_session_sn,10)) '實際上課的時間
			int_now_time = sCDbl(getFormatDateTime(now(),10)) '現在時間
			if ( true = bolDebugMode ) then
                response.Write "str_session_sn : " & str_session_sn & "<br/>"
                response.Write "int_dcgs_now : " & int_dcgs_now & "<br/>"
                response.Write "int_session : " & int_session & "<br/>"
                response.Write "int_now_time : " & int_now_time & "<br/>"
            end if
			if ( int_dcgs_now >= int_session ) then '是否在開課前3.5小時內，即顯示進入教室按鈕
				bol_room_bout = 1
			end if 
			'--------------- 課程為3.5小時內，才開放"進入教室"按鈕 -------------- end -----------
		else
			str_show_today = getWord("no_class_today")
		end if 
		'response.write ""&int_now_time&"......"&int_session&"<br>"
		'response.write ""&int_dcgs_now&"......"&int_session&"<br>"
		%>
        <li><%=str_show_today%><span class="black"><%=str_today_session%></span>
        <%
        if ( bol_room_bout = 1 ) then 
            if ( true = bolShortSession ) then
                strJS1 = "window.open('/program/aspx/AspxLogin.asp?url_id=7','_blank');"
                strJS2 = "window.open('/program/aspx/AspxLogin.asp?url_id=7','_blank');"
            else
                strJS1 = "location.href='class_notice.asp';"
                strJS2 = "location.href='class.asp';"
            end if
			if ( int_now_time > int_session ) then 
			'超過上課時間，快速進教室按鈕
			'課程三分鐘內進教室，直接進課前注意事項
			'立即进教 室enter_class
			%>
				<span id="span_query_room" class="right_02">
                    <input type="button" name="btn_query_room" id="btn_query_room" value="+ <%=getWord("enter_class")%>" class="btn_1 m-left5" onclick="javascript:<%=strJS1%>"/>
                </span>
			<%
			else
			'進入教室
			%>
                <span id="span_in_room" class="right_02">
                	<input type="button" name="btn_in_room" id="btn_in_room" value="+ <%=getWord("CLASS_3")%>" class="btn_1 m-left5" onclick="javascript:<%=strJS2%>"/>
                </span>
            <%		
			end if 
		end if 
        %>
        <!--
        <span id="span_query_room_js" class="right_02">
            <input type="button" name="btn_query_room" id="btn_query_room_js" value="+ <%=getWord("enter_class")%>" class="btn_1 m-left5" onclick="javascript:<%=strJS1%>"/>
        </span>
        -->

        <%'倒數計時器%>
        <span id="span_in_room_time" class="right bold color_1" ></span>
        </li>
        <li><%=getWord("other_class")%>：
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

	'---客戶資料 COM 物件
    'dim str_now_use_product_sn : str_now_use_product_sn = ""
    Dim obj_member_opt_exe : Set obj_member_opt_exe = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
    Dim int_member_result_exe : int_member_result_exe = obj_member_opt_exe.prepareData(g_var_client_sn, CONST_VIPABC_RW_CONN)
    Dim strPurchaseAccountSn : strPurchaseAccountSn = "" '離目前時間最近的一筆 客戶於PURCHASE的序號

    if (int_member_result_exe = CONST_FUNC_EXE_SUCCESS) then
        strPurchaseAccountSn = obj_member_opt_exe.getData(CONST_LAST_CONTRACT_ACCOUNT_SN) '離目前時間最近的一筆 客戶於PURCHASE的序號
        int_attend_level = obj_member_opt_exe.getData(CONST_CUSTOMER_NOW_LEVEL)                       '取得客戶的等級
		int_client_data_rate = obj_member_opt_exe.getData(CONST_DATA_COMPLETE_RATE) 
        str_now_use_product_sn = obj_member_opt_exe.getData(CONST_CONTRACT_PRODUCT_SN)                  '離目前時間最近的一筆 客戶合約的產品序號
    else
        '錯誤訊息
        Response.Write "member prepareData error = " & obj_member_opt_exe.PROCESS_STATE_RESULT & "<br>"
    end if
	set int_member_result_exe = nothing
	set obj_member_opt_exe = nothing
	
    '**** 取得產品資訊 START ****
    'dim int_client_purchase_product_type, arr_cfg_product_tmp
    int_client_purchase_product_type = 0
    if (str_now_use_product_sn <> "") then
        '取得產品的資訊
        arr_cfg_product_tmp = getProductInfo(str_now_use_product_sn, CONST_VIPABC_RW_CONN)
        if ( IsArray(arr_cfg_product_tmp) ) then
            '取得 客戶購買合約的產品類型 1:一般產品 2:定時定額
            int_client_purchase_product_type = CInt(arr_cfg_product_tmp(10, 0))	
        end if
    end if
    '**** 取得產品資訊 END   ****

    '**** 取得剩餘堂數資訊 START ****
    Dim arrSessionPoints, fltAvailableSession, flt_remain_week_session
    arrSessionPoints =  getCustomerContractSessionPointInfo("", "", strPurchaseAccountSn, CONST_VIPABC_RW_CONN)
    fltAvailableSession = 0
    flt_remain_week_session = 0
    if ( IsArray(arrSessionPoints) ) then
        if ( not isEmptyOrNull(arrSessionPoints(20)) ) then
            '可用堂數
            fltAvailableSession = Round(CDbl(arrSessionPoints(20)), 2)
        end if

        '當週剩餘堂數
        if (Not isEmptyOrNull(arrSessionPoints(21))) then
            flt_remain_week_session = arrSessionPoints(21)
        end if
    end if
    '**** 取得剩餘堂數資訊 END   ****
		%>
       <a href="reservation_class/watch_class.asp"><%=getWord("other_class_timetable")%></a><a href="learning_record.asp"><%=getWord("my_class_record")%></a></li>
        <li>
            <!--可用周期总堂数<span class="red">&nbsp;<%=CDbl(fltAvailableSession) + CDbl(flt_remain_week_session)%></span>，
            剩余周期堂数<span class="red">&nbsp;<%=fltAvailableSession%></span>
            <%
            '定時定額產品
            'if ( CONST_PRODUCT_QUOTA = int_client_purchase_product_type ) then
            %>
            <!--当周期应用堂数<span class="red"><%=flt_remain_week_session%></span>。
            <% 'else %>
            可用周期总堂数<span class="red">&nbsp;<%=CDbl(fltAvailableSession) + CDbl(flt_remain_week_session)%></span>，
            剩余周期堂数<span class="red">&nbsp;<%=fltAvailableSession%></span>。-->
            <% 'end if %>
        </li>
        <%'預訂課程按鈕%>
        <span class="right_02">
        <%if (checkReservationClassQualification(CONST_CLASS_TYPE_ONE_ON_ONE & "," & CONST_CLASS_TYPE_ONE_ON_THREE, CONST_VIPABC_RW_CONN)) then%>
			<input type="button" name="" id="btn_member_reservation_class" value="+ <%=getWord("SHCEDULE_CLASS")%>" class="btn_1 m-left5" onclick="javascript:location.href='reservation_class/class_choose.asp'"/>
        <%end if %>
        </span>
        </ul>
        </div>
        <div class="box_shadow" style="background-position:-80px 0px;"></div> 
        <!--標題end-->
        <!--NEWS標頭-->        
        <!--NEWS標頭end-->  
		<!--
        <div id="divLobbySessionBanner_other">
			<a href="/program/member/reservation_class/lobby_near_order.asp"><img src="/images/ielts_session.gif" border="0" /></a>
			<br />
			<a href="/program/member/reservation_class/lobby_near_order.asp"><img src="/images/BANNER_20120815.jpg" border="0" /></a>
		</div>
		-->
        <!--個人資料標頭-->
        <div class="con_close">

        <div class="con_header_right"><%=getWord("data_complete_rate")%>：<%=int_client_data_rate%>% <a href="profile.asp"><%=getWord("edit_personal_data")%></a><a href="learning_dcgs.asp"><%=getWord("set_DCGS")%></a><!--<a href="#">打印证书</a>--></div>
        <div class="con_header_right_02">
        <a id="a_open_and_close_commend_member_info" class="close_btn01" href="javascript:void(0);" onclick="showCommendLobbyClass();"><%=getWord("Close")%></a></div>
        <div class="clear"></div>
        </div>
         <!--個人資料標頭end-->
         <%
        dim str_client_email
        dim int_company_tel
        dim str_client_member_sql
        dim str_client_dcgs_sql
        dim arr_result
        dim dcgs_result
        dim var_arr
        dim dcgs_arr
        dim int_count
        dim str_industry : str_industry = ""
        dim str_role : str_role = ""
        dim str_interest : str_role = ""
        dim str_focus
		Dim str_cphone_cycode : str_cphone_cycode = "" '手機國碼
		Dim str_cphone_cycode_2 : str_cphone_cycode_2 = "" '第二組手機國碼
		Dim str_htel_cycode : str_htel_cycode = ""	'固定電話國碼
		Dim str_otel_cycode : str_otel_cycode = ""	'辦公室電話國碼  
        Dim int_client_sex
        Dim int_client_level
		'---------add by ryanlin 20100401 修改地址BUG-------start------
		Dim int_province_number '國碼
		Dim int_area_number '區碼
		'---------add by ryanlin 20100401 修改地址BUG-------end  ------

        var_arr = Array(g_var_client_sn)
        str_client_member_sql = "select email,cname,fname,lname,sex,nlevel,birth,client_address," 
		str_client_member_sql = str_client_member_sql & " cphone_code+cphone as cphone "
        str_client_member_sql = str_client_member_sql & ",htel_code,htel,otel_code,otel,clb_cphone_2,focus " 
		str_client_member_sql = str_client_member_sql & ",clb_cphone_cycode,clb_cphone_cycode_2,clb_htel_cycode,clb_otel_cycode, clb_province ,clb_area, clb_region "
 		str_client_member_sql = str_client_member_sql & " from client_basic where sn=@g_var_client_sn "
			
        arr_result = excuteSqlStatementRead(str_client_member_sql, var_arr, CONST_VIPABC_RW_CONN)       
		'response.write g_str_sql_statement_for_debug
		'response.end
          if (isSelectQuerySuccess(arr_result)) then
            '有資料
			
			'手機國碼
			str_cphone_cycode = arr_result(15,0)
			if ( not isEmptyOrNull(str_cphone_cycode) ) then 
				str_cphone_cycode = "+"&str_cphone_cycode
			end if 
			
			'第二組手機國碼
			str_cphone_cycode_2 = arr_result(16,0)
			if ( not isEmptyOrNull(str_cphone_cycode_2) ) then 
				str_cphone_cycode_2 = "+"&str_cphone_cycode_2
			end if 
			
			'固定電話國碼
			str_htel_cycode = arr_result(17,0)
			if ( not isEmptyOrNull(str_htel_cycode) ) then 
				str_htel_cycode = "+"&str_htel_cycode
			end if 
			
			'辦公室電話國碼
			str_otel_cycode = arr_result(18,0)
			if ( not isEmptyOrNull(str_otel_cycode) ) then 
				str_otel_cycode = "+"&str_otel_cycode
			end if 
			
            if ( Ubound(arr_result) >= 0 ) then
                int_client_sex = arr_result(4,0)
                if int_client_sex = 1 then
                    int_client_sex = getWord("male")
                else
                    int_client_sex = getWord("female")
                end if 
                int_client_level = arr_result(5,0)
				
                if (int_client_level <> "") then 
                    int_client_level = convertNumber(int_client_level, CONST_TW_NOR_NUMBER)
				else
				'ryanlin 20100205 判斷為空 show無任何資訊
					int_client_level = getWord("NO_DATA")
                end if 
                '找尋客戶的FOCUS
                
                if isnull(arr_result(14,0) ) or  arr_result(14,0) = "" then
                    str_focus = getWord("no_data") 
                else
                    Select Case arr_result(14,0) 
                           Case 1
                               str_focus= getWord("DCGS_8")
                           Case 2
                               str_focus= getWord("CLASS_VOCABULARY_7")
                           Case 3
                               str_focus= getWord("CLASS_VOCABULARY_8")
                           Case 4
                               str_focus= getWord("DCGS_6")                                                      
                           Case 5
                               str_focus= getWord("DCGS_7")                              
                           Case 6
                               str_focus= getWord("DCGS_9")      
                    End Select
                end if
                
                if  isnull(arr_result(11,0)) or isnull(arr_result(12,0))  then
                    int_company_tel=arr_result(11,0) &"-"& arr_result(12,0)
                else
                    int_company_tel= getWord("no_data")     
                end if
                    'Response.Write int_company_tel
                
			'----------add by ryanlin 20100401 修改地址BUG-------start------
            '20121226 阿捨新增 地址帶入三層地址架構值
            dim strClientAddress : strClientAddress = ""
            dim strValueName : strValueName = "" '地區省分值
            dim intProvinceNumber : intProvinceNumber = trim(arr_result(19,0))
            dim intAreaNumber : intAreaNumber = trim(arr_result(20,0))
            dim intRegionNumber : intRegionNumber = trim(arr_result(21,0))
            dim strPath : strPath = Server.Mappath("/xml/ChinaProvinceAndArea.xml")
	        strValueName = findXmlValueName(strPath, intProvinceNumber, intAreaNumber, intRegionNumber)
			'----------add by ryanlin 20100401 修改地址BUG--------end-----

        dcgs_arr = Array(g_var_client_sn,g_var_client_sn,g_var_client_sn)
      
        'str_client_dcgs_sql = "select client_profile_log.*,lesson_industry.cname from client_profile_log" 
        'str_client_dcgs_sql = str_client_dcgs_sql & " left join lesson_industry on lesson_industry.sn= client_profile_log.profile_sn " 
        'str_client_dcgs_sql = str_client_dcgs_sql & "where client_sn=@g_var_client_sn and profile_type=1" 
    
         str_client_dcgs_sql = "select client_profile_log.profile_type,client_profile_log.profile_sn,lesson_industry.sname from client_profile_log "
         str_client_dcgs_sql =  str_client_dcgs_sql & "left join lesson_industry on lesson_industry.sn= client_profile_log.profile_sn "
         str_client_dcgs_sql =  str_client_dcgs_sql & " where client_sn=@g_var_client_sn_1 and profile_type=1 and edate is null"
         str_client_dcgs_sql =  str_client_dcgs_sql & " union all"
         str_client_dcgs_sql =  str_client_dcgs_sql & " select client_profile_log.profile_type,client_profile_log.profile_sn,lesson_role.sname from client_profile_log"
         str_client_dcgs_sql =  str_client_dcgs_sql & " left join lesson_role on lesson_role.sn= client_profile_log.profile_sn"
         str_client_dcgs_sql =  str_client_dcgs_sql & " where client_sn=@g_var_client_sn_2 and profile_type=2 and edate is null"
         str_client_dcgs_sql =  str_client_dcgs_sql & " union all"
         str_client_dcgs_sql =  str_client_dcgs_sql & " select client_profile_log.profile_type,client_profile_log.profile_sn,cfg_interest.sname from client_profile_log"
         str_client_dcgs_sql =  str_client_dcgs_sql & " left join cfg_interest on cfg_interest.sn= client_profile_log.profile_sn"
         str_client_dcgs_sql =  str_client_dcgs_sql & " where client_sn=@g_var_client_sn_3 and profile_type=3 and edate is null"

        dcgs_result = excuteSqlStatementRead(str_client_dcgs_sql, dcgs_arr, CONST_VIPABC_RW_CONN)       

		if (isSelectQuerySuccess(dcgs_result)) then
            '有資料
            if (Ubound(dcgs_result) >= 0) then
                for int_count = 0  to Ubound(dcgs_result,2)
                   '职业别
                   if dcgs_result(0,int_count) = 1 then
                       if str_industry = "" then
                            str_industry =  dcgs_result(2,int_count)
                        else
                            str_industry = str_industry &"," & dcgs_result(2,int_count)       
                        end if
                   end if 
                   'role:職業別
                   if dcgs_result(0,int_count) = 2 then
                       if str_role = "" then
                            str_role =  dcgs_result(2,int_count) 
                        else
                            str_role = str_role &"," & dcgs_result(2,int_count) 
                        end if
                   end if 
                   'interest:興趣別
                   if dcgs_result(0,int_count) = 3 then
                       if str_interest = "" then
                            str_interest =  dcgs_result(2,int_count) 
                        else
                            str_interest = str_interest &"," & dcgs_result(2,int_count) 
                        end if
                   end if 
                next
            else
                '錯誤訊息
            '	Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
            ''	Response.End()
	            'call sub alertGo
	           'Call alertGo(getWord("role"), "", CONST_ALERT_THEN_GO_BACK)
	           ' Response.End       
	           str_industry=getWord("no_data") 
	           str_role=getWord("no_data") 
	           str_interest=getWord("no_data") 
            end if
        'else            
        end if
               %>
                <div id="div_commend_member_info">
                     <table cellpadding="0" cellspacing="0" class="tabpf">
                     <tr>
                     <td class="t01"><%=getWord("email_account")%></td>
                     <td colspan="3" class="c02"><%=arr_result(0,0) %></td>
                     <td class="t01"><%=getWord("password")%></td>
                     <td class="c01">******<a href="profile_password.asp"><%=getWord("modify_password")%></a></td>
                     </tr>
                     <tr>
                       <td class="t01"><%=getWord("CONTRACT_35")%></td>
                       <td class="c01"><%=arr_result(1,0) %></td>
                       <td class="t01"><%=getWord("ENAME")%></td>
                       <td class="c01"><%=arr_result(2,0) %></td>
                       <td class="t01"><%=getWord("fNAME")%></td>
                       <td class="c01">
                        <% if isnull(arr_result(3,0)) or arr_result(3,0)="" then 
                          response.write getWord("no_data") 
                        else 
                          response.write arr_result(3,0) 
                        end if 
                        %>
                       </td>
                     </tr>
                     <tr>
                       <td class="t01"><%=getWord("rank")%></td>
                       <td class="c01"><%=int_client_level %>
					   <%
						if int_client_level <> "" then  
							response.write getWord("rank") 
						end if 
					   %><!--<a href="#">APPLY_FOR_NEW_LEVEL</a>--></td>
                       <td class="t01"><%=getWord("sex")%></td>
                       <td class="c01"><%=int_client_sex %></td>
                       <td class="t01"><%=getWord("BIRTHDAY")%></td>
                       <td class="c01"><%=arr_result(6,0) %></td>
                     </tr>
                     <tr>
                       <td class="t01"><%=getWord("communication_address")%></td>
                       <td colspan="5" class="c03">
                         <%if isnull(arr_result(7,0)) or arr_result(7,0)="" then 
                          response.write getWord("no_data") 
                        else 
                          response.write strValueName&arr_result(7,0)
                        end if 
                        %>                     
                        </td>
                       </tr>
                     <tr>
                       <td class="t01"><%=getWord("CONTACT_US_MOBIL_PHONE")%>1</td>
                       <td class="c01"><%=str_cphone_cycode&arr_result(8,0) %></td>
                       <td class="t01"><%=getWord("CONTACT_US_MOBIL_PHONE")%>2</td>
                       <td class="c01">
                       <%if isnull(arr_result(13,0)) or arr_result(13,0)="" then 
                          response.write getWord("no_data") 
                        else 
                          response.write str_cphone_cycode_2&arr_result(13,0)
                        end if 
                        %>
                       </td>
                       <td class="t01"><%=getWord("DCGS_5")%></td>
                       <td class="c01"><%=str_focus %></td>
                     </tr>
                     <tr>
                       <td class="t01"><%=getWord("home_phone")%></td>
                       <td class="c01"><%=str_htel_cycode&arr_result(9,0) %>-<%=arr_result(10,0) %></td>
                       <td class="t01"><%=getWord("office_phone")%></td>
                       <td colspan="3" class="c02"><%=str_otel_cycode&int_company_tel %>
                       </td>
                       </tr>
                     <tr>
                       <td class="t01"><%=getWord("ROLE")%></td>
                       <td colspan="5" class="c03"><%=str_industry %></td>
                       </tr>
                     <tr>
                       <td class="t01"><%=getWord("position")%></td>
                       <td colspan="5" class="c03"><%=str_role %></td>
                       </tr>
                     <tr>
                       <td class="t01"><%=getWord("DCGS_12")%></td>
                       <td colspan="5" class="c03"><%=str_interest %></td>
                       </tr>
                     </table>
                 </div>
               <%     
            else
                '錯誤訊息
            '	Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
            ''	Response.End()
	            'call sub alertGo
	           Call alertGo(getWord("data_error")&" "&getWord("email_error"), "", CONST_ALERT_THEN_GO_BACK)
	           Response.End
            end if
        end if               
          %>
         <!---個人資料---->
         <!---個人資料end---->
         <!--我的收藏標頭-->
        <!-- <div class="con_close">
           <div class="con_header_left">我的收藏</div>
           <div class="con_header_right"><a href="#">编辑我的收藏</a></div>
           <div class="con_header_right_02"><a href="#">关闭 ─</a></div>
           <div class="clear"></div>
         </div>-->
         <!--我的收藏標頭end-->
         <!--課程回顧-->
         <!--選單-->
        <!--<div class="main_memtitle">
           <div class="mune_01">课程回顾</div>
           <div ><a href="member_2.html" class="mune_02">最爱顾问名单</a></div>
           <div ><a href="member_3.html" class="mune_03">精选电子报</a></div>
         </div>
         <div class="clear"></div>-->
         <!--選單end-->
         <!--列表內容-->
        <!-- <div class="menber_content_list">
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
         <!-- <div class="main_storage"> <span class="left_member">您收藏的课程回顾共 4 堂</span> <span class="right_page"><a href="#">上一頁</a><span class="right_text">1</span><a href="#">2</a><a href="#">3</a><a href="#">4</a><a href="#">5</a><a href="#">6</a><a href="#">7</a><a href="#">8</a><a href="#">下一頁</a></span>
             <div class="clear"></div>
         </div>-->
         <!--儲存選單emd-->
         <!--列表內容end-->
         <!--課程回顧end-->
         <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
       </div>
       <!--內容end-->