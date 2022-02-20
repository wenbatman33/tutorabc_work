<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/member/reservation_class/include/include_prepare_and_check_data.inc"-->
<!--#include virtual="/lib/include/class/DataOutputEncode.asp" -->
<!--#include virtual="/program/class/lobby/CourseUtility.asp" -->
<!--標題start-->
<body lang="zh">
<link href="/css/session_20120807.css" rel="stylesheet" type="text/css" />
<!--標題end-->
<%

'被 /program/member/reservation_class/lobby_near_order.asp ajax 動態載入
'被 /program/member/reservation_class/lobby_near_cancel.asp ajax 動態載入

Dim bolDebugMode : bolDebugMode = false '正式改為false

Dim str_search_topic : str_search_topic = trim(unEscape(Request("lobby_topic"))) '搜尋條件: 主題
Dim str_search_con : str_search_con = trim(Request("con_sn")) '搜尋條件: 顧問編號
Dim str_search_lev : str_search_lev = trim(Request("lobby_suit_lev")) '搜尋條件: 適合等級
Dim str_search_time : str_search_time = trim(Request("lobby_time")) '搜尋條件: 課程時段
Dim strImageUrl : strImageUrl = "" '顧問大頭照圖片路徑
Dim str_suit_lev : str_suit_lev = "" '適合的等級
Dim str_lobby_condition_topic : str_lobby_condition_topic = "" 'lobby_session.topic where 條件
Dim str_lobby_condition_con : str_lobby_condition_con = "" 'lobby_session.consultant where 條件
Dim str_lobby_condition_lev : str_lobby_condition_lev = "" 'lobby_session.lev where 條件
Dim str_lobby_condition_time : str_lobby_condition_time = "" 'lobby_session.session_time where 條件
Dim str_lobby_condition_lobby_type : str_lobby_condition_lobby_type = "" 'lobby_session.lobby_type where 條件
Dim str_lobby_condition_datetime : str_lobby_condition_datetime = "" 'lobby_session.session_date & lobby_session.session_time where 條件
Dim str_attend_condition_con : str_attend_condition_con = "" 'client_attend_list.attend_consultant where 條件
Dim str_attend_condition_time : str_attend_condition_time = "" 'client_attend_list.attend_sestime where 條件
Dim str_attend_condition_datetime : str_attend_condition_datetime = "" 'client_attend_list.attend_date & client_attend_list.attend_sestime where 條件
Dim str_lobby_session_introduction : str_lobby_session_introduction = "" 'lobby_session.introduction 條件
Dim str_lobby_session_topic : str_lobby_session_topic = "" 'lobby_session.topic 條件
Dim str_from_other_page_con_sn : str_from_other_page_con_sn = trim(Request("from_other_page_con_sn")) '搜尋頁面和推薦大會堂也可以訂課 若有值則輸出顧問時要打勾
Dim strLastFirstSessionHour : strLastFirstSessionHour = 0 '檢查客戶是否是首次上課且有無預約First Session，若有預約的話只能訂一小時後的課
Dim strSql, arr_lobby_data, arrParam
Dim intClientSn, i, str_arr_tmp, strCancelClassDatetime, strAlreadyOrderTitle
dim strTotalSession : strTotalSession = getRequest("strTotalSession", CONST_DECODE_ESCAPE) '更新剩餘堂數
dim strReservedCalOrList : strReservedCalOrList = "list" '確認訂課後 判斷載入頁面
Dim objDencode : set objDencode = new DataOutputEncode

'20120316 阿捨新增 無限卡產品判斷	
Dim strUnlimited : strUnlimited = getSession("unlimited", CONST_DECODE_NO)
if ( isEmptyOrNull(strUnlimited) ) then
    Dim obj_member_opt_exe : Set obj_member_opt_exe = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
    Dim int_member_result_exe : int_member_result_exe = obj_member_opt_exe.prepareData(g_var_client_sn, CONST_VIPABC_RW_CONN)
    if ( int_member_result_exe = CONST_FUNC_EXE_SUCCESS ) then        
        strNowUseProductSn = obj_member_opt_exe.getData(CONST_CONTRACT_PRODUCT_SN) '離目前時間最近的一筆 客戶合約的產品序號
    else
	    '錯誤訊息
	    'response.Write "member prepareData error = " & obj_member_opt_exe.PROCESS_STATE_RESULT & "<br>"
    end if
    if ( Not isEmptyOrNull(strNowUseProductSn) ) then
        if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (strNowUseProductSn & ",")) > 0 ) then
            session("unlimited") = "true" 
            strUnlimited = session("unlimited")
        end if
    end if
end if

'取得客戶編號
if ( Not isEmptyOrNull(g_var_client_sn) ) then
    intClientSn = g_var_client_sn
end if


'**** 2012-11-08增加 預約測試 排除預約測試之前的時間  Johnny
Set rs = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp") '建立物件
set course_ = new CourseUtility
re_valArr = course_.getFirstSessionTestTime ( rs , intClientSn)
if "N"=re_valArr(0) then
	if (true = isDate(re_valArr(1))) then
		int_diff_orderEnvTest_minute = DateDiff("n", now(), re_valArr(1))
	end if
end if		
rs.close()
set rs = nothing
set course_ = nothing
'**** 2012-11-08 End
''設定 lobby_session 時間條件 (大會堂只允許訂24時後的課)
'g_str_lobby_condition_datetime = " AND ((CONVERT(VARCHAR, lobby_session.session_date, 111) = CONVERT(VARCHAR, GETDATE()+1, 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_
	                                '"      OR (CONVERT(VARCHAR, lobby_session.session_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

'設定 lobby_session 時間條件 (大會堂允許開課前5分鐘都可以訂)
str_lobby_condition_datetime = " AND ((CONVERT(VARCHAR, lobby_session.session_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND lobby_session.session_time = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
                                " OR (CONVERT(VARCHAR, lobby_session.session_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_ 
	                            "      OR (CONVERT(VARCHAR, lobby_session.session_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

''設定 client_attend_list 時間條件 (大會堂只允許訂24時候的課)
'g_str_attend_condition_datetime = " AND ((CONVERT(VARCHAR, client_attend_list.attend_date, 111) = CONVERT(VARCHAR, GETDATE()+1, 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
	                                '"      OR (CONVERT(VARCHAR, client_attend_list.attend_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

'設定 client_attend_list 時間條件 (大會堂允許開課前5分鐘都可以訂)
str_attend_condition_datetime = " AND ((CONVERT(VARCHAR, client_attend_list.attend_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND client_attend_list.attend_sestime = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
                                " OR (CONVERT(VARCHAR, client_attend_list.attend_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
	                            "      OR (CONVERT(VARCHAR, client_attend_list.attend_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

'大會堂主題 條件設定
if (Not isEmptyOrNull(str_search_topic)) then		
    str_lobby_condition_lobby_type = " AND (lobby_session.lobby_type = '"&str_search_topic&"') "
end if
    
'顧問編號 條件設定
if (Not isEmptyOrNull(str_search_con)) then
    str_lobby_condition_con = " AND (lobby_session.consultant = '"&str_search_con &"')  "
    str_attend_condition_con = " AND (client_attend_list.attend_consultant = '"&str_search_con &"')  "
end if

'大會堂適合的等級 條件設定
if (Not isEmptyOrNull(str_search_lev)) then
    str_lobby_condition_lev = " AND (lobby_session.lev LIKE '%, " & str_search_lev & ",%')  "
end if

'大會堂開課時段 條件設定
if (Not isEmptyOrNull(str_search_time)) then
    str_lobby_condition_time = " AND (lobby_session.session_time = '"&str_search_time&"')  "
    str_attend_condition_time = " AND (client_attend_list.attend_sestime = '"&str_search_time&"')  "
end if

'TODO: 大會堂的標題 判斷語系
if (session("language") = CONST_LANG_CN) then
    str_lobby_session_topic = " lobby_session.topic_cn "
else
    str_lobby_session_topic = " lobby_session.topic "
end if

'TODO: 大會堂內容簡介 判斷語系
if (session("language") = CONST_LANG_CN) then
    str_lobby_session_introduction = " lobby_session.introduction_cn "
else
    str_lobby_session_introduction = " lobby_session.introduction "
end if

''
'設定 lobby_session 時間條件 (大會堂只允許訂24時後的課)
'g_str_lobby_condition_datetime = " AND ((CONVERT(VARCHAR, lobby_session.session_date, 111) = CONVERT(VARCHAR, GETDATE()+1, 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_
									'"      OR (CONVERT(VARCHAR, lobby_session.session_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

'設定 lobby_session 時間條件 (大會堂允許開課前5分鐘都可以訂)
str_lobby_condition_datetime = " AND ((CONVERT(VARCHAR, lobby_session.session_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND lobby_session.session_time = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
								" OR (CONVERT(VARCHAR, lobby_session.session_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_ 
								"      OR (CONVERT(VARCHAR, lobby_session.session_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

''設定 client_attend_list 時間條件 (大會堂只允許訂24時候的課)
'g_str_attend_condition_datetime = " AND ((CONVERT(VARCHAR, client_attend_list.attend_date, 111) = CONVERT(VARCHAR, GETDATE()+1, 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
									'"      OR (CONVERT(VARCHAR, client_attend_list.attend_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

'設定 client_attend_list 時間條件 (大會堂允許開課前5分鐘都可以訂)
str_attend_condition_datetime = " AND ((CONVERT(VARCHAR, client_attend_list.attend_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND client_attend_list.attend_sestime = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
								" OR (CONVERT(VARCHAR, client_attend_list.attend_date, 111) = CONVERT(VARCHAR, GETDATE(), 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
								"      OR (CONVERT(VARCHAR, client_attend_list.attend_date, 111) > CONVERT(VARCHAR, GETDATE(), 111))) "

'20121004 Create 計算總筆數 Start 
strSql = " SELECT "
'--大會堂的編號 9
strSql = strSql & " COUNT(lobby_session.sn)  "
strSql = strSql & " FROM ( "
'--有開課的大會堂及顧問的資料(>=今天)
strSql = strSql & "   SELECT lobby_session.sn "
strSql = strSql & "   FROM lobby_session "
strSql = strSql & "   WHERE (lobby_session.valid = 1 AND lobby_session.sn <> 2622 ) " 
strSql = strSql & " AND ( lobby_session.SessionPeriod = 45 OR lobby_session.SessionPeriod IS NULL ) " 
strSql = strSql & str_lobby_condition_datetime & str_lobby_condition_topic & str_lobby_condition_con & str_lobby_condition_lev & str_lobby_condition_time & str_lobby_condition_lobby_type
strSql = strSql & " AND lobby_session.session_date >= CONVERT(varchar(10), GETDATE(),111) AND lobby_session.session_date <= CONVERT(varchar(10), (GETDATE()+7),111)  "
'--加入判斷品牌非JR
strSql = strSql & " AND lobby_session.BrandType <> 2 "
strSql = strSql & " ) AS lobby_session  "
arr_lobby_data = excuteSqlStatementRead(strSql, "", CONST_VIPABC_RW_CONN)

if ( true = bolDebugMode ) then
	response.Write ("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
end if

dim intTotal : intTotal = 0 '總筆數
if ( isSelectQuerySuccess(arr_lobby_data) ) then
	if ( Ubound(arr_lobby_data) >= 0 ) then
		intTotal = arr_lobby_data(0,0)
	end if
end if
'<!-- 20121004 Create 計算總筆數 End -->
	
dim intPageCount : intPageCount = 8 '每頁筆數
dim intNowPage : intNowPage = getRequest("page", CONST_DECODE_ESCAPE) '目前頁面
dim intTopNum : intTopNum = "" '開頭筆數
dim intBottomNum : intBottomNum = "" '結尾筆數
dim intTotalPage : intTotalPage = 0 '總頁數
intTotalPage = ( intTotal / intPageCount ) '總頁數無條件進位
if ( InStr(intTotalPage,".") = 0 ) then
	intTotalPage = intTotalPage 
else 
	intTotalPage = sCInt(Left(cstr(intTotalPage),InStr(1,cstr(intTotalPage),".")-1))+ 1 
end if

'預設第一頁
if ( true = bolDebugMode ) then
	response.Write "內層_總頁數intTotalPage : " &intTotalPage&"<br/>"
	response.Write "內層_總筆數intTotal : " &intTotal&"<br/>"
end if

if ( isEmptyOrNull(intNowPage) ) then
    intNowPage = 1 
end if

intBottomNum = intNowPage * intPageCount
intTopNum = intBottomNum - intPageCount + 1

if ( true = bolDebugMode ) then
	response.Write "內層_此頁最後一筆intBottomNum : " &intBottomNum& "<br/>"
	response.Write "內層_此業第一筆intTopNum : " &intTopNum& "<br/>"
	response.Write "內層_目前頁數intNowPage : " &intNowPage&"<br/>"
end if

dim intSessionSn : intSessionSn = 0
dim intSessionStatus : intSessionStatus = 0
dim intSessionSTime : intSessionSTime = ""
dim strSessionTopic : strSessionTopic = ""
dim strSessionSDate : strSessionSDate = ""
dim strConsultantFName : strConsultantFName = ""
dim strConsultantLName : strConsultantLName = ""
dim strConsultantEName : strConsultantEName = ""
dim intMaterialSn : intMaterialSn = ""
dim strSessionIntro : strSessionIntro = ""
dim strSessionApplyLv : strSessionApplyLv = ""
dim intSessionMaxNum : intSessionMaxNum = 0
dim intSessionNowNum : intSessionNowNum = 0
dim intSessionMemberLv : intSessionMemberLv = 0
dim intClientAttendSn : intClientAttendSn = ""
dim intClientAttendType : intClientAttendType = 0
dim intSessionType : intSessionType = 0
dim intI : intI = 0

strSql = " SELECT * FROM ( "
strSql = strSql & " SELECT "
'strSql = " SELECT top 50 "
'--開課日期 0
strSql = strSql & " CONVERT(VARCHAR, lobby_session.sdate, 111) AS sdate, "
'--開課時間 1
strSql = strSql & " lobby_session.stime, "
'--顧問名字 2
strSql = strSql & " lobby_session.ename, "
'--大會堂的標題 3
strSql = strSql & str_lobby_session_topic & ", "
'--大會堂內容簡介 4
strSql = strSql & str_lobby_session_introduction & ", "
'--大會堂session_sn 5
strSql = strSql & " lobby_session.dcode, "
'--此時段客戶是否有訂位(訂的是一對一或一對多或大會堂) 6
strSql = strSql & " CASE ISNULL(client_order_lobby.dt, 0) WHEN 0 THEN 0 ELSE 1 END AS status, "
'--該堂大會堂的等級 7
strSql = strSql & " ISNULL(CAST(client_order_lobby.attend_level AS varchar(5)), 0) AS lev, "
'--大會堂人數上限 8
strSql = strSql & " ISNULL(lobby_session.limit, 0) AS limit,  "
'--大會堂的編號 9
strSql = strSql & " lobby_session.sn AS session_sn,  "
'--已經預定大會堂的人數 10
'strSql = strSql & " ISNULL(lobby_order_count.snum, 0) AS snum, " ' TODO : 已達上限人數判斷，submit之後再判斷
'--大會堂適合的等級 11
strSql = strSql & " lobby_session.lev as apply_lev, "
'顧問的大頭照 hex code 12
'strSql = strSql & " lobby_session.con_img, "
'顧問的編號 13
strSql = strSql & " lobby_session.con_sn, "
'client_attend_list.sn 14
strSql = strSql & " client_order_lobby.client_attend_list_sn, "
'--顧問名字 fname 15
strSql = strSql & " lobby_session.basic_fname, "
'--顧問名字 lname 16
strSql = strSql & " lobby_session.basic_lname, "
'--客戶已預定的課程類型 lname 17
strSql = strSql & " client_order_lobby.attend_livesession_types, "
'--課程教材 18  <!--Allison 新增-->
strSql = strSql & " lobby_session.material , "
'--新增session_type判別一對一/一對多
strSql = strSql & " client_order_session.session_type , "
'--新增special_sn
strSql = strSql & " client_order_lobby.special_sn , "
'--分頁語法
strSql = strSql & " ROW_NUMBER() OVER (ORDER BY lobby_session.sdate, lobby_session.stime) AS row "
strSql = strSql & " FROM ( "
'--有開課的大會堂及顧問的資料(>=今天) <!-- 新增教材-->
strSql = strSql & " SELECT lobby_session.sn, lobby_session.material, CONVERT(varchar, lobby_session.session_date, 111) AS sdate, lobby_session.session_time AS stime,  "
strSql = strSql & " isnull(con_basic.basic_fname, '') + ' ' + isnull(con_basic.basic_lname,'') AS ename, " & str_lobby_session_topic & ", " & str_lobby_session_introduction & ", "
strSql = strSql & " CONVERT(VARCHAR, lobby_session.session_date, 112)+ RIGHT('0' + CAST(lobby_session.session_time AS VARCHAR(2)),2)+ CAST(lobby_session.sn AS VARCHAR(5)) AS dcode,	lobby_session.limit, lobby_session.lev, " ''''
'strSql = strSql & " ISNULL(con_basic.img,'') AS con_img, "
strSql = strSql & " con_basic.con_sn, isnull(con_basic.basic_fname, '') AS basic_fname, isnull(con_basic.basic_lname,'') AS basic_lname, "
'--新增datetime
strSql = strSql & " CONVERT(VARCHAR, lobby_session.session_date, 112) + RIGHT('0' + CAST(lobby_session.session_time AS VARCHAR(2)), 2) AS date_time "
strSql = strSql & " FROM lobby_session "
strSql = strSql & " LEFT JOIN con_basic ON lobby_session.consultant = con_basic.con_sn "
strSql = strSql & " WHERE (lobby_session.valid = 1 and lobby_session.sn <> 2622  "
strSql = strSql & " AND ( lobby_session.SessionPeriod = 45 OR lobby_session.SessionPeriod IS NULL ) "  
'--加入判斷品牌非JR
strSql = strSql & " AND lobby_session.BrandType <> 2 "
'--加入只顯示七天大會堂課程 若下面註解拿掉，上面最後的括號要拿掉
strSql = strSql & " AND lobby_session.session_date >= CONVERT(varchar(10), GETDATE(),111) AND lobby_session.session_date <= CONVERT(varchar(10), (GETDATE()+7),111) ) " 
strSql = strSql &   str_lobby_condition_datetime & str_lobby_condition_topic & str_lobby_condition_con & str_lobby_condition_lev & str_lobby_condition_time & str_lobby_condition_lobby_type
strSql = strSql & " ) AS lobby_session "
'--找出所有大會堂的訂課人數
'strSql = strSql & " LEFT JOIN ( "
''strSql = strSql & " SELECT DISTINCT ISNULL(attend_level,0) AS attend_level, COUNT(sn) AS snum FROM client_attend_list "
'strSql = strSql & " SELECT DISTINCT special_sn, COUNT(sn) AS snum FROM client_attend_list "
''strSql = strSql & " WHERE (valid = 1)  " & str_attend_condition_con & str_attend_condition_time
'strSql = strSql & " WHERE (valid = 1)  AND attend_level > 12 " & str_attend_condition_con & str_attend_condition_time
''strSql = strSql & " GROUP BY attend_level, attend_date "
'strSql = strSql & " GROUP BY special_sn "
''strSql = strSql & " HAVING (ISNULL(attend_level,0) > 12)  "
'strSql = strSql & " ) AS lobby_order_count "
'strSql = strSql & " ON lobby_session.sn = lobby_order_count.special_sn "
'--找出客戶有訂位的大會堂
strSql = strSql & " LEFT JOIN ( "
strSql = strSql & " SELECT CONVERT(varchar, attend_date, 112) + RIGHT('0'+CAST(attend_sestime AS varchar(2)),2) AS dt, ISNULL(attend_level,0) AS attend_level, special_sn, client_attend_list.sn AS client_attend_list_sn, client_attend_list.attend_livesession_types "
strSql = strSql & " FROM client_attend_list "
strSql = strSql & " WHERE (client_sn = @client_sn) AND (valid = 1) " & str_attend_condition_datetime & str_attend_condition_con & str_attend_condition_time
strSql = strSql & " ) AS client_order_lobby "
'strSql = strSql & " ON client_order_lobby.attend_level = lobby_session.sn "
strSql = strSql & " ON client_order_lobby.special_sn = lobby_session.sn "
'--找出客戶有訂位的一對一、一對多
strSql = strSql & " LEFT JOIN ( "
'strSql = strSql & " SELECT CONVERT(varchar, attend_date, 112) + RIGHT('0'+CAST(attend_sestime AS varchar(2)),2) AS date_time, ISNULL(attend_level,0) AS attend_level, client_attend_list.sn AS client_attend_list_sn, client_attend_list.attend_livesession_types AS session_type  "
strSql = strSql & " SELECT CONVERT(varchar, attend_date, 112) + RIGHT('0'+CAST(attend_sestime AS varchar(2)),2) AS date_time, client_attend_list.sn AS client_attend_list_sn, client_attend_list.attend_livesession_types AS session_type  "
strSql = strSql & " FROM client_attend_list "
strSql = strSql & " WHERE (client_sn = @clientsn ) AND (valid = 1) AND ( attend_level <= 12 ) " & str_attend_condition_datetime & str_attend_condition_con & str_attend_condition_time
strSql = strSql & " ) AS client_order_session "
strSql = strSql & " ON client_order_session.date_time = lobby_session.date_time "
'LEFT(lobby_session.dcode, 10) = client_order_lobby.dt AND client_order_lobby.attend_level = lobby_session.sn "
'strSql = strSql & " ORDER BY lobby_session.sdate, lobby_session.stime 
strSql = strSql & " ) a "
strSql = strSql & " WHERE row BETWEEN @top AND @bottom "
strSql = strSql & " ORDER BY a.sdate, a.stime "
arrParam = Array(intClientSn, intClientSn, intTopNum, intBottomNum)
'arr_lobby_data = excuteSqlStatementRead(strSql, arrParam, CONST_VIPABC_RW_CONN)
Set objLobbySessionData = excuteSqlStatementReadEach(strSql, arrParam, CONST_VIPABC_R_CONN)

if ( true = bolDebugMode ) then
    response.Write(objLobbySessionData.eof & "<br>")
	response.Write ("錯誤訊息sql for 一周大會堂顧問 : " & g_str_sql_statement_for_debug & "<br>")
    response.write "查詢條件str_lobby_condition_lev : " & str_lobby_condition_lev & "<br/>"
    response.write "查詢條件str_lobby_condition_topic : " & str_lobby_condition_topic & "<br/>"
    response.write "查詢條件str_lobby_condition_con : " & str_lobby_condition_con & "<br/>"
    response.Write "str_search_lev : " & str_search_lev & "<br/>"
	response.Write "str_search_topic : " & str_search_topic & "<br/>"
	response.Write "str_search_con : " & str_search_con & "<br/>"
end if

dim strMinDate : strMinDate = ""
dim strMaxDate : strMaxDate = ""
strMinDate = DATE()
strMinDate = getFormatDateTime(strMinDate, 5)
strMaxDate = DateAdd("d",5,strMinDate)
strMaxDate = getFormatDateTime(strMaxDate, 5)
 'Function getOtherNewSession(ByVal p_intClientSn, ByVal p_strSdate, ByVal p_strEdate, ByVal p_intDeBugMode, ByVal p_intConnectType)
'/****************************************************************************************
'描述      ：取得此客戶訂的新課程類型 包括10mins大會堂(98), 20mins大會堂(97)
'傳入參數  ：
'[p_intClientSn:Integer] client_sn
'[p_strSdate:String] 開始日期
'[p_strEdate:String] 結束日期
'[p_intDeBugMode:Integer] 除錯模式 開:1 關:0
'[p_intConnectType:Integer] 連線字串
'回傳值    ：[objCfgSession] 課程資訊
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) : ArthurWu 2013/05/24 Created
'*****************************************************************************************/
 dim objNewSessionInfo : objNewSessionInfo = null
 set objNewSessionInfo = getOtherNewSession(intClientSn, strMinDate, strMaxDate, 0, CONST_VIPABC_R_CONN)    
%>
<!--課程列表start-->
<div class="session_booking_body">
    <div class="session_booking_list_h">
        <div class="l_c clearfix">
            <div class="l_r r_01">预约</div>
            <div class="l_r r_02">顾问 / 日期</div>
            <div class="l_r r_03">课程主题</div>
            <div class="l_r r_04">适用程度</div>
        </div>
    </div>
    <% '查詢大會堂課程 有資料符合就顯示資料 上面標頭不管有無值都要顯示 所以放外面
        if ( not objLobbySessionData.eof ) then
     %>
    <ul class="class_list_b">
        <%
            while not objLobbySessionData.eof

		    intSessionSn = objLobbySessionData("session_sn")
            intSessionStatus = objLobbySessionData("status")
            intSessionSTime = objLobbySessionData("stime")
            strSessionTopic = objLobbySessionData("topic_cn")
            strSessionSDate = objLobbySessionData("sdate")
            strConsultantFName = objLobbySessionData("basic_fname")
            strConsultantLName = objLobbySessionData("basic_lname")
            strConsultantEName = objLobbySessionData("ename")
            intMaterialSn = objLobbySessionData("material")
            strSessionIntro = objLobbySessionData("introduction_cn")
            strSessionApplyLv = objLobbySessionData("apply_lev")
            intSessionMaxNum = objLobbySessionData("limit")
            'intSessionNowNum = objLobbySessionData("snum")
            intSessionMemberLv = objLobbySessionData("lev")
            intClientAttendSn = objLobbySessionData("client_attend_list_sn")
            intClientAttendType = objLobbySessionData("attend_livesession_types")
            intSessionType = objLobbySessionData("session_type") '新增 判斷一對一/一對多

            strSessionTopic = replace(strSessionTopic, """","")

		    if ( Not isHardCodeComplied("lobby_session_taiwan_class", intSessionSn) ) then '非台灣相關課程
			    '檢查客戶是否是首次上課且有無預約First Session，若有預約的話只能訂一小時後的課
			    if ( isDate(str_last_first_session_datetime) ) then
				    if ( DateDiff("d", DateValue(str_last_first_session_datetime), DateValue(trim(strSessionSDate))) = 0 ) then
				        strLastFirstSessionHour = Hour(str_last_first_session_datetime) + 1
                    end if
                end if

			    '若是首次上課的話且有預約First Session的話，只能訂一小時後的課
			    if ( sCInt(intSessionSTime) >= strLastFirstSessionHour ) then

                'Create 20121009 若為雙數行，變色 。0~7故雙數為1,3,5,7 --Start --by Allison
                if ( intI mod 2 = 1 ) then
                    response.Write "<div class='even'>"
                    strEvenTag = "even"
                else 
                    response.Write "" 
                    strEvenTag = "" '把雙數tag清空
                end if
        %>
        <li>
            <%
            bolDisabled = false
            ' Create 20121024 一對一/一對多 不可訂課 by Allison
            'if ( 4 = sCint(intSessionType) OR 7 = sCint(intSessionType) OR 14 = sCint(intSessionType) ) then
                strDisabled = "disabled"
                strCannotReserved = "ordered " 'css
                bolDisabled = true
                if ( 4 = sCint(intSessionType) OR 14 = sCint(intSessionType) ) then
                    'strNotReservedDetail = "title = '您已预订同时段一对多课程。'"
                    strNotReservedDetail = "title = '您已预订同时段小班制课程。'"
                elseif ( 7 = sCint(intSessionType) ) then
                    strNotReservedDetail = "title = '您已预订同时段一对一课程。'"
                else
                    strNotReservedDetail1 = ""
                    bolHaving = false
                    strSessionWord = ""
                    bolShortSession = false
                    if ( not objNewSessionInfo.eof ) then
                        while not objNewSessionInfo.eof
                            ' : 30~ : 15
                            strShortSessionSTime = objNewSessionInfo("SessionStartTime")
                            strShortSessionETime = objNewSessionInfo("SessionEndTime")
                            strShortSessionSTime = getFormatDateTime(strShortSessionSTime, 1)
                            strShortSessionETime = getFormatDateTime(strShortSessionETime, 1)
                            intNewSessionType = objNewSessionInfo("attend_livesession_types")
                            if ( 98 = sCint(intNewSessionType) ) then
                                strSessionWord = "10mins大会堂"
                            elseif ( 97 = sCint(intNewSessionType) ) then
                                strSessionWord = "20mins大会堂"
                            end if

                            if ( true = bolDebugMode ) then
                                response.write "intNewSessionType : " & intNewSessionType &"<br/>"
                                response.write "strShortSessionSTime : " & strShortSessionSTime &"<br/>"
                                response.write "strShortSessionETime : " & strShortSessionETime &"<br/>"
                            end if

                            str_class_datetime = getFormatDateTime(strSessionSDate & " " & intSessionSTime &":30", 1) '1 = YYYY/MM/DD HH:MM    (HH 為24小時制)
                            strClassDateETime = DateAdd("n", 45, str_class_datetime)
                            strClassDateETime = getFormatDateTime(strClassDateETime, 1) '1 = YYYY/MM/DD HH:MM    (HH 為24小時制)
                            'Function isDateTimeMixed(ByVal p_strStartTime1, ByVal p_strEndTime1, ByVal p_strStartTime2, ByVal p_strEndTime2, ByVal p_intDebugMode)
                            if ( true = isDateTimeMixed(str_class_datetime, strClassDateETime, strShortSessionSTime, strShortSessionETime, 0) ) then
                                strNotReservedDetail1 = strNotReservedDetail1 & "您已预订 "& strShortSessionSTime & " ~ " & strShortSessionETime &" [" & strSessionWord & "]。" & "&#10;"
                                bolHaving = true
                                bolShortSession = true
                            end if

                            ' : 00~ : 59
                            strClassWholeDateSTime = strSessionSDate & " " & intSessionSTime &":00"
                            strClassWholeDateSTime = getFormatDateTime(strClassWholeDateSTime, 1)
                            strClassWholeDateETime = DateAdd("n", 59, strClassWholeDateSTime)
                            strClassWholeDateETime = getFormatDateTime(strClassWholeDateETime, 1)
                            if ( true = isDateTimeMixed(strClassWholeDateSTime, strClassWholeDateETime, strShortSessionSTime, strShortSessionETime, 0) ) then
                                if ( true = bolHaving ) then
                                else
                                    strNotReservedDetail1 = strNotReservedDetail1 & "您已预订 "& strShortSessionSTime & " ~ " & strShortSessionETime &" [" & strSessionWord & "]。" & "&#10;"
                                end if
                            end if
                        objNewSessionInfo.movenext
                        wend
                        objNewSessionInfo.movefirst
                        strNotReservedDetail = "title = '"&strNotReservedDetail1&"'"
                    end if
                    if ( true = bolShortSession ) then
                    else
                        strDisabled = ""
                        strCannotReserved = ""
                        strNotReservedDetail = ""
                    end if
                end if
                strEvenTag = ""
            'else
                'strDisabled = ""
                'strCannotReserved = ""
                'strNotReservedDetail = ""
            'end if
            ' Create 20121009 若已訂課 將背景變成藍色(l_c checked) by Allison
               dim strChecked : strChecked = ""
               dim strCheckedCss : strCheckedCss = ""
                if ( 0 = sCInt(intSessionStatus) ) then '狀態未訂課
                    strChecked = "" 'css不顯示checked
                    strCheckedCss = "check_off"
                else
                    strChecked = "checked" 'css顯示checked表示有訂課
                    strEvenTag = "" '要把單雙列tag清空 否則css跑不出來
                    strCheckedCss = "check_on"
                end if

                if ( true = bolDisabled ) then
                    strCheckedCss = ""
                end if
            %>
            <div class="session_booking_list_l">
                <div  class="l_c clearfix <%=strChecked%> <%=strEvenTag%> <%=strCannotReserved%>"  <%=strNotReservedDetail%> id="chk_lobby_class_<%=intSessionSn%>_color">
                <!--id="<%=intSessionSn%>" onclick="boxCheck('<%=intSessionSn%>','<%=strSessionTopic%>','<%=strSessionSDate%>','<%=Right("0"&trim(intSessionSTime), 2)%>','<%=intClientAttendSn%>','<%=intSessionSn%>','<%=intNowPage%>','<%=strChecked%>');"-->
                    <h1>
                        <div class="l_r r_01">
                            <!--
                            該時段未訂位會顯示 checkbox
                            if (intSessionStatus = 0) then
                            if (Not isEmptyOrNull(str_from_other_page_con_sn) and CStr(intSessionSn) = CStr(str_from_other_page_con_sn)) then checked="checked" end if 
                            'Modify 20121009 若已訂課，checkbox打勾
                            -->
                             <!--Create 20121012 新增取消訂課警示視窗 cancelCheck function-->
                             <!--Modify 20121029 把checkbox變成圖片(chrome難點擊)--><!--判斷非大會堂課程已訂位 不可click-->
                             <!--<div id="chk_lobby_class_<%=intSessionSn%>_IMG" class="checkbox <%=strCheckedCss%><%=strDisabled%>"  
                             <% if ( false = bolDisabled ) then %>
                             onclick="cancelCheck('chk_lobby_class_<%=intSessionSn%>','<%=strSessionTopic%>','<%=strSessionSDate%>','<%=Right("0"&trim(intSessionSTime), 2)%>','<%=intClientAttendSn%>','<%=intSessionSn%>','<%=intNowPage%>','<%=strChecked%>');"
                             <% end if %>
                             > 
                            </div>-->
                            <% '排除特殊字元 topic %>
							<%
							'****  2012-11-08增加 排除預約測試之前的時間  Johnny
							if ( "N"=re_valArr(0) and  true = isDate(re_valArr(1))  and   (DateDiff("n" , re_valArr(1) , strSessionSDate &" "& intSessionSTime &":30" ) < 30)  ) then							
							%>
								<input type="checkbox" style="WIDTH: 30px; HEIGHT: 30px" id="chk_lobby_class_<%=intSessionSn%>" <%=strChecked%> <%=strDisabled%>  onclick="this.checked=false;alert('亲爱的客户，您的电脑尚未完成测试！\n您预约的电脑测试时间：<%=getFormatDateTime( re_valArr(1) , 1)%>\n为确保您的上课品质，请于电脑测试完成后再进行课程预约。');"/>
							<%	else%>
								<input type="checkbox" style="WIDTH: 30px; HEIGHT: 30px" id="chk_lobby_class_<%=intSessionSn%>" <%=strChecked%> <%=strDisabled%>  onclick="cancelCheck('chk_lobby_class_<%=intSessionSn%>','','<%=strSessionSDate%>','<%=Right("0"&trim(intSessionSTime), 2)%>','<%=intClientAttendSn%>','<%=intSessionSn%>','<%=intNowPage%>','<%=strChecked%>');"/>
							<%
							end if		
							'****  2012-11-08 End							
							%>
									<input type="hidden" id="<%=intSessionSn%>_topic" value="<%=strSessionTopic%>" />
                        </div>
                    </h1>
                    <h2>
                        <div class="l_r r_02" style="padding-left: 35px;">
                            <div class="date" style="color: Black;">
                            <%
                            '開課日期
                            response.Write "<span id='span_lobby_class_date_" & intSessionSn & "'>" & trim(strSessionSDate) & "</span>" 
                            '若為今日開課，顯示今天標籤
                            if  ( getFormatDateTime(strSessionSDate, 5) = getFormatDateTime(date(), 5) ) then
                                response.Write "  <span class='today'>今天</span>"
                                ' <span class="today"></span>
                            end if
                            %>
                            </div>
                            <div class="time" style="color: Black;">
                            <%
                            '開課時間
                            response.Write "<span id='span_lobby_class_time_"&intSessionSn&"'>" & Right("0" & trim(intSessionSTime), 2) & ":30" & "</span><br/>"
                            %>
                            </div>
                            <div class="consultant_name" >
                                <%
                                '顧問名字
                                'Create 20121012 若已經訂位，則顯示已訂位標籤 by Allison
                                'Create 20121012 若已經訂位(一對一/一對三)，則顯示已訂位(一對一/一對三)標籤 by Allison
                                dim strBooked : strBooked =""
                                if ( true = bolShortSession ) then
                                     strBooked = "<div class='booked'>已订位<br/>短课程大会堂</div>"
                                elseif  ( 1 = sCInt(intSessionStatus) ) then 
                                    strBooked = "<div class='booked'>已订位</div>"
                                elseif ( 4 = sCint(intSessionType) OR 14 = sCint(intSessionType) ) then
                                    'strBooked = "<div class='booked'>已订位<br/>一对三</div>"
                                    strBooked = "<div class='booked'>已订位<br/>小班制</div>"
                                elseif ( 7 = sCint(intSessionType) ) then
                                    strBooked = "<div class='booked'>已订位<br/>一对一</div>"
                                else
                                    strBooked = ""
                                end if
                                response.Write "<span>" & trim(strConsultantFName) & " " & trim(strConsultantLName) & "</span>" & strBooked
                                %>
                                <input type="hidden" id="hdn_lobby_class_con_ename_<%=intSessionSn%>" value="<%=trim(strConsultantEName)%>" />
                            </div>
                            <%
                            dim objFSO : objFSO = null
                            dim strFilePath : strFilePath = ""
                            '顧問的大頭照 <!--20121003 修改 教材的圖片 -->
                            if ( Not isEmptyOrNull(intMaterialSn) ) then
                                strImageUrl = "/images/mtl_thumb/" & intMaterialSn & ".jpg"
                                Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
                                if ( true = objFSO.fileExists(Server.MapPath(strImageUrl)) ) then  '有圖片，就show教材圖片
                                    strImageUrl = strImageUrl
                                else '找不到圖片，show預設圖
                                    'strImageUrl = "/images/no_face_img.jpg"
                                    strImageUrl = "/images/mtl_thumb/mov.jpg"
                                end if
                                Set objFSO = nothing
                            else '找不到圖，show預設圖
                                '若找不到顧問的大頭照 則show出預設的圖片
                                'strImageUrl = "/images/no_face_img.jpg"
                                strImageUrl = "/images/mtl_thumb/mov.jpg"
                            end if
                            %>
                            <img src="<%=strImageUrl%>" width="70" /><!--09/11/19新增放圖：此圖比例是依mysession的顧問圖等比例縮小 width=70 height=81-->
                        </div>
                    </h2>
                    <h3>
                        <div class="l_r r_03" style="padding-left: 100px;">
                            <p>
                                <div class="title_c">
                                    <%
                                    '大會堂的標題
                                    response.Write "<span id='span_lobby_class_topic_"&intSessionSn&"' style='display:none;'>" & objDencode.HtmlEncode(Replace(trim(strSessionTopic), "<br>", chr(10))) & "</span>"
                                    response.Write trim(strSessionTopic)
                                    %>
                            </p>
                        </div>
                        <div class="desc">
                            <%
                            '大會堂內容簡介
                            response.Write objDencode.HtmlEncode(replace(trim(strSessionIntro), "<br>", chr(10)))
                            %>
                        </div>
                </div>
                </h3>
                <h4>
                    <div class="l_r r_04" style="padding-left: 155px;">
                        <%
                        '大會堂適合的等級
                        if ( Not isEmptyOrNull(strSessionApplyLv) ) then
                            strSessionApplyLv=trim(strSessionApplyLv) '去空白
				            if ( Left(strSessionApplyLv, 1) = "," ) then
                                strSessionApplyLv = Right(strSessionApplyLv, Len(strSessionApplyLv)-1)
                                'response.Write " strSessionApplyLv : " & strSessionApplyLv & "<br/>"
                            end if
				            if ( Right(strSessionApplyLv, 1) = "," ) then 
                                strSessionApplyLv = Left(strSessionApplyLv, Len(strSessionApplyLv)-1)
                                'response.Write " strSessionApplyLv : " & strSessionApplyLv & "<br/>"
                            end if
				            arrSessionApplyLv = split(strSessionApplyLv, ",")
                            'response.Write " arrSessionApplyLv(0) : " & arrSessionApplyLv(0) & "<br/>"
				            response.Write "LV" & arrSessionApplyLv(0) & " -" & arrSessionApplyLv(UBound(arrSessionApplyLv)) & "<br/>" '适用程度LV
                            'response.Write " UBound(arrSessionApplyLv) : " & UBound(arrSessionApplyLv) & "<br/>"
                        end if                       
                        %>
                    </div>
                </h4>
                <h5>
                    <%
                    '已額滿 訊息
			        'if ( sCInt(intSessionNowNum) >= sCInt(intSessionMaxNum) ) then                
                    '    response.Write str_msg_font_limit_full & "<br/>"    '已額滿
                    'end if

                    '已預定 訊息 (已經預定的大會堂 課程的lev > 13)
                    '此時段已經預定了
                    if ( 1 = sCInt(intSessionStatus) AND sCInt(intSessionMemberLv) = sCInt(intSessionSn) ) then
                        '大會堂
                        if ( 13 < sCInt(intSessionMemberLv) ) then
                            'response.Write "<span id='cancel_class_info_" & intClientAttendSn &"'>" & str_msg_font_already_order_class & "</span><br/>"   '您已预订

                            '四小時內無法取消 不顯示四小時內的課程 或合約已到期
                            strCancelClassDatetime = strSessionSDate & " " & Right("0"&intSessionSTime, 2) & ":30"
                            if ( DateDiff("n", now(), strCancelClassDatetime) >= 240 AND bol_client_have_valid_contract ) then
                                'response.Write "<a id='cancel_class_link_" & intClientAttendSn & "' href='javascript:void(0);' onclick=""javascript:g_obj_reservation_class_ctrl.cancelClass('"&intClientAttendSn&"', '"&trim(strSessionSDate)&"', '"&trim(intSessionSTime)&"');"">" & str_msg_a_cancel_class & "</a>"
                                'response.Write "<img id='img_cancel_class_loading_" & intClientAttendSn & "' src='/lib/javascript/JQuery/images/ajax-loader.gif'  width='25' height='25' border='0' style='display:none;' />"
                            end if
                            'response.Write "</span>" '取消
                        '一對一或一對三
                        else
                            strAlreadyOrderTitle = ""
                            if ( Not isEmptyOrNull(intClientAttendType) ) then
                                '一對三
                                if ( sCInt(intClientAttendType) = CONST_CLASS_TYPE_ONE_ON_THREE OR sCInt(intClientAttendType) = CONST_CLASS_TYPE_ONE_ON_FOUR ) then
                                    'strAlreadyOrderTitle = getWord("NORMAL_CLASS_NOTE") '一般课(1-3人)
                                    strAlreadyOrderTitle = "小班制"
                                '一對一
                                elseif ( sCInt(intClientAttendType) = CONST_CLASS_TYPE_ONE_ON_ONE ) then
                                    strAlreadyOrderTitle = getWord("VIP_CLASS_NOTE") '一般课(1對1)
                                elseif ( 98 = sCInt(intClientAttendType) ) then
                                    strAlreadyOrderTitle = "10mins大会堂"
                                end if
                            end if
                            response.Write "<span id='cancel_class_info_"&intClientAttendSn&"' title='"&strAlreadyOrderTitle&"'>" & getWord("ALREADY_RESERVATION_CLASS") & "</span><br/>"   '此时段您已有预订课程
                        end if
                    end if
                    %>
                </h5>
            </div>
</div>
</li>
<!--課程列表end-->
<br class="clear" />
<%
            'Create 20121009 若為雙數行，變色 。0~7故雙數為1,3,5,7 --End --by Allison
            if ( intI mod 2 = 1 ) then  
                response.Write "</div>"
            end if
            intI = intI + 1
        end if
    end if
    objLobbySessionData.movenext
wend
%>
</ul>
<input type="hidden" id="page" name="page" value="<%=intNowPage%>" />
<div class="session_booking_body">
    <!--Create 頁數-->
    <div class="page"> 
        <% if ( sCInt(intNowPage) <> 1 ) then %><!--若不是第一頁，顯示可點選之 "<" 和 "最前頁"  -->
            <span><%="<a href='javascript:;' style='cursor:hand;text-decoration:none;' id='page_" & 1 & "' onclick=""ChangePage('" & 1 & "', '1');"">" %>最前页</a></span>
            <span><%="<a href='javascript:;' style='cursor:hand;text-decoration:none;' id='page_" & intNowPage - 1 & "' onclick=""ChangePage('" & intNowPage - 1 & "', '1');"">" %>&lt;</a></span>
        <% else %> <!--若是第一頁，只顯示字，無超連結  -->
            <span>最前页</span> <span>&nbsp;&lt;</span>
        <% end if %>
        <% '顯示頁數
	        Dim intJ : intJ = 0
	        for intJ = 1 to sCInt(intTotalPage)
		        if ( sCInt(intJ) <> sCInt(intNowPage) ) then
			        response.Write "<a href='javascript:;' style='cursor:hand;text-decoration:none;' id='page_" & intJ & " ' onclick=""ChangePage('" & intJ & "', '1');"">"
		        else
			        response.Write "&nbsp;<font color='red'><strong>"
		        end if
		        response.Write intJ
		        if ( sCInt(intJ) <> sCInt(intNowPage) ) then
			        response.Write "</a>"
		        else
			        response.Write "</strong></font>&nbsp;"
		        end if
	        next
        %>
        <% if ( sCInt(intNowPage) <> sCInt(intTotalPage) ) then %><!--若不是最後頁，顯示 ">" 和 "最後頁" -->
            <span><%="<a href='javascript:;' style='cursor:hand;text-decoration:none;' id='page_" & intNowPage + 1 & "' onclick=""ChangePage('" & intNowPage + 1 & "', '1');"">" %>&gt;</a></span>
            <span><%="<a href='javascript:;' style='cursor:hand;text-decoration:none;' id='page_" & intTotalPage & "' onclick=""ChangePage('" & intTotalPage & "', '1');"">" %>最后页</a></span>
        <% else %>
            <span>&gt;</span> <span>最后页</span>
        <% end if %>
    </div><!--end .page-->
    <div class="btn_submit">
        <input type="button" id="submit_button" value="<%=str_msg_btn_confirm_order%>" class="btn_1"  onclick="submitReservate();"/><!--確認訂課-->
    </div>
</div><!--end .session_booking_body-->
<%
else
    'TODO: 沒資料時的預設值
    'response.write"<script language=javascript>alert('无相关大会堂课程，请重新查询。');</script>"
    response.Write str_msg_font_search_no_data  '沒有符合搜尋條件的資料
end if
Set objDencode = Nothing
%>
</body>