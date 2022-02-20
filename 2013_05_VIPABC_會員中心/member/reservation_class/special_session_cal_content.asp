<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="include/include_prepare_and_check_data.inc"-->
<!--#include virtual="/lib/include/class/DataOutputEncode.asp" -->
<!--#include virtual="/program/class/lobby/CourseUtility.asp" -->
<% 
dim bolDebugMode : bolDebugMode = false '正式改為false

'日曆用到的參數Start
dim intMinTime : intMinTime = getRequest("Min_Time", CONST_DECODE_ESCAPE)
dim intMaxTime : intMaxTime = getRequest("Max_Time", CONST_DECODE_ESCAPE)
dim strMinDate : strMinDate = request("Min_Date")
dim strMaxDate : strMaxDate =  request("Max_Date")
dim intJ : intJ = 0
dim intI : intI = 0
dim intK : intK = 0 '課堂時間迴圈
dim intL : intL = 0 '日期迴圈
dim strShowHour : strShowHour = "" '開課小時
dim strShowMinute : strShowMinute = ":30" '開課分鐘
dim intMaxNum : intMaxNum = DATEDIFF("d", strMinDate, strMaxDate) '間隔幾天
dim arrWeekDay : arrWeekDay = Array("","星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六") ' 記錄星期幾的陣列
dim intDateCount : intDateCount = 2 '變換div class變數
dim strTodayTag : strTodayTag = "" '今日tag
dim strEven : strEven = "" 
dim intNowPage : intNowPage = getRequest("intNowPage", CONST_DECODE_ESCAPE)
dim strCheckDivOpen : strCheckDivOpen = getRequest("strCheckDivOpen", CONST_DECODE_ESCAPE)
dim bolHasData : bolHasData = false '查詢預設變數

'dim strSession : strSession =  getSession(strCheckDivOpen, CONST_DECODE_ESCAPE)
Session(strCheckDivOpen) = "true"

if ( true  = bolDebugMode ) then
	response.Write "間隔 : " &intMaxNum&"<br/>"
    response.Write "星期幾 : " &WeekDay(strMinDate)&":" &arrWeekDay(WeekDay(strMinDate))&"<br/>"
    response.write "只顯示月日 : " & right(strMinDate, len(strMinDate)-5)&"<br/>" '只顯示月/日
    response.write "test : " & "0" & intDateCount &"<br/>"
    response.write "intMinTime : " & "0" & intMinTime &"<br/>"
    response.write "strMinDate : " & strMinDate &"<br/>"
    response.write "date : " & date &"<br/>"
end if
'日曆用到的參數End

'DB query Start
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
Dim intClientSn, str_arr_tmp, strCancelClassDatetime, strAlreadyOrderTitle
Dim objDencode : set objDencode = new DataOutputEncode

'20120316 阿捨新增 無限卡產品判斷	
Dim strUnlimited : strUnlimited = getSession("unlimited", CONST_DECODE_NO)
if ( isEmptyOrNull(strUnlimited) ) then
    Dim obj_member_opt_exe : Set obj_member_opt_exe = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
    Dim int_member_result_exe : int_member_result_exe = obj_member_opt_exe.prepareData(g_var_client_sn, CONST_VIPABC_RW_CONN)
    if ( int_member_result_exe = CONST_FUNC_EXE_SUCCESS ) then        
        strNowUseProductSn = obj_member_opt_exe.getData(CONST_CONTRACT_PRODUCT_SN) '離目前時間最近的一筆 客戶合約的產品序號
    else
'        '錯誤訊息
'        'response.Write "member prepareData error = " & obj_member_opt_exe.PROCESS_STATE_RESULT & "<br>"
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
 'g_str_lobby_condition_datetime = " AND ((CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE()+1, 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_
'                                    '"      OR (CONVERT(varchar, lobby_session.session_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
 
 '設定 lobby_session 時間條件 (大會堂允許開課前5分鐘都可以訂)
 str_lobby_condition_datetime = " AND ((CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE(), 111) AND lobby_session.session_time = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
                                 " OR (CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE(), 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_ 
                                "      OR (CONVERT(varchar, lobby_session.session_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
 
 ''設定 client_attend_list 時間條件 (大會堂只允許訂24時候的課)
 'g_str_attend_condition_datetime = " AND ((CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE()+1, 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
'                                    '"      OR (CONVERT(varchar, client_attend_list.attend_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
 
 '設定 client_attend_list 時間條件 (大會堂允許開課前5分鐘都可以訂)
 str_attend_condition_datetime = " AND ((CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE(), 111) AND client_attend_list.attend_sestime = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
                                 " OR (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE(), 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
                               "      OR (CONVERT(varchar, client_attend_list.attend_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
 
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
 'if (Not isEmptyOrNull(str_search_time)) then
     'str_lobby_condition_time = " AND (lobby_session.session_time = '"&str_search_time&"')  "
     'str_attend_condition_time = " AND (client_attend_list.attend_sestime = '"&str_search_time&"')  "
 'end if

  '大會堂開課時段 條件設定
 if ( Not isEmptyOrNull(intMinTime) AND Not isEmptyOrNull(intMaxTime)  ) then
     str_lobby_condition_time = " AND (lobby_session.session_time BETWEEN " & intMinTime &"AND " & intMaxTime & ")  "
 end if

  if ( Not isEmptyOrNull(strMinDate) AND Not isEmptyOrNull(strMaxDate)  ) then
    strMinDate = getFormatDateTime(strMinDate, 5)
    strMaxDate = getFormatDateTime(strMaxDate, 5)
    str_lobby_condition_time = str_lobby_condition_time & " AND ( CONVERT(varchar, lobby_session.session_date, 111)  BETWEEN CONVERT(varchar, '"&strMinDate&"', 111)  AND CONVERT(varchar, '"& strMaxDate &"', 111)  )  "
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
 'g_str_lobby_condition_datetime = " AND ((CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE()+1, 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_
'								    '"      OR (CONVERT(varchar, lobby_session.session_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
 
 '設定 lobby_session 時間條件 (大會堂允許開課前5分鐘都可以訂)
 str_lobby_condition_datetime = " AND ((CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE(), 111) AND lobby_session.session_time = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
							    " OR (CONVERT(varchar, lobby_session.session_date, 111) = CONVERT(varchar, GETDATE(), 111) AND lobby_session.session_time > '"&Hour(now)&"' ) " &_ 
							    "      OR (CONVERT(varchar, lobby_session.session_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
 
 ''設定 client_attend_list 時間條件 (大會堂只允許訂24時候的課)
 'g_str_attend_condition_datetime = " AND ((CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE()+1, 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
'								    '"      OR (CONVERT(varchar, client_attend_list.attend_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
 
 '設定 client_attend_list 時間條件 (大會堂允許開課前5分鐘都可以訂)
 str_attend_condition_datetime = " AND ((CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE(), 111) AND client_attend_list.attend_sestime = '"&Hour(now)&"' AND " & Minute(now) & " < 25 ) " &_
							    " OR (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, GETDATE(), 111) AND client_attend_list.attend_sestime > '"&Hour(now)&"' ) " &_
							    "      OR (CONVERT(varchar, client_attend_list.attend_date, 111) > CONVERT(varchar, GETDATE(), 111))) "
  
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
'strSql = strSql & " ISNULL(lobby_order_count.snum, 0) AS snum, "  ' TODO : 已達上限人數判斷，submit之後再判斷
'--大會堂適合的等級 11
strSql = strSql & " lobby_session.lev as apply_lev, "
'顧問的大頭照 hex code 12
'strSql = strSql & " lobby_session.con_img, "  '使用顾问头像 paul
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
strSql = strSql & " ROW_NUMBER() OVER (ORDER BY lobby_session.sdate, lobby_session.stime, client_order_lobby.status DESC) AS row "
strSql = strSql & " FROM ( "
'--有開課的大會堂及顧問的資料(>=今天) <!-- 新增教材-->
strSql = strSql & " SELECT lobby_session.sn, lobby_session.material, CONVERT(varchar, lobby_session.session_date, 111) AS sdate, lobby_session.session_time AS stime,  "
strSql = strSql & " isnull(con_basic.basic_fname, '') + ' ' + isnull(con_basic.basic_lname,'') AS ename, " & str_lobby_session_topic & ", " & str_lobby_session_introduction & ", "
strSql = strSql & " CONVERT(VARCHAR, lobby_session.session_date, 112)+ RIGHT('0' + CAST(lobby_session.session_time AS VARCHAR(2)),2)+ CAST(lobby_session.sn AS VARCHAR(5)) AS dcode,	lobby_session.limit, lobby_session.lev, " ''''
'strSql = strSql & " CASE WHEN con_basic.Img IS NULL THEN 0 ELSE 1 END  AS con_img, " '使用顾问头像 paul 
'strSql = strSql & " ISNULL(con_basic.img,'') AS con_img, " '使用顾问头像 paul 
strSql = strSql & " con_basic.con_sn, isnull(con_basic.basic_fname, '') AS basic_fname, isnull(con_basic.basic_lname,'') AS basic_lname, "
'--新增dt
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
strSql = strSql & " ) AS lobby_session  "
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
strSql = strSql & " SELECT CONVERT(varchar, attend_date, 112) + RIGHT('0'+CAST(attend_sestime AS varchar(2)),2) AS dt, ISNULL(attend_level,0) AS attend_level, special_sn, client_attend_list.sn AS client_attend_list_sn, client_attend_list.attend_livesession_types, "
'strSql = strSql & " SELECT CONVERT(varchar, attend_date, 112) + RIGHT('0'+CAST(attend_sestime AS varchar(2)),2) AS dt, special_sn, client_attend_list.sn AS client_attend_list_sn, client_attend_list.attend_livesession_types "
'--新增status
strSql = strSql & " CASE ISNULL(CONVERT(varchar, attend_date, 112) + RIGHT('0'+CAST(attend_sestime AS varchar(2)),2), 0) WHEN 0 THEN 0 ELSE 1 END AS status "
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
strSql = strSql & " WHERE (client_sn = @clientsn ) AND (valid = 1) AND ( attend_level <= 12 )" & str_attend_condition_datetime & str_attend_condition_con & str_attend_condition_time
strSql = strSql & " ) AS client_order_session "
strSql = strSql & " ON client_order_session.date_time = lobby_session.date_time "

'LEFT(lobby_session.dcode, 10) = client_order_lobby.dt AND client_order_lobby.attend_level = lobby_session.sn "
'strSql = strSql & " ORDER BY lobby_session.sdate, lobby_session.stime 
strSql = strSql & " ) a "
'strSql = strSql & " WHERE row BETWEEN @top AND @bottom "
strSql = strSql & " ORDER BY a.sdate, a.stime "
arrParam = Array(intClientSn, intClientSn)
'arrParam = Array(intClientSn, intTopNum, intBottomNum)
'arr_lobby_data = excuteSqlStatementRead(strSql, arrParam, CONST_VIPABC_RW_CONN)
Set objLobbySessionData = excuteSqlStatementReadEach(strSql, arrParam, CONST_VIPABC_R_CONN)

 if ( true = bolDebugMode ) then
    response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
    response.write "查詢條件str_lobby_condition_lev : " & str_lobby_condition_lev & "<br/>"
    response.write "查詢條件str_lobby_condition_topic : " & str_lobby_condition_topic & "<br/>"
    response.write "查詢條件str_lobby_condition_con : " & str_lobby_condition_con & "<br/>"
    response.Write "str_search_lev : " & str_search_lev & "<br/>"
	response.Write "str_search_topic : " & str_search_topic & "<br/>"
	response.Write "str_search_con : " & str_search_con & "<br/>"
 end if

'字典檔
Set objSessionTopic = CreateObject("Scripting.Dictionary")
Set objSessionApplyLv = CreateObject("Scripting.Dictionary")
Set objSessionSn = CreateObject("Scripting.Dictionary")
Set objSessionStatus = CreateObject("Scripting.Dictionary")
Set objSessionSTime = CreateObject("Scripting.Dictionary")
Set objSessionSDate = CreateObject("Scripting.Dictionary")
Set objConsultantFName = CreateObject("Scripting.Dictionary")
Set objConsultantLName = CreateObject("Scripting.Dictionary")
Set objConsultantEName = CreateObject("Scripting.Dictionary")
Set objintConSn = CreateObject("Scripting.Dictionary")
'Set objbyteImg = CreateObject("Scripting.Dictionary")
Set objMaterialSn = CreateObject("Scripting.Dictionary")
Set objSessionIntro = CreateObject("Scripting.Dictionary")
Set objSessionMaxNum = CreateObject("Scripting.Dictionary")
Set objSessionNowNum = CreateObject("Scripting.Dictionary")
Set objSessionMemberLv = CreateObject("Scripting.Dictionary")
Set objClientAttendSn = CreateObject("Scripting.Dictionary")
Set objClientAttendType = CreateObject("Scripting.Dictionary")
Set objSessionType = CreateObject("Scripting.Dictionary") '判斷一對一/一對多

'字典檔設陣列 因key值為date+time 若同時段有多筆課程 用陣列存 
dim arrSessionTopic : arrSessionTopic = null
dim arrSessionApplyLv : arrSessionApplyLv =  null
dim arrSessionSn : arrSessionSn = null
dim arrSessionStatus : arrSessionStatus = null
dim arrConsultantFName : arrConsultantFName = null
dim arrConsultantLName : arrConsultantLName = null
dim arrConsultantEName : arrConsultantEName = null
dim arrConSn : arrConSn= null
dim arrMaterialSn : arrMaterialSn = null
'dim arrByteImg : arrByteImg = null
dim arrSessionIntro : arrSessionIntro = null
dim arrSessionMaxNum : arrSessionMaxNum = null
dim arrSessionNowNum : arrSessionNowNum = null
dim arrSessionMemberLv : arrSessionMemberLv = null
dim arrClientAttendSn :   arrClientAttendSn = null
dim arrClientAttendType : arrClientAttendType = null
dim arrSessionSTime : arrSessionSTime = null
dim arrSessionSDate : arrSessionSDate = null
dim arrSessionType : arrSessionType = null '判斷一對一/一對多

dim strSessionWord : strSessionWord = ""
dim intNewSessionType : intNewSessionType = ""
dim bolSession : bolSession = false
'if ( not objLobbySessionData.eof ) then
'else
'    response.Write str_msg_font_search_no_data  '沒有符合搜尋條件的資料
'end if

'字典檔 給值
 if ( not objLobbySessionData.eof ) then
    while not objLobbySessionData.eof
	    intConSn = objLobbySessionData("con_sn")
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
        intSessionMaxNum = objLobbySessionData("limit") '訂課上限
        'intSessionNowNum = objLobbySessionData("snum") '訂課人數
        intSessionMemberLv = objLobbySessionData("lev") '若有人訂課，則寫session_Sn值，否則0
        intClientAttendSn = objLobbySessionData("client_attend_list_sn") 
        intClientAttendType = objLobbySessionData("attend_livesession_types")
        intSessionType = objLobbySessionData("session_type") '判斷一對一/一對多

        if ( true = bolDebugMode) then
            response.write "intSessionSTime : " & intSessionSTime &"<br/>"
            response.write "strSessionSDate : " & strSessionSDate &"<br/>"
        end if
        strSessionSDate = getFormatDateTime( strSessionSDate ,5)
        intSessionSTime = right("0" & intSessionSTime, 2) 
        strSessionSTime = intSessionSTime & ":30" '時間格式 00:30
        strKeyWord = strSessionSDate & " " & strSessionSTime  '字典檔的keyword
        if ( true = bolDebugMode) then
            response.write "strKeyWord _strKeyWord : " & strKeyWord &"<br/>"
        end if
        '上一個keyword值和目前這筆相同 則寫入陣列
        if ( strKeyWord = strTempKeyword ) then
            intTempNum = intTempNum + 1
            arrSessionTopic(intTempNum) = strSessionTopic
            arrSessionApplyLv(intTempNum) = strSessionApplyLv
            arrSessionSn(intTempNum) = intSessionSn
            arrSessionStatus(intTempNum) = intSessionStatus
            arrSessionSTime(intTempNum) = intSessionSTime
            arrSessionSDate(intTempNum) = strSessionSDate
            arrSessionType(intTempNum) = intSessionType '判斷一對一/一對多
            arrConsultantFName(intTempNum) = strConsultantFName
            arrConsultantLName(intTempNum) = strConsultantLName
		    arrConSn(intTempNum) = intConSn
            arrConsultantEName(intTempNum) = strConsultantEName 'fname+lname
            arrMaterialSn(intTempNum) = intMaterialSn
            arrSessionIntro(intTempNum) = strSessionIntro
            arrSessionMaxNum(intTempNum) = intSessionMaxNum
            'arrSessionNowNum(intTempNum) = intSessionNowNum
            arrSessionMemberLv(intTempNum) = intSessionMemberLv '若有人訂課，則寫session_Sn值，否則0
            arrClientAttendSn(intTempNum) = intClientAttendSn
            arrClientAttendType(intTempNum) = intClientAttendType
        else 'keyword只有一筆 存到陣列的第0個位置 要先把陣列清空再存
            redim arrSessionTopic(20)
            redim arrSessionApplyLv(20)
            redim arrSessionSn(20)
            redim arrSessionStatus(20)
            redim arrConsultantFName(20)
            redim arrConsultantLName(20)
            redim arrConsultantEName(20)
			redim arrConSn(20)
            redim arrMaterialSn(20)
            redim arrSessionIntro(20)
            redim arrSessionMaxNum(20)
            'redim arrSessionNowNum(20)
            redim arrSessionMemberLv(20)
            redim arrClientAttendSn(20)
            redim arrClientAttendType(20)
            redim arrSessionSTime(20)
            redim arrSessionSDate(20)
            redim arrSessionType(20)  '判斷一對一/一對多

            arrSessionTopic(0) = strSessionTopic
            arrSessionApplyLv(0) = strSessionApplyLv
            arrSessionSn(0) = intSessionSn
            arrSessionStatus(0) = intSessionStatus
            arrConsultantFName(0) = strConsultantFName
            arrConsultantLName(0) = strConsultantLName
            arrConsultantEName(0) = strConsultantEName 'fname+lname
			arrConSn(0) = intConSn
            arrMaterialSn(0) = intMaterialSn
            arrSessionIntro(0) = strSessionIntro
            arrSessionMaxNum(0) = intSessionMaxNum
            'arrSessionNowNum(0) = intSessionNowNum
            arrSessionMemberLv(0) = intSessionMemberLv
            arrClientAttendSn(0) = intClientAttendSn
            arrClientAttendType(0) = intClientAttendType
            arrSessionSDate(0) = strSessionSDate
            arrSessionSTime(0) = intSessionSTime
            arrSessionType(0) = intSessionType '判斷一對一/一對多

            intTempNum = 0
        end if
        '存好的字典檔陣列 塞值給相對應的位置
        objSessionApplyLv(strKeyWord) = arrSessionApplyLv
        objSessionTopic(strKeyWord) = arrSessionTopic
        objSessionSn(strKeyWord) = arrSessionSn
        objSessionStatus(strKeyWord) = arrSessionStatus
        objSessionSTime(strKeyWord) = arrSessionSTime
        objSessionSDate(strKeyWord) =arrSessionSDate
        objConsultantFName(strKeyWord) = arrConsultantFName
        objConsultantLName(strKeyWord) = arrConsultantLName
		objintConSn(strKeyWord) = arrConSn
        objConsultantEName(strKeyWord) = arrConsultantEName
        objMaterialSn(strKeyWord) = arrMaterialSn
        objSessionIntro(strKeyWord) = arrSessionIntro
        objSessionMaxNum(strKeyWord) = arrSessionMaxNum
        'objSessionNowNum(strKeyWord) = arrSessionNowNum
        objSessionMemberLv(strKeyWord) = arrSessionMemberLv
        objClientAttendSn(strKeyWord) = arrClientAttendSn
        objClientAttendType(strKeyWord) = arrClientAttendType
        objSessionType(strKeyWord) = arrSessionType '判斷一對一/一對多

        strTempKeyword = strKeyWord
        objLobbySessionData.movenext
    wend
 end if
 'DB query End

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
<!--課程內容Start-->
<%
for intI = sCInt(intMinTime) to sCInt(intMaxTime) '時間迴圈
    if ( intI = sCInt(intMinTime) ) then '第一列 串日期標頭的css
        strShowFirstString = "<div class='l_c c_h clearfix'>"
    else '非第一列 內容 單雙數列css
        if ( 0 = (intI mod 2) ) then 
            strShowOtherString = "<div class='l_c c_l clearfix even'> "
        else
            strShowOtherString = "<div class='l_c c_l clearfix'> "
        end if
        strShowTime = right("0" & intI, 2) & ":30" '串第一欄，時間
        strShowOtherString = strShowOtherString & "<div class='l_r r_01'>" 
        strShowOtherString = strShowOtherString & strShowTime
        strShowOtherString = strShowOtherString & "</div>" 'end for "<div class='l_r r_01'>" 
    end if

    for intJ = -1 to (intMaxNum-1) '日期迴圈 日期區間天數 intMaxNum
        if ( intI = sCInt(intMinTime) ) then '第一列的日期呈現
            intK = intJ + 2
            if ( -1 = intJ ) then '第一列的第一欄 若從0開始 日期會往左偏移 所以從-1開始
                if ( intNowPage <= 0 ) then '第一頁 不show前三天tag
                    strPreTag="<div class='btn_pre'>&nbsp;</div>"
                else
                    strPreTag = "<div class='btn_pre' style='cursor:hand;text-decoration:none;' onclick='ChangePage(" & intNowPage -1 & ",2)'><a href='javascript:;'>前三日</a></div>"
                end if
                strShowPreTag = "<div class='l_r r_01'>" & strPreTag &"</div>" '前三天tag
                strShowFirstString = strShowFirstString & strShowPreTag
            else
                strShowFirstString = strShowFirstString & "<div class='l_r r_0" & intK & "'>"
                strShowDate = DateAdd("d", intJ, strMinDate) '日期增加
                strShowDate = getFormatDateTime( strShowDate ,5)
                '判斷後三天tag
                if ( (intMaxNum-1) = intJ ) then
                    if  ( intNowPage >= 2 ) then '第三頁 不show後三天tag
                        strNextTag = "<div class='btn_next'>&nbsp;</div>"
                    else
                        strNextTag = "<div class='btn_next' style='cursor:hand;text-decoration:none;' onclick='ChangePage("& intNowPage + 1 &", 2)' ><a href='javascript:;'>後三天</a></div>" 
                    end if
                    strShowFirstString = strShowFirstString & strNextTag
                end if
                
                '判斷今日tag
                if ( strShowDate = getFormatDateTime(date(), 5) ) then
                    strTodayTag = "<span class=""today"">今天</span>"
                else 
                    strTodayTag = ""
                end if
                strShowFirstString = strShowFirstString & getFormatDateTime(strShowDate, 12) & " " & arrWeekDay(WeekDay(strShowDate)) 
                strShowFirstString = strShowFirstString & strTodayTag
                strShowFirstString = strShowFirstString & "</div>" ' end for "<div class='l_r r_0" & intK & "'>"
            end if
			
        else '非第一列的內文呈現
            if ( -1 = intJ ) then
            else
                strShowDate = DateAdd("d", intJ, strMinDate) '日期增加
                strShowDate = getFormatDateTime( strShowDate ,5)
                intK = intJ + 2
                strShowOtherString = strShowOtherString & "<div class='l_r r_0" & intK & "'>"
                'strShowInfo = "<div class='r_session'><div class='title_c'>"
                
                '取得字典檔的keyword
                strKeyWord = strShowDate & " " & strShowTime
                if ( true = bolDebugMode) then
                    response.write "strShowTime : " & strShowTime &"<br/>"
                    response.write "intSessionSTime : " & intSessionSTime &"<br/>"
                    response.write "strSessionSDate : " & strSessionSDate &"<br/>"
                    response.write "strKeyWord : " & strKeyWord &"<br/>"
                end if
                '主題多筆陣列
                arrSessionTopic = objSessionTopic(strKeyWord)
                '程度多筆陣列
                arrSessionApplyLv = objSessionApplyLv(strKeyWord)
                '課程多筆陣列
                arrSessionSn = objSessionSn(strKeyWord)
                '訂課狀態多筆陣列
                arrSessionStatus = objSessionStatus(strKeyWord)
                '開課時間多筆陣列
                arrSessionSTime = objSessionSTime(strKeyWord)
                '開課日期多筆陣列
                arrSessionSDate = objSessionSDate(strKeyWord)
                '顧問多筆陣列FirstName
                arrConsultantFName = objConsultantFName(strKeyWord)
                '顧問多筆陣列LastName
                arrConsultantLName = objConsultantLName(strKeyWord)
                '顧問多筆陣列AllName
                arrConsultantEName = objConsultantEName(strKeyWord)
				'顾问con_sn
				arrConSn = objintConSn(strKeyWord)
                '教材多筆陣列
                arrMaterialSn = objMaterialSn(strKeyWord)
				'介紹多筆陣列
                arrSessionIntro = objSessionIntro(strKeyWord)
                '上限人數多筆陣列
                arrSessionMaxNum = objSessionMaxNum(strKeyWord)
                '目前人數多筆陣列
                'arrSessionNowNum = objSessionNowNum(strKeyWord)
                '
                arrSessionMemberLv = objSessionMemberLv(strKeyWord)
                '客戶Sn多筆陣列
                arrClientAttendSn = objClientAttendSn(strKeyWord)
                '
                arrClientAttendType = objClientAttendType(strKeyWord)
                '判斷一對一/一對多
                arrSessionType = objSessionType(strKeyWord)

                if ( true = isArray(arrSessionTopic) ) then '當同時段有多筆資料時
                    intTotal = Ubound(arrSessionTopic)
                    for intL = 0 to intTotal
                        strSessionTopic = arrSessionTopic(intL)
                        if ( Not isEmptyOrNull(strSessionTopic) ) then

                            strSessionTopic = replace(strSessionTopic, """","")

                            strSessionIntro = arrSessionIntro(intL) 'title資訊
                            strConsultantFName = arrConsultantFName(intL)
                            strConsultantLName = arrConsultantLName(intL)
                            strConsultantEName  = arrConsultantEName(intL) 'title資訊
							intConSn = arrConSn(intL)
                            strSessionApplyLv = arrSessionApplyLv(intL) '程度，下面有只留頭尾
                            intMaterialSn = arrMaterialSn(intL) '教材Sn
                            intSessionStatus = arrSessionStatus(intL) '判斷是否訂課，訂課:1；未訂課:0
                            intSessionType = arrSessionType(intL) '判斷一對一/一對多
                            
                            '程度 重新命名並在此format
                            strTempSessionApplyLv = strSessionApplyLv 
                            strTempSessionApplyLv=trim(strTempSessionApplyLv)
				            if ( Left(strTempSessionApplyLv, 1) = "," ) then
                                strTempSessionApplyLv = Right(strTempSessionApplyLv, Len(strTempSessionApplyLv)-1)
                            end if
				            if ( Right(strTempSessionApplyLv, 1) = "," ) then 
                                strTempSessionApplyLv = Left(strTempSessionApplyLv, Len(strTempSessionApplyLv)-1)
                            end if
                            arrTempSessionApplyLv = split(strTempSessionApplyLv, ",")
                            
                            'dim objFSO : objFSO = null
                            'dim strFilePath : strFilePath = ""
                            '顧問的大頭照 <!--20121003 修改 教材的圖片 -->
                            '有教材sn 再判斷是否有圖 沒有皆放預設圖
                            'if ( isEmptyOrNull(byteimg) ) then
								'此部分为课程图片url设置begin
                                'strImageUrl = "/images/mtl_thumb/" & intMaterialSn & ".jpg"
                                'Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
                                'if ( true = objFSO.fileExists(Server.MapPath(strImageUrl)) ) then 
                                    'strImageUrl = strImageUrl
                                'else
                                    'strImageUrl = "/images/no_face_img.jpg"
                                    'strImageUrl = "/images/mtl_thumb/mov.jpg"
                                'end if
                                'Set objFSO = nothing
                            'else
                                'strImageUrl = "/images/no_face_img.jpg" '顾问头像 
                                'strImageUrl = "/images/mtl_thumb/mov.jpg"
                            'end if

                            '滑鼠移至課程上 顯示的課程詳細資訊
                            strTempSessionLv = "LV" & arrTempSessionApplyLv(0) & "-" & arrTempSessionApplyLv(UBound(arrTempSessionApplyLv)) '程度字串
                            strSessionTopic = objDencode.HtmlEncode(Replace(trim(strSessionTopic), "<br>", chr(10)))
                            strShowMouseOver = strSessionTopic & "&#13; 适用程度：" & strTempSessionLv '主題 + 適用程度
                            strSessionIntro = objDencode.HtmlEncode(replace(trim(strSessionIntro), "<br>", chr(10)))
                            strShowMouseOver = strShowMouseOver & "&#13; 顾问：" & strConsultantEName & "&#13; &#13; " & strSessionIntro'顧問+介紹

                            bolHasData = true '查詢條件 若有符合的課程 true

                            intSessionSn = arrSessionSn(intL)
                            intClientAttendSn = arrClientAttendSn(intL)
                            intSessionSTime = arrSessionSTime(intL) 'detail資訊
                            strSessionSDate = arrSessionSDate(intL) 'detail資訊
                            if ( true = bolDebugMode) then
                                response.write "intSessionSTime : " & intSessionSTime &"<br/>"
                                response.write "strSessionSDate : " & strSessionSDate &"<br/>"
                                response.write "strKeyWord : " & strKeyWord &"<br/>"
                            end if
                            
                            if ( strKeyWord <> strTempKeyword ) then '這表示是第一筆, 所以清空同時段訂課判斷
                                dim bolHasReserved : bolHasReserved = false
                            end if

                            '判斷大會堂已訂位 背景變成藍色
                            if ( 1 = intSessionStatus ) then
                                strChecked = " checked  " '大會堂已訂課 背景變藍色
                                bolHasReserved = true '大會堂已訂課 
                            else
                                strChecked = "" '大會堂已訂課 背景變藍色
                            end if 
                            '第8 9天 或 一對一/一對多  不能OpenMsg 
                            if ( true = bolDebugMode ) then
                                response.Write "intSessionType: " & intSessionType & "<br/>"
                                response.Write "intNowPage: " & intNowPage & "<br/>"
                                response.Write "intK: " & intK & "<br/>"
                             end if
                             
							if ( "N" = re_valArr(0) and true = isDate(re_valArr(1)) and (DateDiff("n", re_valArr(1), strSessionSDate &" "& intSessionSTime &":30" ) < 30) ) then
								strShowInfo =  "<div class='r_session " & strChecked & " ' style='color:#aaa;cursor:hand;text-decoration:none;' onclick='alert(""亲爱的客户，您的电脑尚未完成测试！\n您预约的电脑测试时间：" &  getFormatDateTime( re_valArr(1) ,  1) & "\n为确保您的上课品质，请于电脑测试完成后再进行课程预约。"")'><div class='title_c' title='" & strShowMouseOver &"'>"
                            elseif ( (intNowPage = 2 AND intK >2) OR intK<=2 ) then
                                if ( 4 = sCint(intSessionType) OR 7 = sCint(intSessionType) OR 14 = sCint(intSessionType) ) then '判斷一對一/一對多
                                    bolSession = false
                                    if ( 4 = sCint(intSessionType) OR 14 = sCint(intSessionType) ) then
                                        'strNotReservedDetail = "您已预订同时段一对多课程。"
                                        strNotReservedDetail = "您已预订同时段小班制课程。"
                                        bolSession = true
                                    elseif ( 7 = sCint(intSessionType) ) then
                                        strNotReservedDetail = "您已预订同时段一对一课程。"
                                        bolSession = true
                                    end if 
                                    'if ( true = bolSession OR (strKeyWord = strTempKeyword AND true = bolHasReserved) ) then
                                    '    strShowInfo =  "<div class='r_session ordered' onmouseover=""this.style.backgroundColor='#BBBBBB'""><div class='title_c' title='" & strNotReservedDetail &"'>"
                                    'else
                                    '    strShowInfo =  "<div class='r_session " & strChecked & " ' style='cursor:hand;text-decoration:none;' onclick=""OpenMsgS('" & strSessionSDate & "', '" & intSessionSTime & "', '" & strConsultantEName & "', '" & strImageUrl & "', '" & "" &"', '" & "" &"', '" & strTempSessionLv &"', " & intSessionStatus &", '" & intSessionSn &"', '" & intClientAttendSn &"','"&intConSn&"');""><div class='title_c' title='" & strNotReservedDetail & strShowMouseOver & "');"">"
                                    'end if
                                else 
                                    strNotReservedDetail = ""
                                    bolHaving = false
                                    bolShortSession = false
                                    strSessionWord = ""
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
                                                strNotReservedDetail = strNotReservedDetail & "您已预订 "& strShortSessionSTime & " ~ " & strShortSessionETime &" [" & strSessionWord & "]。" & "&#10;"
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
                                                    strNotReservedDetail = strNotReservedDetail & "您已预订 "& strShortSessionSTime & " ~ " & strShortSessionETime &" [" & strSessionWord & "]。" & "&#10;"
                                                end if
                                            end if
                                        objNewSessionInfo.movenext
                                        wend
                                        objNewSessionInfo.movefirst
                                    else
                                        '判斷第8 9天
                                        strShowInfo =  "<div class='r_session' ><div class='title_c'>"
                                    end if
                                    if (  true = bolSession OR true = bolShortSession OR (strKeyWord = strTempKeyword AND true = bolHasReserved) ) then
                                        strShowInfo =  "<div class='r_session ordered' onmouseover=""this.style.backgroundColor='#BBBBBB'""><div class='title_c' title='" & strNotReservedDetail &"'>"
                                    else
                                        strShowInfo =  "<div class='r_session " & strChecked & " ' style='cursor:hand;text-decoration:none;' onclick=""OpenMsgS('" & strSessionSDate & "', '" & intSessionSTime & "', '" & strConsultantEName & "', '" & strImageUrl & "', '" & "" &"', '" & "" &"', '" & strTempSessionLv &"', " & intSessionStatus &", '" & intSessionSn &"', '" & intClientAttendSn &"','"&intConSn&"');""><div class='title_c' title='" & strNotReservedDetail & strShowMouseOver & "');"">"
                                    end if
                                end if
                            elseif ( strKeyWord = strTempKeyword AND true = bolHasReserved ) then '同時段大會堂 已訂課 非第一筆 判斷
                                strNotReservedDetail = "您已预订同时段大会堂课程。"
                                strShowInfo =  "<div class='r_session ordered' onmouseover=""this.style.backgroundColor='#BBBBBB'""><div class='title_c' title='" & strNotReservedDetail &"'>"
                            else 						
                                '判斷大會堂已訂課 mouseover也要是藍色					
								if ( 1 = intSessionStatus ) then 
                                    strShowInfo =  "<div class='r_session " & strChecked & " ' onmouseover=""this.style.backgroundColor='#D1E3ED'"" style='cursor:hand;text-decoration:none;' onclick=""OpenMsgS('" & strSessionSDate & "', '" & intSessionSTime & "', '" & strConsultantEName & "', '" & strImageUrl & "', '" & "" &"', '" & "" &"', '" & strTempSessionLv &"', " & intSessionStatus &", '" & intSessionSn &"', '" & intClientAttendSn &"','"&intConSn&"');""><div class='title_c' title='" & strShowMouseOver &"'>"
                                else
                                    strShowInfo =  "<div class='r_session " & strChecked & " ' style='cursor:hand;text-decoration:none;' onclick=""OpenMsgS('" & strSessionSDate & "', '" & intSessionSTime & "', '" & strConsultantEName & "', '" & strImageUrl & "', '" & "" &"', '" & "" &"', '" & strTempSessionLv &"', " & intSessionStatus &", '" & intSessionSn &"', '" & intClientAttendSn &"','"&intConSn&"');""><div class='title_c' title='" & strShowMouseOver &"');"">"
                                end if
                            end if
                            
                            '主題太多特殊字元 放到隱藏欄位內
                            response.Write "<input type='hidden' id='" & intSessionSn &"_topic' value='" & strSessionTopic & "' />"
                            '介紹太多特殊字元 放到隱藏欄位內
                            response.Write "<input type='hidden' id='" & intSessionSn &"_intro' value='" & strSessionIntro &"' />"

                            '主題
                            '第8 9天 不顯示課程(主題)
                            if ( intNowPage = 2 AND intK > 2 ) then
                                strSessionTopic = ""
                            else
                                strSessionTopic = strSessionTopic
                            end if

                            strShowInfo = strShowInfo & strSessionTopic
                            strShowInfo = strShowInfo & "</div>" 'end for <div class='title_c'> 

                            '字典檔 未用的值
                            strConsultantFName = arrConsultantFName(intL) 
                            strConsultantLName = arrConsultantLName(intL)
                            intSessionMaxNum = arrSessionMaxNum(intL) 
                            'intSessionNowNum = arrSessionNowNum(intL)
                            intSessionMemberLv = arrSessionMemberLv(intL)
                            intClientAttendType = arrClientAttendType(intL)

                            '是否已訂課tag判斷
                            if ( 1 = intSessionStatus ) then
                                strReservedTag = "<span class = 'reserved'>已订位</span>" '大會堂
                                strIsReserved = "true"
                            elseif ( 4 = sCint(intSessionType) OR 14 = sCint(intSessionType) ) then
                                'strReservedTag = "<span class = 'reserved'>已订位一对三</span>" '一對三
                                strReservedTag = "<span class = 'reserved'>已订位小班制</span>" '小班制
                            elseif ( 7 = sCint(intSessionType)  ) then
                                strReservedTag = "<span class = 'reserved'>已订位一对一</span>" '一對一
                            elseif ( strKeyWord = strTempKeyword AND true = bolHasReserved ) then '大會堂 同時段已訂位
                                strReservedTag = "<span class = 'reserved'>同时段已订位</span>"
                            elseif ( true = bolShortSession ) then
                                strReservedTag = "<span class = 'reserved'>已订位短课程大会堂</span>" '一對一
                            else
                                strReservedTag = ""
                                strIsReserved = "false"
                            end if
                            '串程度字串 原 ,1,2,3...,12 更改為 LV 1-12
                            if ( Not isEmptyOrNull(strSessionApplyLv) ) then
                                strSessionApplyLv = trim(strSessionApplyLv)
                                strShowInfo = strShowInfo & "<div class='lv'>"
				                if ( Left(strSessionApplyLv, 1) = "," ) then
                                    strSessionApplyLv = Right(strSessionApplyLv, Len(strSessionApplyLv)-1)
                                end if
				                if ( Right(strSessionApplyLv, 1) = "," ) then 
                                    strSessionApplyLv = Left(strSessionApplyLv, Len(strSessionApplyLv)-1)
                                end if
				                arrShowSessionApplyLv = split(strSessionApplyLv, ",")
                                '第8 9天 不顯示課程
                                if ( intNowPage = 2 AND intK >2 ) then
                                    strSessionApplyLv = "" '適用程度清空
                                    strReservedTag = "" '已訂位tag清空
				                else
                                    strSessionApplyLv = "适用程度:LV" & arrShowSessionApplyLv(0) & "-" & arrShowSessionApplyLv(UBound(arrShowSessionApplyLv)) '適用程度LV
                                end if 
                                strShowInfo = strShowInfo & strSessionApplyLv
                                strShowInfo = strShowInfo & "</div>" 'end for <div class='lv'>
                                strShowInfo = strShowInfo & strReservedTag 'show已訂課tag
                            end if
                            strShowInfo = strShowInfo & "</div>" ' end for<div class='r_session'>
                            strShowOtherString = strShowOtherString & strShowInfo
                        end if 'end if ( Not isEmptyOrNull(strSessionTopic) ) 
                        strTempKeyword = strKeyWord
                    next ' for intJ
                else '當沒有資料時 一樣要有 r_session div 的css
                    strShowInfo = "<div class='r_session'>"
                    strShowInfo = strShowInfo & "</div>" 'end for <div class='r_session'>  
                    strShowOtherString = strShowOtherString & strShowInfo
                end if 'if ( true = isArray(arrSessionTopic) ) then
                strShowOtherString = strShowOtherString & "</div>" '"<div class='l_r r_0" & intK & "'>"
            end if
        end if
    next 'for intI
    strShowFirstString = strShowFirstString & "</div>" 'end for "<div class='l_c c_h clearfix'>" 'show日期
    strShowOtherString = strShowOtherString & "</div>" 'end for "<div class='l_c c_h clearfix'>" 'show內容
    if (  intI = sCInt(intMinTime) ) then '日期
        response.Write strShowFirstString
    else '其他
        response.Write strShowOtherString
    end if
next
 if ( true = bolDebugMode ) then
     response.write "strCheckDivOpen : " & strCheckDivOpen & "<br/>"
     response.write "bolHasData : " & bolHasData & "<br/>"
 end if
%>
<!--內容End-->
<script type="text/javascript" language="javascript">
 <!--
    //原本OpenMsg()在此，移至最外層lobby_near_order.asp
    $("#<%=strCheckDivOpen%>4").attr("value", "<%=bolHasData%>"); //cal時段查詢變數 special_session_cal.asp
//-->
</script>