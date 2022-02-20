<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>
<%
dim str_check_test_sql
dim StrToMail
dim var_arr
dim arr_result
dim room
dim YMD
dim hour1
dim attend_room
dim session_sn
dim room_rand_str
dim Room_Link
dim bolTutorMeet : bolTutorMeet = true '新增tutormeet測試教室權限 

var_arr = Array(g_var_client_sn)
'判斷是否有開啟測試教室
str_check_test_sql="select room,YMD,hour1 from add_client_datetime where flag=4 and room is not null and datediff(minute,dateadd(minute,-30,isnull(convert(varchar,YMD)+' '+convert(varchar,hour1)+':'+convert(varchar,min1)+':00','2000/1/1')),getdate())>=0 and datediff(minute,getdate(),dateadd(minute,30,isnull(convert(varchar,YMD)+' '+convert(varchar,hour1)+':'+convert(varchar,min1)+':00','2000/1/1')))>=0  and client_sn=@client_sn"
arr_result = excuteSqlStatementRead(str_check_test_sql, var_arr, CONST_VIPABC_RW_CONN)  
if (isSelectQuerySuccess(arr_result)) then
    if ( Ubound(arr_result) >= 0 ) then
        room = arr_result(0,0)
        YMD = arr_result(1,0)
        hour1 = arr_result(2,0)
        attend_room = "test"&Room
        session_sn = year(YMD)&right("0"&month(YMD),2)&right("0"&day(YMD),2)&right("0"&hour1,2)&"00"&Room       
        '取得教室亂數
        room_rand_str = GetTutorMeetVar(attend_room,session_sn,"abc")
        bolTutorMeet = true
    end if
else 
    bolTutorMeet = false
end if

if ( room <> ""  and room_rand_str <> "" ) then   
    dim str_check_client_sql
    dim cname
    dim ename
    dim now_level

    var_arr = Array(g_var_client_sn)
      '取得學生基本資料
    str_check_client_sql="SELECT cname,isnull(fname,'')+' '+isnull(lname,'') as ename,isnull(nlevel,1) as now_level FROM client_basic  WHERE sn = @sn"
    arr_result = excuteSqlStatementRead(str_check_client_sql, var_arr, CONST_VIPABC_RW_CONN)  
    if (isSelectQuerySuccess(arr_result)) then
        if ( Ubound(arr_result) >= 0 ) then
            cname = arr_result(0,0)
            ename = arr_result(1,0)
            now_level = arr_result(2,0)
        end if
    end if
   
    dim str_insert_tutormeet_sql
    '檢查該User的資料是否轉入TutorMeet，若否，則直接insert進一筆
    str_check_client_sql="SELECT * from tutormeet.tutormeet_fms.dbo.fms_client_audio_setting_info  where client_sn=@g_var_client_sn and comp_status='abc'"
    arr_result = excuteSqlStatementRead(str_check_client_sql, var_arr, CONST_VIPABC_RW_CONN)  
    if (isSelectQuerySuccess(arr_result)) then
        if (Ubound(arr_result)<0) then
            var_arr = Array(g_var_client_sn,cname,ename,now_level)
            str_insert_tutormeet_sql="insert into tutormeet.tutormeet_fms.dbo.fms_client_audio_setting_info (client_sn, cname, ename, mic_setting, volume_setting, nlevel, profile, comp_status, client_status, lang,class_type,comp_status_logo) values(@g_var_client_sn,@cname,@ename,6,6,@now_level,'首/','abc','首/',2,9,'vipabc') "
            arr_result = excuteSqlStatementWrite(str_insert_tutormeet_sql, var_arr, CONST_VIPABC_RW_CONN)
        else
            var_arr = Array(cname,ename,now_level,g_var_client_sn)
            str_insert_tutormeet_sql="update tutormeet.tutormeet_fms.dbo.fms_client_audio_setting_info set cname=@cname&, ename=@ename,nlevel = @now_level, class_type=9,comp_status_logo='vipabc' where client_sn=@g_var_client_sn  and comp_status='abc' "
            arr_result = excuteSqlStatementWrite(str_insert_tutormeet_sql, var_arr, CONST_VIPABC_RW_CONN)
        end if
    end if
end if

'20100823 阿捨新增 VIPCS測試教室 for JoinNet
dim bolEnterJoinNetSession '進入joinnet測試教室
dim arrParam : arrParam = null '傳入sql的array
dim strSql : strSql = "" 'sql
dim arrSqlResult : arrSqlResult = "" 'sql結果 
dim bolJoinNet : bolJoinNet = true '進入joinnet測試教室權限
dim strUrlPath : strUrlPath = "" 'joinnet固定網址
dim intRoomNum : intRoomNum ="" '測試教室編號

strUrlPath = "http://" & CONST_ABC_WEBHOST_NAME & "/jnj_file/"

'判斷客戶是否有測試教室 有帶出教室編號
arrParam = Array(g_var_client_sn)
strSql = "SELECT vip_cs_room "
strSql = strSql & "FROM add_client_datetime "
strSql = strSql & "WHERE client_sn=@client_sn AND vip_cs_room IS NOT NULL order by vip_cs_room"
arrSqlResult = excuteSqlStatementRead(strSql, arrParam, CONST_TUTORABC_RW_CONN)
if ( isSelectQuerySuccess(arrSqlResult) ) then
    if ( Ubound(arrSqlResult) >= 0 ) then
        intRoomNum = arrSqlResult(0,0)
        bolJoinNet = true 
    else
        bolJoinNet = false
    end if
end if

'有教室編號 帶出教室連結
if ( intRoomNum <> "" ) then
    arrParam = Array(intRoomNum)
    strSql = "SELECT url "
    strSql = strSql & "FROM session_room_status "
    strSql = strSql & "WHERE room = @room"
    arrSqlResult = excuteSqlStatementRead(strSql, arrParam, CONST_WEBOFFICE_CONN)
    if ( isSelectQuerySuccess(arrSqlResult) ) then
	    if ( Ubound(arrSqlResult) >= 0 ) then
            strVipCsTestRoomUrl = strUrlPath & arrSqlResult(0,0)
	    else
            
	    end if
    end if
end if

if (false = bolJoinNet AND false = bolTutorMeet) then
    Call alertGo("无开放测试教室", "/program/member/reservation_class/normal_and_vip.asp", CONST_ALERT_THEN_REDIRECT)
end if
%>
<div class="temp_contant">
	<!--內容start-->
    <div class="main_membox">
<%
if room<>""  and room_rand_str<>"" then 
    Room_Link = "javascript:showIntegratedTutorMeet('"&session_sn&"', '"&g_var_client_sn&"', '"&attend_room&"', '"&room_rand_str&"', 9,'','abc')"   
    response.Write "&nbsp;&nbsp;<a href="""& Room_Link &""" title='TutorMeet测试教室'><img src='../../images/images_cn/entertestingroom.gif' border='0'></a>"
end if

if ( true = bolJoinNet AND intRoomNum <> "" ) then
    response.Write "<br />&nbsp;&nbsp;<a href="""& strVipCsTestRoomUrl &""" title='JoinNet测试教室' target='_blank'><img src='../../images/images_cn/entertestingroom.gif' border='0'></a>"
end if
%>
    </div>
	<!--內容end-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
   