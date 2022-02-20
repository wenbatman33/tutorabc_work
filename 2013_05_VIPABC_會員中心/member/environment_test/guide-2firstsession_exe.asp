<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/lib/functions/functions_global.asp"-->
<%'!--#include file="../reservation_class/include/include_prepare_and_check_data.inc"--%>
<!--#include virtual="/program/class/environment_test/EnvironmentTestService.asp"-->
<!--#include virtual="/program/member/environment_test/environment_test_config.inc" -->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"-->
<!--#include virtual="/lib/functions/functions_email_to.asp" --><%'發送mail人員%>
<!--#include virtual="/lib/functions/functions_instant_demo.asp" --><%'發送mail人員%>
<%
function conn()
	Set conn = Server.CreateObject("ADODB.Connection")
    If ( int_enabled_simple_login_process <> 1 ) Then
        conn.open abc_connectionString

    ElseIf ( int_enabled_simple_login_process = 1 ) Then
        If ( CheckSimpleLoginProcessEnabled ) Then
            conn.open new_abc_connectionString
        Else
            conn.open abc_connectionString
        End If
    End If

end function
%>
<%
Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")  

'紀錄首測重測 進入點  Log Insert
set logService = new EnvironmentTestServiceLog
call logService.AddLogBaseArrCol( objPCInfo , Array("OrderTest_save") , Array( Now() ) )
set logService = nothing

'抄襲E:\web\www.tutorabc.com\asp\member\customer_session_first_exe.asp
Dim strlanguage : strlanguage = getRequest("language", CONST_DECODE_NO)     '語系
Dim strExternalLink : strExternalLink = getRequest("externalLink", CONST_DECODE_NO) '是不是後來外部連結 是=1 不是=空
Dim strTestflag_computer : strTestflag_computer = getRequest("testflag_computer", CONST_DECODE_NO) '重新測試的原應為 換電腦重測?電腦重灌?雜音....

'先將問卷資料存起來等預約電腦測試後再寫入 應為以前邏輯問卷需要預約的資料(sn)
Dim strAnswerValue : strAnswerValue = getRequest("AnswerValue", CONST_DECODE_NO) '存問卷的答案選項編號
Dim strProblemName : strProblemName = getRequest("ProblemName", CONST_DECODE_NO) '存問卷的題目選項編號
Dim add_client_datetime_sn '預約SN編號
Dim strclass_time_date_already_order : strclass_time_date_already_order = "" '本次預約時間=上次預約時間->"1"  不要重複寄信  也不用判斷是否過期

class_time=trim(request("chk"))                             '客戶本次預約時間
date_already_order=trim(request("date_already_order"))      '客戶前次預約時間
'如果客戶沒有修改本次預約測試的時間，表示客戶不修改時間所以要用前一次的時間
if (isEmptyOrNull(class_time)) then
    class_time = date_already_order
    strclass_time_date_already_order = "1"
end if
'如果客戶完沒有預約時間就送出，也沒有預約過，提示訊息並回到上一頁
if (isEmptyOrNull(class_time)) then
    call alertGo("请选择预约时间","", CONST_ALERT_THEN_GO_BACK )'請選擇預約時間
    response.end 
end if

aclient_sn=session("client_sn")
language=request("language")
cust_name=request("cust_name")&" "&request("cust_lame")
date_year=Year(class_time)
date_month=Month(class_time)
date_day=Day(class_time)
date_hour=Hour(class_time)
date_min=Minute(class_time)

class_date=date_year&"/"&date_month&"/"&date_day
YMD=class_date
if ( "1" <> strclass_time_date_already_order ) then '本次預約時間=上次預約時間->"1"  這時就不要判斷是否過期
    '首冊重測 後端判斷時間是否已超過
    if DateDiff("n",Cdate(Year(now())&"-"&Month(now())&"-"&Day(now())&" "&Time()),Cdate(date_year&"-"&date_month&"-"&date_day&" "&date_hour&":"&date_min)) < 30 then
    %>
         <script type="text/javascript">
             alert("您选择的时间已超过"); //您選擇的時間已超過		
		    history.go(-1)  				
	     </script>
    <%
	    response.end
    end if
end if
'判斷是不是首冊的重新測試，如果是不可以在預約同一個時段 alert（你已經預約過）
'第一先判斷預約的同一個時段是不是已經有了
' ' ' Dim bol_is_first_session : bol_is_first_session =false '是否為首測
' ' ' 'Dim str_chk_date_time_sql '確定當時段的SQL
' ' ' 'str_chk_date_time_sql = "select * from add_client_datetime where client_sn ='"&session("client_sn")&"' and YMD='"&getFormatDateTime(YMD,5)&"' and hour1='"&date_hour&"' and min1='"&date_min&"'"
' ' ' Dim g_arr_result '回傳的加總值
' ' ' Dim arr_chk_data 'function getFirstTestQuestionaryPassedStatus 傳入參數
' ' ' Redim arr_chk_data(2)
' ' ' Dim str_questionary_type '問卷類型
' ' ' arr_chk_data(0)="1" '首測問卷
' ' ' arr_chk_data(1)="T" 'tutorABC
' ' ' arr_chk_data(2)="tw" '語系
' ' ' '1: 通過 2: 沒通過 0:沒做
' ' ' if getFirstTestQuestionaryPassedStatus(session("client_sn"),arr_chk_data,CONST_TUTORABC_RW_CONN) <> 1 then
	' ' ' bol_is_first_session = true
	' ' ' str_questionary_type = "1"
' ' ' else
	' ' ' str_questionary_type = "2"
' ' ' end if

Dim bol_is_first_session : bol_is_first_session =true '是否為首測
strSql = "SELECT TOP 1 * FROM add_client_datetime WITH(NOLOCK)  WHERE client_sn =@add_client_datetime_sn  and flag=5  and status=6  ORDER BY sn desc "  'and staff_sn =275 and testflag=1 考慮到 ims建立 故不考慮設限自動化測試
intResult = objPCInfo.excuteSqlStatementEach( strSql , Array( session("client_sn")  ) , CONST_TUTORABC_RW_CONN)
if not objPCInfo.EOF then
	bol_is_first_session = false
end if

'line 97
sqlE="SELECT * FROM customer_first_sesion where mod_type=1"
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


'CS08122614 20090109deney修改增加會自動帶服務客服
'sqlF="SELECT * FROM  client_basic where sn='"&aclient_sn&"'"
sqlF="SELECT a.*, b.staff_sn, a.fname + ISNULL(a.lname,'') ename FROM  client_basic a,client_service_staff b where a.sn='"&aclient_sn&"' and a.service_code=b.service_code"
intResult = objPCInfo.excuteSqlStatementEach( sqlF , Array() , CONST_TUTORABC_RW_CONN)
if not objPCInfo.EOF then
	Email= objPCInfo("email")
	call_phone= objPCInfo("cphone_code")&objPCInfo("cphone")
	client_name= objPCInfo("cname")
	service_staff_sn= objPCInfo("staff_sn")
	ename = objPCInfo("ename")
end if

class_time1=class_date&" "&date_hour&":"&date_min
class_time1 = getFormatDateTime(class_time1,1)
'-----------首測重測新增判斷 為首測還是重測----------
'判斷之前有沒有預約過，如果有update
Dim arr_already_order_test_data '已經預約過得那筆資料
Dim g_arr_result_update '回傳的加總值
arr_already_order_test_data = getAlreadyOrderTestTime(aclient_sn,CONST_TUTORABC_RW_CONN)  'Tree 回傳客戶預定最近一次的預約測試時間
'Dim str_redirect_url : str_redirect_url = "http://"&CONST_ABC_WEBHOST_NAME&"/asp/member/first_time_to_go.asp?language="& Request("language") '要轉址的網址
'str_redirect_url = escape(str_redirect_url)

Dim str_add_client_time_sn '最近一次的問卷sn


if ( Not IsArray(arr_already_order_test_data)) then ''Tree 無回傳客戶預定最近一次的預約測試時間
'如果沒有長一筆預約資料
	Dim flag_is_first_time_session : flag_is_first_time_session = request("test_session_type")
	set rs=server.createobject("adodb.recordset")
	sql = "SELECT top 0 *  from add_client_datetime"

	rs.open sql,conn,3,3
	rs.AddNew
	rs("client_sn") = aclient_sn
	class_date = getFormatDateTime(class_date,5)
	rs("YMD") =class_date
	rs("hour1") = date_hour
	rs("min1") = date_min
	'CS08122614 20090109deney修改增加會自動帶服務客服 --stare
	rs("staff_sn") = service_staff_sn
	'CS08122614 20090109deney修改增加會自動帶服務客服 --end
	rs("flag") = "4"


	if flag_is_first_time_session = "R" then
		rs("testflag") = "2" '重測
		rs("retest_item") = strTestflag_computer '重測的結果項目
		'電腦重新測試的項目 add_client_datetime.retest_item
		'1=換電腦重測,  2=電腦重灌,  3=調整耳麥音量,  4=回音,  5=雜音,  6=測試頻寬,  7=錄影檔無法下載,  8=其他
		
		rs("description") = "[電腦重新測試]["&request("str_computer_problems")&"]預約日期:"&class_date&"。<br/>新增備註:官網預約。<br/><font color=red>客戶</font>&nbsp;預約於"&Now()&"<hr/>"
	else
		rs("description") = "[首次上課測試]預約日期:"&class_date&"。<br/>新增備註:官網預約。<br/><font color=red>客戶</font>&nbsp;預約於"&Now()&"<hr/>"
		rs("testflag") = "1" '首測
	end if
	'官網不管為首測還是重測
	'預約後的狀態: 1=已預約,  2=更改預約,  3=取消預約, 4=未出席/更改預約,  5=未出席/取消預約,  6=已出席/測試通過(已完成測試),  7=已出席/測試未通過,
	rs("status") = "1" '已預約
	rs.update
	rs.movelast

	str_add_client_time_sn =rs("sn")
	rs.close

 
else

	'有預約
	sql = "UPDATE add_client_datetime "&_
	" SET YMD = '"&getFormatDateTime(YMD,5)&"',"&_  
	" hour1 ='"&date_hour&"',"&_ 
	" min1 ='"&date_min&"',"&_
	" description ='"&arr_already_order_test_data(5)&"原預約時間"&arr_already_order_test_data(0)&"更改為"&class_time1&"。<br/>新增備註:官網改約。<br/><font color=red>客戶</font>&nbsp;更改預約時間於"&Now()&"<hr/>',"&_
    " retest_item ='"&strTestflag_computer&"',"&_
    " status = 2"&_ 
	" WHERE sn = '"&arr_already_order_test_data(1)&"'" 
    str_add_client_time_sn = arr_already_order_test_data(1) 'Tree 回傳這筆資料的ＳＮ對應
	g_arr_result_update = excuteSqlStatementWriteQuick(sql, CONST_TUTORABC_RW_CONN)
	'response.write sql
	'response.end
	'回傳二維陣列
	if (g_arr_result_update >= 0) then
        if ( "1" <> strclass_time_date_already_order ) then '本次預約時間=上次預約時間->"1"  不要重複寄信
            if ( bol_is_first_session ) then    '首測才寄首測信件
                '20111019 tree ABC系統信件 --開始--
                set SendMailabc = new VIPSystemSendMail											'建立class物件
                SendMailabc.str_sendlettermail_sn 			= "envTestOrder_0002"							'第0002封信 	主旨：更改預約「首次上課測試」通知		
				if CFG_ENVTEST_DEBUG_MODEL=1 then
					SendMailabc.strEmailToAddr 					= CFG_ENVTEST_MAIL_TO			'客戶的電子郵件
				else
					SendMailabc.strEmailToAddr 					= Email					        '客戶的電子郵件
				end if
				SendMailabc.strEmailTitle					=	"更改预约「首次上课测试」通知"
                SendMailabc.client_ename 					= ename							    '客戶姓名
				SendMailabc.str_now_date                    = arr_already_order_test_data(0)    '原先預約上課日期
                SendMailabc.client_date                     = class_time                        '預約上課日期
                Call SendMailabc.MainSendMail()													'呼叫class物件的函式來發信
                set SendMailabc = nothing														'釋放CLASS物件
                '20111019 tree ABC系統信件 --結束--	
				call sendToIMS()
			else
				call sendToIMS()
				call sendToCustomer()
            end if
        end if

		'Response.Write("影響筆數 " & g_arr_result_update & "<br>")
        '如果有效的預約資料將 session寫入 session("bolEffectiveReservation") = "1"  ---開始---
        session("bolEffectiveReservation") = "1" '有效的預約資料
        '如果有效的預約資料將 session寫入 session("bolEffectiveReservation") = "1"  ---結束---

		'Call alertGo("您預約測試時間為:"& class_time1 &" 客服人員將於該時段致電給您,有任何問題請洽客服專線:02-33659999，謝謝!", "/program/member/member.asp", CONST_ALERT_THEN_REDIRECT )
        
		
		' ' ' if ( "1" = strExternalLink ) then 
            ' ' ' '您預約測試時間為:XXX 客服人員將於該時段致電給您,有任何問題請洽客服專線:02-33659999，謝謝!
            ' ' ' Call alertGo(strLangMemGuide2fexe00003 &":"&class_time1 & strLangMemGuide2fexe00004, "/asp/member/first_time/guide-3environment.asp?externalLink="& strExternalLink &"&language="& Request("language"), CONST_ALERT_THEN_REDIRECT )
			' ' ' 'Call alertGo(strLangMemGuide2fexe00003 &":"&class_time1 & strLangMemGuide2fexe00004, "/asp/member/first_time/dcgsLink.asp?language="& Request("language"), CONST_ALERT_THEN_REDIRECT )
        ' ' ' elseif ( "3" = strExternalLink ) then 
			' ' ' Call alertGo("感謝您的預約！\n客服人員將於【"&class_time1 & "】致電給您進行電腦設備測試，有任何問題亦可洽詢客服專線：02-33659999。\n按下「確定」鈕後，系統將導至「DCGS學習偏好設定」。", "/asp/member/first_time/dcgsLink.asp?language="& Request("language"), CONST_ALERT_THEN_REDIRECT )
		' ' ' else
            ' ' ' Call alertGo(strLangMemGuide2fexe00003 &":"&class_time1 & strLangMemGuide2fexe00004, "/asp/member/first_time/guide-3environment.asp?externalLink=3&language="& Request("language"), CONST_ALERT_THEN_REDIRECT )
			' ' ' 'Call alertGo(strLangMemGuide2fexe00003 &":"&class_time1 & strLangMemGuide2fexe00004, "/asp/member/first_time/dcgsLink.asp?language="& Request("language"), CONST_ALERT_THEN_REDIRECT )
        ' ' ' end if
		Call alertGo("感谢您的预约！ \n客服人员将于【"&class_time1 & "】致电给您进行电脑设备测试，有任何问题亦可洽询客服专线：4006-30-30-22。 \n按下「确定」钮后，系统将导至「DCGS学习偏好设定」。", "/program/member/learning_dcgs.asp", CONST_ALERT_THEN_REDIRECT )
        response.end
	else
		Response.Write("錯誤訊息" & g_str_run_time_err_msg & "<br>")
	end if
end if

sqlGetsn="SELECT TOP (1) sn FROM  add_client_datetime where client_sn = "&aclient_sn&" ORDER BY sn DESC"
set rssqlGetsn=server.createobject("adodb.recordset")
rssqlGetsn.open sqlGetsn,conn,1,1
Getsn=rssqlGetsn("sn")


sql="SELECT * FROM client_customer_care"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,3,3
rs.addnew'------------------------------------------------------addnew
rs("client_sn")=aclient_sn
rs("agent_sn")="99999"
rs("agent")="C,99999"
rs("nextcare")=class_date   
rs("types")="22"
rs("acd_sn")=Getsn
			

rs.update
rs.close
set rs=nothing

if bol_is_first_session then

	'mail_url = "http://" & CONST_ABC_WEBHOST_NAME & "/asp/mail/BookFirstSession/index.asp?BookName=" & Escape(cust_name) & "&Bookdate=" & Escape(class_time)
'tree	Call sendEmailFromUrl(Email, cust_name, "感謝您預約「First Session」測試", mail_url, CONST_TUTORABC_WEBSITE)
        
        '20111019 tree ABC系統信件 --開始--
        set SendMailabc = new VIPSystemSendMail											'建立class物件
        SendMailabc.str_sendlettermail_sn 			= "envTestOrder_0001"							'第0001封信 	主旨：感謝您預約「首次上課測試」		
		SendMailabc.strEmailTitle					=	"感谢您预约「首次上课测试」"		
        'SendMailabc.strEmailToAddr 					= Email					            '客戶的電子郵件
		if CFG_ENVTEST_DEBUG_MODEL=1 then
			SendMailabc.strEmailToAddr 					= CFG_ENVTEST_MAIL_TO			'客戶的電子郵件
		else
			SendMailabc.strEmailToAddr 					= Email					        '客戶的電子郵件
		end if	
        SendMailabc.client_ename 					= ename							    '客戶姓名
        SendMailabc.client_date                     = class_time                        '預約上課日期
		Call SendMailabc.MainSendMail()													'呼叫class物件的函式來發信
        set SendMailabc = nothing														'釋放CLASS物件
        '20111019 tree ABC系統信件 --結束--	
		call sendToIMS()
else
    if ( "1" <> strclass_time_date_already_order ) then '本次預約時間=上次預約時間->"1"  不要重複寄信
		set SendMailabc = new VIPSystemSendMail
		SendMailabc.str_sendlettermail_sn 			= "envTestOrder_0003"							'感谢您预约「电脑重新测试」	
		SendMailabc.strEmailTitle					=	"感谢您预约「电脑重新测试」"
		if CFG_ENVTEST_DEBUG_MODEL=1 then
			SendMailabc.strEmailToAddr 					= CFG_ENVTEST_MAIL_TO			'客戶的電子郵件
		else
			SendMailabc.strEmailToAddr 					= Email					        '客戶的電子郵件
		end if	
		SendMailabc.client_ename 						= ename							'客戶姓名
		SendMailabc.client_date                     	= class_time					'預約上課日期
		Call SendMailabc.MainSendMail()
		set SendMailabc = nothing
		
		call sendToIMS()
    end if
end if
%>
<%
''''''''''''''''''變更上面流程 改為下半流程 Johnny 2012-07-05 END

'首測重測寄給TS人員信
'2012-06-12 Johnny  因為之前的邏輯很亂，也不敢亂刪，故另新寫function 直接被call
function sendToIMS()
	if CFG_ENVTEST_DEBUG_MODEL=1 then
		aEmail 					= CFG_ENVTEST_MAIL_TO			'客戶的電子郵件
		UserName				= CFG_ENVTEST_FORMAL_MAIL_NAME
	else
		aEmail 					= CFG_ENVTEST_FORMAL_MAIL_TO
		UserName				= CFG_ENVTEST_FORMAL_MAIL_NAME
	end if


	set SendMailabc = new VIPSystemSendMail
	if bol_is_first_session then
		SendMailabc.strEmailTitle					=	"客户预约「首次上课测试」"
		SendMailabc.str_sendlettermail_sn 			= "envTestOrder_0004"							'客戶 XX 「電腦首次上課測試/电脑重新测试」預約時間
		SendMailabc.strEmailToAddr					= aEmail
		SendMailabc.client_ename 					= ename	& "「电脑首次上课测试」"				'客戶姓名  & 「電腦首次上課測試」
		SendMailabc.client_date                     = class_time					'預約上課日期
		SendMailabc.str_para1						= session("lead_sn")
		SendMailabc.str_para2						= session("client_sn")
	
		' ' ' mail_url = "http://" & CONST_VIPABC_WEBHOST_NAME & "/program/member/environment_test/mail/BookFirstSession/index_cs.asp?"
		' ' ' mail_url = mail_url & "aName=" & Escape(UserName) & "&client_name=" & Escape(client_name)
		' ' ' mail_url = mail_url & "&client_Email=" & Escape(Email) & "&call_phone=" & Escape(call_phone)
		' ' ' mail_url = mail_url & "&aclass_time1=" & Escape(class_time)
		' ' ' mail_url = mail_url & "&lead_sn=" & session("lead_sn")
		' ' ' mail_url = mail_url & "&client_sn=" & session("client_sn")
		' ' ' Call sendEmailFromUrl(aEmail, UserName, "客戶預約「電腦首次上課測試」預約時間:"& class_time, mail_url, CONST_TUTORABC_WEBSITE)
	else
		SendMailabc.strEmailTitle					=	"客户预约「电脑重新测试」"
		SendMailabc.str_sendlettermail_sn 			= "envTestOrder_0004"							'客戶 XX 「電腦首次上課測試/电脑重新测试」預約時間
		SendMailabc.strEmailToAddr					= aEmail
		SendMailabc.client_ename 					= ename	& "「电脑重新测试」"				'客戶姓名  & 「電腦首次上課測試」
		SendMailabc.client_date                     = class_time					'預約上課日期
		SendMailabc.str_para1						= session("lead_sn")
		SendMailabc.str_para2						= session("client_sn")
	
		' ' ' mail_url = "http://" & CONST_VIPABC_WEBHOST_NAME & "/program/member/environment_test/mail/BookReTest/index_cs.asp?"
		' ' ' mail_url = mail_url & "aName=" & Escape(UserName) & "&client_name=" & Escape(client_name)
		' ' ' mail_url = mail_url & "&client_Email=" & Escape(Email) & "&call_phone=" & Escape(call_phone)
		' ' ' mail_url = mail_url & "&aclass_time1=" & Escape(class_time)
		' ' ' mail_url = mail_url & "&lead_sn=" & session("lead_sn")
		' ' ' mail_url = mail_url & "&client_sn=" & session("client_sn")
		' ' ' Call sendEmailFromUrl(aEmail, UserName, "客戶預約「電腦重新測試」測試時間:"& class_time, mail_url, CONST_TUTORABC_WEBSITE)
	end if
	Call SendMailabc.MainSendMail()
	set SendMailabc = nothing
end function
'首測重測寄給客戶信
'2012-06-12 Johnny 因為之前的邏輯很亂，也不敢亂刪，故另新寫function 直接被call
function sendToCustomer()
	if bol_is_first_session then
		set SendMailabc = new VIPSystemSendMail											'建立class物件
		SendMailabc.strEmailTitle					=	"感谢您预约「首次上课测试」"
        SendMailabc.str_sendlettermail_sn 			= "envTestOrder_0001"							'第0001封信 	主旨：感謝您預約「首次上課測試」						
        'SendMailabc.strEmailToAddr 					= Email					            '客戶的電子郵件
		if CFG_ENVTEST_DEBUG_MODEL=1 then
			SendMailabc.strEmailToAddr 					= CFG_ENVTEST_MAIL_TO			'客戶的電子郵件
		else
			SendMailabc.strEmailToAddr 					= Email					        '客戶的電子郵件
		end if	
        SendMailabc.client_ename 					= ename							    '客戶姓名
        SendMailabc.client_date                     = class_time                        '預約上課日期
		' ' ' Call SendMailabc.MainSendMail()													'呼叫class物件的函式來發信
        set SendMailabc = nothing														'釋放CLASS物件
	else
		set SendMailabc = new VIPSystemSendMail
		SendMailabc.strEmailTitle					=	"感谢您预约「电脑重新测试」"
		SendMailabc.str_sendlettermail_sn 			= "envTestOrder_0003"							'感谢您预约「电脑重新测试」	
		if CFG_ENVTEST_DEBUG_MODEL=1 then
			SendMailabc.strEmailToAddr 					= CFG_ENVTEST_MAIL_TO			'客戶的電子郵件
		else
			SendMailabc.strEmailToAddr 					= Email					        '客戶的電子郵件
		end if	
		SendMailabc.client_ename 						= ename							'客戶姓名
		SendMailabc.client_date                     	= class_time					'預約上課日期
		Call SendMailabc.MainSendMail()
		set SendMailabc = nothing
	
		' ' ' mail_url = "http://" & CONST_VIPABC_WEBHOST_NAME & "/program/member/environment_test/mail/BookReTest/index.asp?BookName=" & Escape(cust_name) & "&Bookdate=" & Escape(class_time)
		' ' ' Dim Email_tmp
		' ' ' if CFG_ENVTEST_DEBUG_MODEL=1 then
			' ' ' Email_tmp 					= CFG_ENVTEST_MAIL_TO			'客戶的電子郵件
		' ' ' else
			' ' ' Email_tmp 					= Email					        '客戶的電子郵件
		' ' ' end if
	    ' ' ' ' ' ' Call sendEmailFromUrl(Email_tmp, cust_name, strLangMemGuide2fexe00005&"", mail_url, CONST_TUTORABC_WEBSITE)    'strLangMemGuide2fexe00005=感謝您預約「電腦重新測試」測試
	end if
end function
%>
<%
objPCInfo.close                               '關閉物件
Set objPCInfo = nothing                       '設定物件為空

Call alertGo("感谢您的预约！ \n客服人员将于【"&class_time1 & "】致电给您进行电脑设备测试，有任何问题亦可洽询客服专线：4006-30-30-22。 \n按下「确定」钮后，系统将导至「DCGS学习偏好设定」。", "/program/member/learning_dcgs.asp", CONST_ALERT_THEN_REDIRECT )
%>