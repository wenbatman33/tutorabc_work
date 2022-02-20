<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/md5.asp"-->
<%

dim bolDeBugMode : bolDeBugMode = false '除錯模式
dim bolTestCase : bolTestCase = false '測試模式

if ( g_var_client_sn = "" ) then
    'Session("client_sn") = Request.Cookies("client_sn")
    g_var_client_sn = Request("client_sn")
end if

'如果session 掉了 導回首頁
if ( g_var_client_sn = "" ) then
	response.Redirect "http://"&CONST_VIPABC_WEBHOST_NAME
end if

Dim var_arr						'小幫手提供的conn開法所需變數
Dim arr_result					'小幫手提供的conn開法所需變數
Dim str_sql						'小幫手提供的conn開法所需變數
Dim str_tablecol
Dim str_tablevalue

function extendTree(ByVal qlNumber, ByVal parent, ByVal divLevel, ByVal typeOfInput, ByVal valueOfClassmate, ByVal valueOfQuestion, ByVal actionType, ByVal recordOfQuestion)
'/****************************************************************************************
'描述      ：產生題目的答案選項[以遞迴的方法]
'傳入參數  ：[qlNumber:Integer] 問卷編號 - questionary_list.ql_number
'			 [parent:Integer] 題目編號 - questionary_question_list.qql_number
'            [divLevel:Integer] 選項遞迴的階層
'            [typeOfInput:Array] 每一選項階層顯示的型別 - 例:Array("0","radio","checkbox","radio")
'			 [valueOfClassmate:Integer] 學生的編號 - 從1開始
'			 [valueOfQuestion:Integer] 題目的編號 - 從1開始
'			 [actionType:String] 儲存或是顯示 - "show" or "save"
'			 [recordOfQuestion:String] 儲存的結果字串, 用來做比對是否有選取
'回傳值    ：回傳 1.html tag 或 2.儲存的字串
'牽涉數據表： 
'歷程(作者/日期/原因) ：[Beson] 2011/12/14 Created
'*****************************************************************************************/
  Dim resData : resData = ""
  '找出這問題所有選項
	sql = "SELECT qqil.qqi_number,qqil.qqi_description,qqil.qqi_answer "
	sql = sql & "FROM questionary_list ql "
	sql = sql & "LEFT JOIN questionary_question_list  qql ON qql.ql_number = ql.ql_number "
	sql = sql & "LEFT JOIN questionary_question_item_list qqil ON qqil.qql_number = qql.qql_number "
	sql = sql & "WHERE ql.ql_number="&qlNumber&" AND qqil.qqi_description IS NOT NULL "
	sql = sql & "AND qqil.qqi_parent_number = '"& parent &"' ORDER BY qqil.qqi_sort_sn"
  rsm = excuteSqlStatementReadQuick(sql, CONST_VIPABC_RW_CONN)

  if ( isArray(rsm) ) then
      if (Ubound(rsm) > 0) then
	    if actionType = "show" then resData =resData&"<div id=""level_"&divLevel&""" style=""padding:0 0 0 "&(divLevel-1)*10&"px;"">"
	    for valueOfInput=0 to Ubound(rsm,2)
	      if actionType = "show" then
		    if instrrev(recordOfQuestion, ","&rsm(2,valueOfInput)&",", -1 , 1)>0 then isChecked=" checked" else isChecked=""
		    if isnup then isDisable = " disabled=true"
		    resData = resData & "<span style=""padding:0 1px 0 1px;""><input type="""&typeOfInput(divLevel)&""" name=""level_"&valueOfClassmate&"_"&valueOfQuestion&"_"&divLevel&"_"&parent&""" id=""level_"&valueOfClassmate&"_"&valueOfQuestion&"_"&divLevel&"_"&parent&""" value="""&rsm(2,valueOfInput) &""""&isChecked&isDisable&"  />&nbsp;" & rsm(1,valueOfInput)&"&nbsp;&nbsp;</span>"
	      elseif actionType = "save" then
		    if valueOfInput =0 then resData = resData & request("level_"&valueOfClassmate&"_"&valueOfQuestion&"_"&divLevel&"_"&parent)&"," '因input tag都是同名字, 所以取一個就好了.
	      end if
	      resData = resData & extendTree(qlNumber,rsm(0,valueOfInput),divLevel+1,typeOfInput,valueOfClassmate,valueOfQuestion,actionType,recordOfQuestion)
	    next
	    if actionType = "show" then resData = resData & "</div>"

      end if
  end if
  extendTree = resData
end function

'feedback的參數
Dim str_session_sn	: str_session_sn 		= request("session_sn")
Dim str_consultant	: str_consultant 		= request("consultant")
Dim str_speed			: str_speed			= request("speed")
Dim str_distribution	: str_distribution 	= request("distribution")
Dim str_tskill		: str_tskill			= request("tskill")
Dim str_behavior		: str_behavior		= request("behavior")
Dim str_materials		: str_materials		= request("materials")
Dim str_difficulty	: str_difficulty		= request("difficulty")
Dim str_requirement	: str_requirement		= request("requirement")
Dim str_overall		: str_overall			= request("overall")
Dim str_connection	: str_connection		= request("connection")
Dim str_suggestions	: str_suggestions		= request("suggestions")
Dim str_compliment	: str_compliment		= request("compliment")
Dim str_help_note		: str_help_note		= request("help_note")

'feedback_classmate
Dim var_arr_classmate
Dim str_sql_classmate
Dim arr_result_classmate
Dim var_tmp_condition_value
Dim var_tmp_performance_value
Dim i

'課後評鑑Log
dim objLogOpt : objLogOpt = null
dim strFileName : strFileName = "" 'log檔名 
dim arrColumn : arrColumn = "" '資料欄位名稱
dim arrData : arrData = "" 'log資料
dim intCount : intCount = 0 '欄位筆數
dim strData : strData = "" 'log資料串

'目前日期時間
currentDateTime = year(now)&"/"&month(now)&"/"&day(now)&" "&right("0"&hour(now),2)&":"&right("0"&minute(now),2)

'---新版的儲存方法 Start---
qlNumberOfEvaluation = getRequest("qlNumberOfEvaluation",CONST_DECODE_NO)
qlNumberOfClassmate = getRequest("qlNumberOfClassmate",CONST_DECODE_NO)
'取得課後評比結果
sql = "SELECT qql_number FROM questionary_question_list WHERE ql_number='"&qlNumberOfEvaluation&"' ORDER BY ql_sort_sn"
Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
intSqlWriteResult = objPCInfoInsert.excuteSqlStatementEachQuick(sql, CONST_VIPABC_RW_CONN)
if (not objPCInfoInsert.eof) then
	for i=1 to objPCInfoInsert.recordcount
		if i = 1 then str_consultant = Replace(Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ",""),",","")
		if i = 2 then str_speed = Replace(Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ",""),",","")
		if i = 3 then str_distribution = Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ","")
		if i = 4 then str_tskill = Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ","")
		if i = 5 then str_behavior = Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ","")
		if i = 6 then str_materials = Replace(Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ",""),",","")
		if i = 7 then str_difficulty = Replace(Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ",""),",","")
		if i = 8 then str_requirement = Replace(Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ",""),",","")
		if i = 9 then str_overall = Replace(Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ",""),",","")
		if i = 10 then str_connection = Replace(extendTree(qlNumberOfEvaluation,objPCInfoInsert("qql_number"),1,Array("0","radio","checkbox","radio"),1,i,"save","")," ","")
		objPCInfoInsert.MoveNEXT
	next
end if
objPCInfoInsert.close
Set objPCInfoInsert = nothing

rdata=str_requirement
if request("rn")<>"" then rdata=rdata&", 10|"&request("rn")&"|"
if str_consultant="" or str_materials="" or str_overall="" or str_session_sn="" then
	response.write "<script language='javascript'>"
	response.write "window.history.go(-1);"
	response.write "</script>"
	response.end
else
	Set objPCInfo1 = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
	sql = "SELECT session_sn FROM client_session_evaluation WHERE session_sn='"&str_session_sn&"' AND client_sn="&g_var_client_sn
	intSqlWriteResult = objPCInfo1.excuteSqlStatementEachQuick(sql, CONST_VIPABC_RW_CONN)
	if objPCInfo1.eof then '沒資料代表沒做過評鑑
		
		if request("C1")="ON" then
			str_tablecol = "session_sn, client_sn, consultant_points, materials_points, overall_points, compliment, suggestions, behavior, distribution, speed, tskill, requirement, difficulty, connection, contact, cttime"
			str_tablevalue = "@session_sn, @client_sn, @consultant_points, @materials_points, @overall_points, @compliment, @suggestions, @behavior, @distribution, @speed, @tskill, @requirement, @difficulty, @connection, @contact, @cttime"
			var_arr = Array(str_session_sn, g_var_client_sn, str_consultant, str_materials, str_overall, str_compliment, str_suggestions, str_behavior, str_distribution, str_speed, str_tskill, rdata, str_difficulty, str_connection,1,request("ctime"))
			arrColumn = array("session_sn", "client_sn", "consultant_points", "materials_points", "overall_points", "compliment", "suggestions", "behavior", "distribution", "speed", "tskill", "requirement", "difficulty", "connection", "contact", "cttime")
		else
			str_tablecol = "session_sn, client_sn, consultant_points, materials_points, overall_points, compliment, suggestions, behavior, distribution, speed, tskill, requirement, difficulty, connection "
			str_tablevalue = "@session_sn, @client_sn, @consultant_points, @materials_points, @overall_points, @compliment, @suggestions, @behavior, @distribution, @speed, @tskill, @requirement, @difficulty, @connection "
			var_arr = Array(str_session_sn, g_var_client_sn, str_consultant, str_materials, str_overall, str_compliment, str_suggestions, str_behavior, str_distribution, str_speed, str_tskill, rdata, str_difficulty, str_connection)
			arrColumn = array("session_sn", "client_sn", "consultant_points", "materials_points", "overall_points", "compliment", "suggestions", "behavior", "distribution", "speed", "tskill", "requirement", "difficulty", "connection")
		end if
		arr_result = excuteSqlStatementWrite("Insert into client_session_evaluation ("&str_tablecol&") values("&str_tablevalue&")" , var_arr, CONST_VIPABC_RW_CONN)           

		Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")
		
		strFileName = "log_FeedBack"
		arrData = var_arr
		for intCount = 0 to UBound(arrColumn)
			if ( intCount > 0 ) then
				strData = strData & ","
			end if
			strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
		next

		strData = strData & vbCrLF
		Call objLogOpt.WriteLog(strFileName, strData)
		Set objLogOpt = nothing

		'儲存其他同學的評鑑 Start ---
		Set objPCInfoInsert = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
		sql = " SELECT  ClassRecord.client_sn , "
		sql = sql & "         ISNULL(member.Profile.ProfileFirstName, '') + ' ' + LEFT(ISNULL(member.Profile.ProfileLastName, ''), 1) AS cname , "
		sql = sql & "         classmate_session_feedback.condition , "
		sql = sql & "         classmate_session_feedback.performance "
		sql = sql & " FROM    member.ClassRecord "
		sql = sql & "         LEFT OUTER JOIN classmate_session_feedback ON ClassRecord.client_sn = classmate_session_feedback.client_sn "
		sql = sql & "                                                       AND member.ClassRecord.session_sn = dbo.classmate_session_feedback.session_sn "
		sql = sql & "                                                       AND classmate_session_feedback.inputer IN ( "&g_var_client_sn&" ) "
		sql = sql & "         LEFT JOIN member.Profile ON member.Profile.client_sn = member.ClassRecord.client_sn "
		sql = sql & " WHERE   ( ClassRecordValid = 1 ) "
		sql = sql & "         AND member.ClassRecord.ClassStatusId = 3 " '該同學有出席課程
		sql = sql & "         AND ( member.ClassRecord.session_sn = '"&str_session_sn&"' ) "
		sql = sql & "         AND ( NOT ( ClassRecord.client_sn IN ( "&g_var_client_sn&" ) )) "
		sql = sql & " ORDER BY ClassRecord.client_sn "
		intSqlWriteResult = objPCInfoInsert.excuteSqlStatementEachQuick(sql, CONST_VIPABC_RW_CONN)
		valueOfClassmate = 1
		while not objPCInfoInsert.eof
			sql = "SELECT qql_number FROM questionary_question_list WHERE ql_number='"& qlNumberOfClassmate &"' ORDER BY ql_sort_sn"
			Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
			intSqlWriteResult = objPCInfo.excuteSqlStatementEachQuick(sql, CONST_VIPABC_RW_CONN)
			condition=""
			performance=""
			if (not objPCInfo.eof) then
				for i=1 to objPCInfo.recordcount
					if i=1 then condition = extendTree(qlNumberOfClassmate,objPCInfo("qql_number"),1,Array("0","radio","checkbox","radio"),valueOfClassmate,i,"save","")
					if i=2 then performance = extendTree(qlNumberOfClassmate,objPCInfo("qql_number"),1,Array("0","radio","checkbox","radio"),valueOfClassmate,i,"save","")
					objPCInfo.MoveNEXT
				next
			end if
			objPCInfo.close
			Set objPCInfo = nothing
			condition = ","&Replace(condition," ","")
			performance = ","&Replace(performance," ","")
			
			'濾掉沒填寫的資料
			if (not isEmptyOrNull(Replace(condition,",",""))) and (not isEmptyOrNull(Replace(performance,",",""))) then 
				sql1 = "SELECT inputer "
				sql1 = sql1 & "	 , client_sn "
				sql1 = sql1 & "	 , session_sn "
				sql1 = sql1 & "	 , condition "
				sql1 = sql1 & "	 , performance "
				sql1 = sql1 & "FROM classmate_session_feedback "
				sql1 = sql1 & "WHERE (inputer = "& g_var_client_sn &") "
				sql1 = sql1 & "	AND (session_sn = '"& str_session_sn &"') "
				sql1 = sql1 & "	AND (client_sn = "& objPCInfoInsert("client_sn") &") "
				Set objPCInfoInsert1 = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")
				intSqlWriteResult = objPCInfoInsert1.excuteSqlStatementEachQuick(sql1, CONST_VIPABC_RW_CONN)
				if objPCInfoInsert1.eof then '沒資料代表沒對其他學生做過評鑑
					str_tablecol = "inputer, client_sn, session_sn, condition, performance"
					str_tablevalue = "@inputer, @client_sn, @session_sn, @condition, @performance"
					var_arr = Array(g_var_client_sn, objPCInfoInsert("client_sn"), str_session_sn, condition, performance)
					arr_result = excuteSqlStatementWrite("Insert into classmate_session_feedback ("&str_tablecol&") values("&str_tablevalue&")" , var_arr, CONST_VIPABC_RW_CONN)
					arrColumn = array("inputer", "client_sn", "session_sn", "condition", "performance")
				end if
				objPCInfoInsert1.close
				Set objPCInfoInsert1 = nothing

				Set objLogOpt = Server.CreateObject("TutorLib.UtilityLib.LogOperation.LogOperation")
		
				strFileName = "log_FeedBack_classmate"
				arrData = var_arr
				for intCount = 0 to UBound(arrColumn)
					if ( intCount > 0 ) then
						strData = strData & ","
					end if
					strData = strData & arrColumn(intCount) & "=" & arrData(intCount)  
				next

				strData = strData & vbCrLF
				Call objLogOpt.WriteLog(strFileName, strData)
				Set objLogOpt = nothing
			end if
			valueOfClassmate = valueOfClassmate + 1
			objPCInfoInsert.MoveNEXT
		wend
		objPCInfoInsert.close
		Set objPCInfoInsert = nothing
		'儲存其他同學的評鑑 End ---
	end IF
	objPCInfo1.close
	set objPCInfo1=nothing

'---新版的儲存方法 End---

        '20110614 阿捨新增 特殊客戶關懷
        '20120926 阿捨新增 特殊客戶關懷include檔
        dim intNoifiyType : intNoifiyType = ""
        dim strSessionDate : strSessionDate = ""
        dim intConsultantScore : intConsultantScore = str_consultant
        dim intMaterialsScore : intMaterialsScore = str_materials
        dim intOverallScore : intOverallScore = str_overall
        dim strCompliment : strCompliment = str_compliment
        dim strSuggestions : strSuggestions = str_suggestions
        intNoifiyType = 3
        strSessionDate = left(str_session_sn,4)&"/"&right(left(str_session_sn,6),2)&"/"&right(left(str_session_sn,8),2)&"  "&right(left(str_session_sn,10),2)&":30"
        strSessionDate = right(strSessionDate, 15)
        if ( bolDeBugMode ) then
            response.Write "strSessionDate: " & strSessionDate & "<br />"
        end if

    Call alertGo(getWord("feedback_send"), "session_feedback_success_jr.asp", CONST_ALERT_THEN_REDIRECT)
end if
%>
