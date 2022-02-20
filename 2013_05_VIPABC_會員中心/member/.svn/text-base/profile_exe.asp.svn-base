<!--#include virtual="/lib/include/global.inc" -->
<!--#include virtual="/lib/functions/functions_leadbasic.asp" -->
<!--#include virtual="/lib/functions/functions_VIPSystemSendMail/functions_VIPSystemSendMail.asp"--><!--TREE VIP系統信-->
<!--#include virtual="/lib/functions/functions_global.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>profile</title>
</head>

<body>
<%
dim bolDeBugMode : bolDeBugMode = false '除錯模式

'判斷從什麼地方導過來的：LearningCard
Dim strEvent : strEvent = getRequest("event",CONST_DECODE_NO)
Dim strEventData : strEventData = getRequest("event_data",CONST_DECODE_NO)

'定義變數；接Post參數
'Dim str_email	 'email
Dim str_email : str_email = Trim(Session("client_email"))
Dim str_cname : str_cname = Trim(Request("txt_cname")) '中文姓名
Dim str_fname : str_fname = Trim(Request("txt_fname")) '英文名
Dim str_lname	 : str_lname = Trim(Request("txt_lname"))	 '英文姓
Dim str_sex	 : str_sex = Trim(Request("chk_gender")) '性別 2:女 1:男
Dim str_birth : str_birth = Trim(Request("txt_birth")) '生日
Dim str_clb_province : str_clb_province = Trim(Request("sel_province")) '省份
Dim str_clb_area	 : str_clb_area = Trim(Request("sel_area")) '地區
dim strClbRegion : strClbRegion = getRequest("sel_region",CONST_DECODE_NO) '地區
if ( bolDeBugMode ) then
    response.Write "strClbRegion:" & strClbRegion & "<br/>"
end if
Dim str_addr : str_addr = Trim(Request("txt_addr")) '地址
'20090109 因為在include_address.asp裡面txt_zip設定disabled，所以用hdn_txt_zip傳值
'Dim str_zip_code : str_zip_code = Trim(Request("hdn_txt_zip")) '郵遞區號
dim str_zip_code : str_zip_code = getRequest("txt_zip",CONST_DECODE_NO)

Dim str_cphone_cycode : str_cphone_cycode = Trim(Request("hdn_txt_cphone_cycode")) '第一組手機國碼
Dim str_cphone : str_cphone = Trim(Request("txt_cphone")) '第一組手機號碼

Dim str_cphone_cycode2 : str_cphone_cycode2 = Trim(Request("hdn_txt_cphone_cycode2")) '第二組手機國碼
Dim str_cphone2	 : str_cphone2 = Trim(Request("txt_cphone2")) '第二組手機號碼

Dim str_htel_code : str_htel_code = Trim(Request("txt_gphone_code")) '家用區碼
Dim str_htel : str_htel = Trim(Request("txt_gphone")) '家用電話
Dim str_htel_branch : str_htel_branch = Trim(Request("txt_gphone_ext")) '家用電話分機
Dim str_otel_code : str_otel_code = Trim(Request("txt_ophone_code")) '公司電話區碼
Dim str_otel : str_otel = Trim(Request("txt_ophone")) '公司電話
Dim str_otel_branch : str_otel_branch = Trim(Request("txt_ophone_ext")) '公司電話分機

Dim str_leadbasic_sn : str_leadbasic_sn = ""   'lead_basic.sn

'其他變數
Dim str_cphone_code '手機前四碼
Dim str_cphone_num '手機後六碼
Dim str_cphone_num_2 '第二組手機號碼

'塞資料的相關設定值
Dim var_arr '傳excuteSqlStatementRead的陣列
Dim arr_result '接回來的陣列值
Dim str_tablecol 'table欄位
Dim str_tablevalue 'table值
Dim int_noteflag

'寄信需要的資料'20110214 tree
Dim DateTime : DateTime = Now()	 '現在的時間
Dim strTime : strTime = Left(DateTime,30) '現在的時間字串

'------------------- update cleint_basic ------------ start ----------------------
'整理手機號碼
if ( str_cphone<>"" ) then
	if ( len(str_cphone)>4 ) then
		str_cphone_code = left(str_cphone,4)
		str_cphone_num = right(str_cphone,len(str_cphone)-4)
	end if
end if

'整理通訊地址,大陸不存地區；台灣032
'if ( str_clb_province = "032") then 
	'str_clb_area = NULL
'end if 
	
'第二組手機國碼
if ( not isEmptyOrNull(str_cphone_cycode2)) then 
	str_cphone_cycode2 = replace(str_cphone_cycode2,"+","")
end if 
	
'第一組手機國碼
if ( not isEmptyOrNull(str_cphone_cycode)) then 
	str_cphone_cycode = replace(str_cphone_cycode,"+","")
end if 
	
'公司電話區碼
if ( not isEmptyOrNull(str_otel_code)) then 
	str_otel_code = replace(str_otel_code,"+","")
end if 
	
'家用區碼
if ( not isEmptyOrNull(str_htel_code)) then 
	str_htel_code = replace(str_htel_code,"+","")
end if 
	
str_tablecol = "cname=@cname,fname=@fname,lname=@lname,sex=@sex,birth=@birth,"
str_tablecol = str_tablecol & "clb_province=@clb_province,clb_area=@clb_area,client_address=@client_address,zip_code=@zip_code, "  
str_tablecol = str_tablecol & "cphone_code=@cphone_code,cphone=@cphone,"  '電話欄位
str_tablecol = str_tablecol & "clb_cphone_cycode_2=@clb_cphone_cycode_2,clb_cphone_2=@clb_cphone_2,"
str_tablecol = str_tablecol & "htel_code=@htel_code,htel=@htel,htel_branch=@htel_branch,"
str_tablecol = str_tablecol & "otel_code=@otel_code,otel=@otel,otel_branch=@otel_branch,"
str_tablecol = str_tablecol & "clb_cphone_cycode=@clb_cphone_cycode, "
'20120918 阿捨新增 地址第三層
str_tablecol = str_tablecol & "clb_region = @clb_region "

redim var_arr(21)
var_arr(0) = str_cname '中文姓名
var_arr(1) = str_fname '英文名
var_arr(2) = str_lname '英文姓
var_arr(3) = str_sex	 '性別 2:女 1:男
var_arr(4) = str_birth '生日
var_arr(5) = str_clb_province '省份
var_arr(6) = str_clb_area '地區
var_arr(7) = str_addr '地址
var_arr(8) = str_zip_code '郵遞區號

var_arr(9) = str_cphone_code '手機前四碼
var_arr(10) = str_cphone_num '手機後六號

var_arr(11) = str_cphone_cycode2 '第二組手機國碼
var_arr(12) = str_cphone2 '第二組手機

var_arr(13) = str_htel_code '家用區碼
var_arr(14) = str_htel '家用電話
var_arr(15) = str_htel_branch '家用電話分機

var_arr(16) = str_otel_code '公司電話區碼
var_arr(17) = str_otel '公司電話
var_arr(18) = str_otel_branch '公司電話分機
var_arr(19) = str_cphone_cycode '第一組手機國碼
'20120918 阿捨新增 地址第三層
var_arr(20) = strClbRegion

var_arr(21) = g_var_client_sn	 'client_basic.sn


' TODO：country是否要給予預設值

arr_result = excuteSqlStatementWrite(" UPDATE client_basic SET " & str_tablecol & " WHERE sn=@sn ", var_arr, CONST_VIPABC_RW_CONN)
if ( bolDeBugMode ) then
    Response.Write("錯誤訊息Sql for 更新客戶資訊:" & g_str_sql_statement_for_debug & "<br>")
end if	   
if (arr_result > 0) then 
	'成功
	'======BD08052615 同步更新Lead_basic的資料===== start ============
	'頁面上無此欄位：otherplacetel --> 其他電話，education-->學壢,Focus -->學習重點，mh_idno-->身份證字號，
	'需要同步：性別，生日
	'country?是否要給予預設值
	
	str_tablecol = "sex=@sex, birth=@birth"
	var_arr = Array(str_sex,str_birth,g_var_client_sn)
	arr_result = excuteSqlStatementWrite("Update lead_basic set "&str_tablecol&" where client_sn=@client_sn" , var_arr, CONST_VIPABC_RW_CONN)
	'======BD08052615 同步更新Lead_basic的資料===== end ============
	
	'====== addnew backup_client_record ===== start ================	
	var_arr = Array(str_email)
	arr_result = excuteSqlStatementRead("select sn from lead_basic where email=@email", var_arr, CONST_VIPABC_RW_CONN)

	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then  
			str_leadbasic_sn = arr_result(0,0)
			int_noteflag = WriteNoteToTable (str_leadbasic_sn,0,6,"select * from lead_basic where sn=@sn" )
		else
		end if 
	end if 
	'====== addnew backup_client_record ===== end ================
		
	'-------- update client_basic_h ------------- start ---------
	var_arr = Array(g_var_client_sn)
	arr_result = excuteSqlStatementRead("select * from client_basic_h where client_sn=@client_sn", var_arr, CONST_VIPABC_RW_CONN)

	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) > 0) then  
			'有資料 (做update處理)
			str_tablecol = "Htel_code=@Htel_code,Htel=@Htel"
			var_arr = Array(str_htel_code,str_htel,g_var_client_sn)
			arr_result = excuteSqlStatementWrite("Update client_basic_h set "&str_tablecol&" where client_sn=@client_sn" , var_arr, CONST_VIPABC_RW_CONN)
		else
			'無資料 (做insert處理)
			var_arr = Array(str_htel_code,str_htel,g_var_client_sn)
			arr_result = excuteSqlStatementWrite("Insert into client_basic_h (Htel_code,Htel,client_sn) values(@Htel_code,@Htel,@client_sn)" , var_arr, CONST_VIPABC_RW_CONN)
		end if 
	end if 
	'20110214 tree VB10110205 VIPABC系統信件 --開始--
	set SendMailvip = new VIPSystemSendMail		   '建立class物件
	SendMailvip.str_sendlettermail_sn = "0009"		   '第0009封信 	主旨：用户基本资料修改						
	SendMailvip.strEmailToAddr = str_email			   '客戶的電子郵件
	SendMailvip.client_cname = str_cname				   '客戶姓名
	SendMailvip.client_modify_data_date = strTime '客戶修改資料時間
	Call SendMailvip.MainSendMail()					      '呼叫class物件的函式來發信
	set SendMailvip = nothing									  '釋放CLASS物件
	'20110214 tree VB10110205 VIPABC系統信件 --結束--
	'-------- update client_basic_h ------------- end ---------
    if ( bolDeBugMode ) then
    else
        if ( "LearningCard" = strEvent ) then
            strCellPhone = sendLearninfCardSMS(strEventData,g_var_client_sn,"VIPABC")
            Call alertGo(getWord("upate_finish"), "http://"&const_vipabc_webhost_name&"/program/learningcard/LearningCardPhoneValidation.asp?CellPhone=" & getNumberEncrypt(strCellPhone) & "&RandomKey=" & strEventData & "&ClientSn=" & g_var_client_sn, CONST_ALERT_THEN_REDIRECT )
        else
		    Call alertGo(getWord("upate_finish"), "profile_success.asp", CONST_ALERT_THEN_REDIRECT )
        end if
	    Response.end
    end if
else
	'錯誤訊息(無修改資料)
	Call alertGo("修改資料失敗", "profile.asp", CONST_ALERT_THEN_REDIRECT )
	Response.end
end if  
%>
</body>
</html>
