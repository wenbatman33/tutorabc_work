<!--#include virtual="/lib/include/global.inc" -->
<%
Dim str_encodestr : str_encodestr = request("encodestr")
if(str_encodestr<> "") then
	Dim str_record_url
	Dim arr_para
	Dim arr_recording_result
	Dim str_record_sql : str_record_sql ="SELECT TOP (1) sf.SN, sf.File_Full_name, CONVERT(varchar, sf.Session_date, 111) AS Session_date, cfg_sessionfile_path.path_text, client_attend_list.attend_mtl_1 FROM SessionRecord_fileinfo AS sf INNER JOIN cfg_sessionfile_path ON sf.File_path = cfg_sessionfile_path.sn LEFT OUTER JOIN client_attend_list ON sf.Session_Record_Number = client_attend_list.session_sn WHERE (sf.encodestr = @encodestr)"
		arr_para = Array(str_encodestr)
		arr_recording_result = excuteSqlStatementRead(str_record_sql, arr_para , CONST_VIPABC_R_CONN)
		if(isSelectQuerySuccess(arr_recording_result) and Ubound(arr_recording_result) >= 0) then
			str_record_url = "http://" & CONST_RECORDINGVIDEO_WEBHOST_NAME & "/" & arr_recording_result(3,0) & Replace(arr_recording_result(2,0),"/","_") & "/" & arr_recording_result(1,0)
			'Response.write str_record_url
			Response.write "<script language='javascript1.2'>window.open('"& str_record_url &"');window.opener=null;window.close();</script>"
		end if
end if
%>