<!--#include virtual="/lib/include/global.inc"-->
<!--#include virtual="/program/class/lobby/LobbyUtility.asp"-->
<%
''檢查某客戶訂大會堂 是否已經重覆上課？
' 從會員點選某個大會堂(Lobby_session & client_attend_list)，判斷該筆 Lobby_session 的 教材編號 material是否重覆了
'
' 沒有資料回傳 0
' 不然回傳 最新一筆的時間

Dim client_sn : client_sn = getSession("client_sn" , 1)
Dim ele_val_str : ele_val_str = getRequest("ele_val_str" , 1)

if not isEmptyOrNull( client_sn ) then
	Set objPCInfo = Server.CreateObject("TutorLib.UtilityLib.DBOperation.DBCmdOp")    '建立連線
	Dim Lobby : set Lobby = new LobbyUtility
	
	obj_count = Lobby.getLobbyIsOrder( objPCInfo , client_sn , ele_val_str)
	if not objPCInfo.EOF then
		d = objPCInfo("attend_date")
		t = objPCInfo("attend_sestime")

		'd = left(d , 10)
        d = getFormatDateTime(d,5)
		t = right( "0"&t , 2 ) &":00"
		
		response.write( d & " " & t )
	else
		response.write( "0" )
	end if
	
	set Lobby = nothing
	objPCInfo.close
	set objPCInfo = nothing
else	
	response.write( "0" )
end if
%>