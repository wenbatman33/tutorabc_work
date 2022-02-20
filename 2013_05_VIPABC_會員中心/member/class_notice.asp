<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%
Dim str_get_session_sn : str_get_session_sn = getSessionRoomTime()	  '欲找的session_sn
Dim str_session_sn
Dim str_attend_level
Dim str_room
Dim str_session_date
Dim str_session_hour : str_session_hour = right(str_get_session_sn,2)
Dim str_sql
Dim var_arr
Dim arr_result  

        str_session_date = left(str_get_session_sn,4)&"/"&right(left(str_get_session_sn,6),2)&"/"&right(left(str_get_session_sn,8),2)
		var_arr = Array(g_var_client_sn,str_session_date,str_session_hour)
		str_sql = "select Top 1 session_sn,isnull(attend_level,0) attend_level " 
		str_sql = str_sql & " from client_attend_list(nolock) "   
		str_sql = str_sql & " where client_sn=@client_sn and attend_date=@attend_date and attend_sestime = @attend_sestime "   
		str_sql = str_sql & " order by attend_date,attend_sestime "
		arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)    
		'Response.Write("錯誤訊息sql...cccc.:" & g_str_sql_statement_for_debug & "<br>")
		'response.write str_get_session_sn
		'response.end 
		if (isSelectQuerySuccess(arr_result)) then
			   
			if (Ubound(arr_result) >= 0) then
				str_session_sn = arr_result(0,0)
				str_attend_level = arr_result(1,0)
				str_room = int(right(str_session_sn,3)) '教室
			end if
		end if		
	'Dim TypeLib, NewGUID
	'Set TypeLib = CreateObject("Scriptlet.TypeLib")
	'NewGUID = TypeLib.Guid
	'Response.Write(NewGUID)
	'Set TypeLib = Nothing
	
%>
<script language="JavaScript" type="text/JavaScript">
<!--
    var str_room_number = "<%=str_room%>";

    function onclickchImage(p_obj_img_none, p_str_img_ok_id) {
        // 隱藏未勾選的圖
        $(p_obj_img_none).hide();

        // 顯示已勾選的圖
        $("#" + p_str_img_ok_id).show();

        var str_chk_passed_count = $("#hdn_chk_passed_count").val();
        var int_chk_passed_count = 0;

        int_chk_passed_count = parseInt(str_chk_passed_count, 10) + 1;
        str_chk_passed_count = int_chk_passed_count.toString(10);
        $("#hdn_chk_passed_count").val(str_chk_passed_count);

        //三個皆勾選後，直接進入教室
        if (int_chk_passed_count >= 3) {
            var int_room_num = parseInt(str_room_number, 10);
            if (int_room_num <= 550) {
                window.location.href = "go_joinnet_room.asp"
            }
            else {
                window.location.href = "go_tutormeet_room.asp"
                //window.open("class_session.asp?rand=<%=NewGUID%>");
                //window.location.href="session_feedback.asp?mod=online&session_sn=<%=str_session_sn%>";
                //window.location.href = "class_session_test.asp?rand=<%=NewGUID%>";		
            }
        }
    }

//-->	
</script>
<%
	   '======================= Denile 增加頻寬偵測功能 ================= Start ================
	   if (str_session_sn <> "" ) then 
	   'response.write g_str_sql_statement_for_debug
%>
<!--
            <div align="center">
            <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
            id="Bandwidth_Detect" width="0" height="0"
            codebase="http://fpdownload.macromedia.com/get/flashplayer/current/swflash.cab">
            <param name="movie" value="Bandwidth_Detect.swf" />
            <param name="quality" value="high" />
            <param name="bgcolor" value="#869ca7" />
            <param name="allowScriptAccess" value="sameDomain" />
            <param name="wmode" value="opaque" />
            <param name="FlashVars" value="user_sn=<%=(g_var_client_sn+133456)&"&session_sn="&str_session_sn%>" />
            </object>
            </div>
		-->
<%
	   end if 
	   '======================= Denile 增加頻寬偵測功能 ================= end ================
%>
<!--內容start-->
<div id="menber_membox">
    <input type="hidden" id="hdn_chk_passed_count" value="0" />
    <!--收藏課程start-->
    <div class="main_mylist">
        <ul>
            <li class="color_1">Welcome to the session!</li>
            <li><span class="normal">
                <%=getWord("class_notice_1")%></span></li>
            <li class="t-right">
                <input type="button" name="button2" id="button2" value="+ <%=getWord("class_notice_4")%> / <%=getWord("class_notice_5")%>"
                    class="btn_3" onclick="location.href='class.asp'" /></li>
        </ul>
    </div>
    <div class="box_shadow" style="background-position:-80px 0px;"></div> 
    <!--標題end-->
    <!--列表內容-->
    <div class="welcome">
        <span class="color_1">
            <%=getWord("class_notice_2")%></span>，<%=getWord("class_notice_3")%>
        <!--硬件start-->
        <div class="box">
            <div class="left">
                <div class="top">
                    <img src="/images/images_cn/welcome_check_none.gif" onclick="onclickchImage(this, 'img_chk_1');" />
                    <img id="img_chk_1" src="/images/images_cn/welcome_check_ok.gif" style="display: none;" />
                    <%=getWord("class_notice_29")%></div>
                <div class="photo">
                    &nbsp;</div>
            </div>
            <div class="right">
                <span class="color_2">1.</span>
                <%=getWord("class_notice_7")%><br />
                <span class="color_2">2.</span>
                <%=getWord("class_notice_25")%><br />
                <span class="color_2">3.</span>
                <%=getWord("class_notice_9")%><br />
                <span class="color_2">4.</span>
                <%=getWord("class_notice_26")%><br />
                <span class="color_2">5.</span>
                <%=getWord("class_notice_11")%>
            </div>
            <div class="clear">
            </div>
        </div>
        <!--硬件end-->
        <!--軟件start-->
        <div class="box">
            <div class="left">
                <div class="top">
                    <img src="/images/images_cn/welcome_check_none.gif" onclick="onclickchImage(this, 'img_chk_2');" />
                    <img id="img_chk_2" src="/images/images_cn/welcome_check_ok.gif" style="display: none;" />
                    <%=getWord("class_notice_30")%></div>
                <div class="photo2">
                    &nbsp;</div>
            </div>
            <div class="right">
                <span class="color_2">1.</span>
                <%=getWord("class_notice_14")%><br />
                <span class="color_2">2.</span>
                <%=getWord("class_notice_27")%><br />
                <span class="color_2">3.</span>
                <%=getWord("class_notice_17")%><br />
                <span class="color_2">4.</span>
                <%=getWord("class_notice_28")%><br />
            </div>
            <div class="clear">
            </div>
        </div>
        <!--軟件end-->
        <!--上課start-->
        <div class="box">
            <div class="left">
                <div class="top">
                    <img src="/images/images_cn/welcome_check_none.gif" onclick="onclickchImage(this, 'img_chk_3');" />
                    <img id="img_chk_3" src="/images/images_cn/welcome_check_ok.gif" style="display: none;" />
                    <%=getWord("class_notice_31")%></div>
                <div class="photo3">
                    &nbsp;</div>
            </div>
            <div class="right">
                <span class="color_2">1.</span>
                <%=getWord("class_notice_20")%><br />
                <span class="color_2">2.</span>
                <%=getWord("class_notice_21")%><br />
                <span class="color_2">3.</span>
                <%=getWord("class_notice_23")%><br />
                <span class="color_2">4.</span>
                <%=getWord("CLASS_NOTICE_24")%><br />
            </div>
            <div class="clear">
            </div>
        </div>
        <!--上課end-->
    </div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<% 
'20101012 阿捨新增 寫入出席狀態
dim arrParam : arrParam = null
dim strSql : strSql = ""
dim arrSqlResult : arrSqlResult = ""

arrParam = Array(9, str_session_sn, g_var_client_sn)
strSql = "UPDATE client_attend_list SET " 
strSql = strSql & "class_yesno = @class_yesno "
strSql = strSql & "WHERE (client_attend_list.session_sn = @session_sn) AND (client_attend_list.client_sn = @client_sn)"
arrSqlResult = excuteSqlStatementWrite(strSql, arrParam, CONST_VIPABC_R_CONN)
if (isSelectQuerySuccess(arrSqlResult)) then
	if (arrSqlResult < 0) then 
		'寫入失敗
	else
		'寫入成功
	end if
end if
%>