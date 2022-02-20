<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include virtual="/backoffice/inc/conn.asp"-->
<!--#include virtual="/lib/functions/functions_free_lobby_session.asp"-->
<script type="text/javascript" src="http://<%=CONST_TUTORMEET_SERVER_NAME%>/js/jcbean.js"></script>
<script type="text/javascript">
function showIntegratedTutorMeet(session_sn, client_sn, attend_room, room_rand_str, class_type)
{
    popupIntegratedTutorMeet(session_sn, client_sn, attend_room, room_rand_str, class_type, "", "abc");
}
</script>
<!--內容start-->
<%	
		Dim int_client_sn : int_client_sn = g_var_client_sn	'client_sn
        Dim str_attend_date, str_attend_sestime, str_attend_consultant ,str_session_sn
		Dim sql, arr_param_data, arr_result, bol_sql_write_result
        Dim str_attend_room : str_attend_room = "" '教室號碼
        Dim str_room_rand : str_room_rand = "" '教室亂數
        Dim strConsultantName : strConsultantName = "" '老師姓名
        Dim strClientName : strClientName = "" '客戶姓名
        Dim arrFreeLobbySessionData : arrFreeLobbySessionData = null '上課相關資訊
        Dim strValidTime : strValidTime = "" '上課時間
        Dim strValidClassName : strValidClassName = "" '課程名稱
        Dim strValidDescription1 : strValidDescription1 = "" '課程描述
        Dim strValidDescription2 : strValidDescription2 = "" '課程描述2
        Dim strDispalyInformation : strDispalyInformation = "" '呈現的上課時間
        Dim arrWeekName '星期幾arr
        Dim strWeekName '星期幾
        arrWeekName = Array("","星期日","星期一","星期二","星期三","星期四","星期五","星期六")
        '如果找不到client_sn 連線逾時 重新登入
        if (isEmptyOrNull(g_var_client_sn)) then
            Call alertGo("您尚未登入会员,请前往登入或立即加入免费会员", "http://"&CONST_VIPABC_WEBHOST_NAME, CONST_ALERT_THEN_REDIRECT)
		    Response.End
        end if
        '找出學生的上課資訊
        sql = "SELECT attend_date, attend_sestime, attend_consultant,session_sn,attend_room,ISNULL(basic_fname,'')+' '+ISNULL(basic_lname,'') as con_ename ,ISNULL(fname,'')+' '+ISNULL(lname,'') as client_ename"&_
              " from free_member_attend_list "&_ 
              " left join con_basic "&_
              " on free_member_attend_list.attend_consultant = con_basic.con_sn "&_
              " left join client_basic "&_ 
              " on free_member_attend_list.client_sn = client_basic.sn "&_
              " where free_member_attend_list.valid=1 and free_member_attend_list.client_sn =@client_sn "
        arr_param_data = Array(int_client_sn)
		arr_result = excuteSqlStatementRead(sql, arr_param_data, CONST_VIPABC_R_CONN)
        'Response.Write("錯誤訊息sql:" & g_str_sql_statement_for_debug & "<br>")
		if (isSelectQuerySuccess(arr_result)) then
			if (Ubound(arr_result) >= 0) then			
				str_attend_date = arr_result(0,0) '上課日期
				str_attend_sestime = arr_result(1,0) '上課時段
				str_attend_consultant = arr_result(2,0) '上課老師id
				str_session_sn = arr_result(3,0) 'session_sn
				str_attend_room = "session"&arr_result(4,0) 
                strConsultantName = arr_result(5,0) '老師姓名
                strClientName = arr_result(6,0) '客戶姓名
                '取得亂數的tutormeet教室
				str_room_rand = GetTutorMeetVar(str_attend_room,str_session_sn,"abc")
                '正規化
                str_attend_date = getFormatDateTime(str_attend_date,5)
             else
		        Call alertGo("您目前没有任何咨询，若您要订位，请至生活英语大会堂订位", "/program/member/reservation_class/expo_reservation.asp", CONST_ALERT_THEN_REDIRECT)
		        Response.End
             end if  
            arrFreeLobbySessionData = getFreeLobbySessionData1(str_attend_date)
            strValidTime = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_TIME)
            strValidClassName = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_CLASS_NAME)
            strValidDescription1 = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION1)
            'strValidDescription2 = arrFreeLobbySessionData(CONST_FREE_LOBBY_SESSION_VALID_DESCRIPTION2)      
         end if
         strDispalyInformation = right(str_attend_date,len(str_attend_date)-5)&" "&arrWeekName(Weekday(str_attend_date))&" "&strValidTime		
%>
<div class="sp_session">
    <div class="topbox"></div>
    <div class="sssbox">
        <div class="line0">
		<div class="area0 area-w4">
            <span class="bold">
                <%=strClientName%></span> 您好！欢迎您加入实用英语大会堂！<br /><br />
            您所预订的课程为：<span class="color_3 bold"><%=strValidClassName%></span>
            <br /><%=strValidDescription1%><br /><br />
            上课时间为：<span class="color_3"><%=strDispalyInformation%></span><br /><br />
            授课顾问为：<%=strConsultantName%>
        </div>
		<div class="area0 area-w3">
            <table width="279" border="0" cellpadding="2" cellspacing="6">
              <tr>
                <td width="13"><img src="/images/images_cn/1.gif" width="10" height="11" /></td>
                <td><span class="STYLE1">本讲座将透过VIPABC在线真人视频平台进行</span></td>
              </tr>
              <tr>
                <td width="13"><img src="/images/images_cn/2.gif" width="10" height="11" /></td>
                <td class="STYLE1">登入帐户为您的电子信箱、密码将透过Email<br />
                  传送给您</td>
              </tr>
              <tr>
                <td width="13"><img src="/images/images_cn/3.gif" width="10" height="11" /></td>
          <td class="STYLE1">若安装与使用过程有任何问题，请致电<br />
                    <span class="STYLE2">VIPABC服务专线</span>:<span class="STYLE2">4006-30-30-30</span></td>
              </tr>
              <tr>
                <td width="13"><img src="/images/images_cn/4.gif" width="10" height="11" /></td>
                <td class="STYLE1">上课前请先进行<span class="STYLE2"><a href="http://www.tutormeet.com/mictest/mictest.html?user_sn=101480&comp_status=abc">平台测试&gt;&gt;&gt;点我测试</a></span></td>
              </tr>
            </table>
          </div>
		</div>  
        <div class="clear">
        </div>
        <div class="ssbtn1">
            <%
            Dim int_attend_time_limit : int_attend_time_limit = 15 '幾分鐘以內才能進教室
            '串出上課的時間
            str_attend_date = getFormatDateTime(str_attend_date&" "&str_attend_sestime&":30",	1)
            '如果小於限制 就秀出button
            '加入上完課之後就不能夠看見的條件
            if  (DateDiff("n", now(),str_attend_date) < int_attend_time_limit) then
            '20100810 joyce 移除'進入諮詢室'按鈕
            %>
            <a href="javascript:void(0)" onclick="showIntegratedTutorMeet('<%=str_session_sn%>','<%=int_client_sn%>','<%=str_attend_room%>','<%=str_room_rand%>',8); return false;">
                <img src="/images/images_cn/ss_btn3.gif"></a>
            <%        
            end if
            %>
        </div>
    </div>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->
