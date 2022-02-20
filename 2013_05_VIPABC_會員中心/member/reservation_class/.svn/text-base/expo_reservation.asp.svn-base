<!--#include virtual="/backoffice/inc/conn.asp"-->
<!--#include virtual="/lib/include/html/template/template_change_keyword.asp"-->
<!--#include virtual="/backoffice/inc/function.asp"-->
<%
'開始去XML裡面撈相關的地區資料
	'Dim obj_date_config_xml
'	Dim obj_date_node : obj_date_node = null '日期節點
'	Dim obj_time_node : obj_time_node = null '時段節點
'	Dim arr_valid_order_date : arr_valid_order_date = null '能夠訂課的日期陣列
'	Dim arr_valid_order_time : arr_valid_order_time = null '能夠訂課的時間陣列
'	Dim date_select_order_date : date_select_order_date = null '選取的訂課時間
'	Dim i
'    Dim dateTempInput
'    Dim arrValidTime(2) '陣列儲存可以上課的日期
'    Dim arrValidDate(2) '陣列儲存可以上課的日期
'    Dim arrValidDescription1(2) '課程描述1
'    Dim arrValidDescription2(2) '課程描述2
'    Dim arrValidClassName(2) '陣列儲存上課名稱
'    Dim arrValidRoom(2) '上課的教室
'    Dim arrSessionSn(2) 'sessionsn
'    Dim intConutForWorkDate : intConutForWorkDate = 0 'count
'    Dim arrWeekName '星期幾arr
'    Dim strWeekName '星期幾
'    arrWeekName = Array("","星期日","星期一","星期二","星期三","星期四","星期五","星期六")
'    '找出可以上課的天數
'    Set obj_date_config_xml = New XML
'    If (obj_date_config_xml.load(Server.Mappath("/xml/free_attend_list_date_time_config.xml"))) Then
'        for i = 0 to 8
'            dateTempInput = DateAdd("d" , i , date)
'            '正規劃時間格式
'            date_select_order_date = getFormatDateTime(dateTempInput, 5)        
'		    For Each obj_date_node in obj_date_config_xml.getXmlObj.selectNodes("/ProjectType[@value='1']/DateTime/Date[@value='"&date_select_order_date&"']/TimeSlot")
'            '取出當天的時段
'		    arr_valid_order_time = obj_date_node.getAttribute("value")             
'			    '取出兩天可以用的時段來上課
'			    if (not isEmptyOrNull(arr_valid_order_time)) then
'                    '陣列不可以超過2
'                    if (intConutForWorkDate <= 1) then
'                        arrValidTime(intConutForWorkDate) = obj_date_node.getAttribute("value")
'                        arrValidClassName(intConutForWorkDate) = obj_date_node.getAttribute("classname")
'                        arrValidRoom(intConutForWorkDate) = obj_date_node.getAttribute("room")
'                        arrValidDescription1(intConutForWorkDate) = obj_date_node.getAttribute("description1")
'                        arrValidDescription2(intConutForWorkDate) = obj_date_node.getAttribute("description2")
'                        arrValidDate(intConutForWorkDate) = date_select_order_date 
'                        arrSessionSn(intConutForWorkDate) =   Replace(arrValidDate(intConutForWorkDate) , "/", "")&left(arrValidTime(intConutForWorkDate),2)&arrValidRoom(intConutForWorkDate)                      
'                        intConutForWorkDate = intConutForWorkDate+1
'                    end if
'			    end if 		
'		    next	    
'        next
'    end if	
'	Set obj_date_config_xml = Nothing

'====读取最近2次大会堂信息start==========================
Set rs=Server.CreateObject("Adodb.RecordSet")
    sql="select top 2 * from lobbysession where SessionDate>getdate() order by SessionDate asc"
    'response.Write sql
    rs.Open sql,conn,1,1
 if rs.eof or rs.bof then
        SessionDateFirst="01/01"
		SessionTimeFirst="20:00"
		weekfirst="星期二"
		SessionNameFirst="尚无课程"
		SessionTypeFirst=""
		SessionDescriptionFirst=""
		ConsultantIDFirst=""
		ConsultantNameFirst=""
		RoomFirst=""
 else
  i=1
  do while not rs.eof
    if i=1 then
		SessionDateFirst=Format_Time(rs("SessionDate"),5)
		SessionTimeFirst=Format_Time(rs("SessionDate"),6)
		weekfirst=WeekdayName(Weekday(rs("SessionDate"))) 
		SessionNameFirst=rs("SessionName")
		SessionTypeFirst=rs("SessionType")
		SessionDescriptionFirst=rs("SessionDescription")
		ConsultantIDFirst=rs("ConsultantID")
		ConsultantNameFirst=rs("ConsultantName")
		RoomFirst=Format_Time(rs("SessionDate"),7)&rs("Room")
		PictureBig=rs("PictureBig")
    end if
	if i=2 then
		SessionDateSecond=Format_Time(rs("SessionDate"),5)
		SessionTimeSecond=Format_Time(rs("SessionDate"),6)
		weekSecond=WeekdayName(Weekday(rs("SessionDate"))) 
		SessionNameSecond=rs("SessionName")
		SessionTypeSecond=rs("SessionType")
		SessionDescriptionSecond=rs("SessionDescription")
		ConsultantIDSecond=rs("ConsultantID")
		ConsultantNameSecond=rs("ConsultantName")
		RoomSecond=Format_Time(rs("SessionDate"),7)&rs("Room")
		
    end if
     i=i+1
	 rs.movenext
   loop
  rs.close
 end if 
Set rs=nothing
if PictureBig = "" then
    PictureBig = "201272791426980.jpg"
end if	
'====读取最近2次大会堂信息end==========================
    session("exporURL") = "1"
%>
<!--內容start-->
<link rel="stylesheet" type="text/css" href="/lib/javascript/framepage/css/expo_reservation.css" />
<style type="text/css">
<!--
<%'''''''''''''重写大会堂主图CSS START''''''''''''''''%>
.sp_session .topbox {background:url(/backoffice/images/lobbysession/<%=PictureBig%>) no-repeat;height:371px;color:#555555;padding-right: 0;padding-bottom: 0}
<%'''''''''''''重写大会堂主图CSS END ''''''''''''''''%>
-->
</style>
<!--<div style="font-size:14px; font-family:微軟正黑體" > 新春春节期间, 免费生活大会堂暂停至 2011/02/15 ! </div>-->
<div class="sp_session">
    <div class="topbox"></div>
    <form id="webform" name="webform" method="post" action="/program/member/reservation_class/order_full.asp">
      <div class="sssbox">
        <div class="line0">
          <div class="area0 area-w1">
            <div class="day"> <%=SessionDateSecond%><span class="small m-left5"><%=weekSecond%></span></div>
            <div class="time">
              <%' if ( right(arrValidDate(0),len(arrValidDate(0))-5) <> "04/05" ) then  %>
              <input type="radio" name="booking" id="booking3" value="<%=RoomSecond%>" />
              <%' end if %>
              <%=SessionTimeSecond%></div>
            <h1> <%=SessionNameSecond%></h1>
            <h2><%'=SessionTypeFirst%></h2>
          <%=SessionDescriptionSecond%>

          </div>
          <div class="area0 area-w2">
            <div class="day"> <%=SessionDateFirst%><span class="small m-left5"><%=weekFirst%></span></div>
            <div class="time">
              <input type="radio" name="booking" id="booking2" value="<%=RoomFirst%>" />
              <%=SessionTimeFirst%></div>
            <h1> <%=SessionNameFirst%></h1>
            <h2><%=SessionTypeFirst%></h2>
            <%=SessionDescriptionFirst%> 
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
                <td class="STYLE1">上课前请先进行<span class="STYLE2"><a href="http://www.tutormeet.com/mictest/mictest.html?lang=3&compstatus=vipabc" target="_blank">平台测试&gt;&gt;&gt;点我测试</a></span></td>
              </tr>
            </table>
          </div>
        </div>
        <!--left-->
<div class="line02"> </div>
        <!--right--><!--right-->
        <div class="clear"> </div>
        <div class="ssbtn"> <a href="javascript:$('#webform').submit()"> <img src="/images/images_cn/ss_btn2.gif" /></a></div>
      </div>
  </form>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font></div>
<!--內容end-->
