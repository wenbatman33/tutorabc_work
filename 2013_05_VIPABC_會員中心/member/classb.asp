<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<!--#include file="reservation_class/include/include_prepare_and_check_data.inc"-->
<%
'取得系統時間
Dim sys_now_datetime : sys_now_datetime = getFormatDateTime(now(), 8)

Dim str_sql
Dim str_attend_sestime : str_attend_sestime = "" 
Dim int_i 
Dim str_today_session : str_today_session = "" '今日下堂課時間
Dim str_other_session : str_other_session = "" '其他咨訊時段
Dim arr_other_session : arr_other_session = "" '其他咨訊時段陣列
	
Dim str_show_today	'顯示今日上課狀態
Dim int_dcgs_time	'現在的時間+3.5小時
Dim int_now_time	'現在的時間
Dim int_session_time '現在的時間+50分鐘
Dim int_session		'實際上課時間
Dim bol_session_date : bol_session_date = 0 'false (表沒課，不會開啟教材、顧問相關資料)
Dim str_session_sn	'session_sn
Dim str_material		'教材編號
Dim str_special_sn : str_special_sn = ""				'大會堂sn
Dim str_attend_consultant  : str_attend_consultant = ""	'顧問con_sn
Dim str_get_session_sn : str_get_session_sn = getSessionRoomTime()	  '欲找的session_sn
Dim str_session_date : str_session_date = ""				'將session_sn轉為日期格式
'table
Dim var_arr
Dim arr_result

'是否連續訂課
Dim isContinuousAttend : isContinuousAttend = false
	
'教材資料
Dim str_material_rating : str_material_rating = "" '教材平均分數
Dim str_material_ltitle : str_material_ltitle = ""	'教材名稱
Dim str_material_ld	: str_material_ld = ""	'教材簡介
Dim str_attend_level	: str_attend_level = ""	'等級
Dim str_foldername : str_foldername ="" '教材圖檔
Dim obj_fs	'讀檔物件
			 
'先判斷現在時間(最後一堂課，會有跨日問題)
Dim str_search_date	'DB撈的日期
Dim str_search_time	'DB撈的時間
Dim str_settime 		'設定是否需要-1 

dim intContractSn : intContractSn = getSession("contract_sn", CONST_DECODE_NO)

dim arrParam : arrParam = null '傳入sql的array
dim strSql : strSql = "" 'sql
dim arrSqlResult : arrSqlResult = "" 'sql結果

'------------------ 判斷合約未確認需導至確認頁 ------------- start -----------
'Session("contract_sn")-->有合約但尚未確認合約才會有值
'20101119 阿捨修正 續約客戶邏輯 當有未點擊同意合約時 且 合約服務日已到 才需導至合約確認頁
if ( not isEmptyOrNull(intContractSn)) then  
	arrParam = Array(g_var_client_sn, intContractSn)
    '修正續約客戶
	strSql = " SELECT contract_sn "
	strSql = strSql & " FROM client_purchase(nolock) "
	strSql = strSql & " WHERE (client_sn = @client_sn) AND (contract_sn = @contract_sn) AND (Valid = 1)  "
	strSql = strSql & " AND datediff(day,product_sdate,GETDATE())>=0 AND datediff(day, GETDATE(),product_edate)>=0 "
	strSql = strSql & " ORDER BY contract_sn "
	arrSqlResult = excuteSqlStatementRead(strSql, arrParam, CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arrSqlResult)) then
		if (Ubound(arrSqlResult) >= 0) then
			'確認合約
			Call alertGo("", "contract_audit.asp", CONST_NOT_ALERT_THEN_REDIRECT)
			response.end
		end if
	end if
end if
'------------------ 判斷合約未確認需導至確認頁 ------------- end -------------

'20121128 Freeman 修改單字預習課前一小時出現 START
bolVIPABCJR = session("VIPABCJR") '是否為VIPABCJR客戶
'檢查是否為VIPABCJR客戶
if ( isEmptyOrNull(bolVIPABCJR) ) then
   bolVIPABCJR = isVIPABCJRClient(g_var_client_sn, CONST_VIPABC_RW_CONN)
end if

    if ( true = bolVIPABCJR ) then
	    arr_result = getClientAttendListJR(g_var_client_sn,str_get_session_sn)
    else
        arr_result = getClientAttendList(g_var_client_sn,str_get_session_sn)
    end if

	if (isSelectQuerySuccess(arr_result)) then		   
		if (Ubound(arr_result) >= 0) then
			for int_i = 0  to Ubound(arr_result,2) '只捉二筆
				dim realSessionSN : realSessionSN = replace(getFormatDateTime(arr_result(0,int_i),5),"/","")
				if ( int_i = 0 ) then
					if( left(str_get_session_sn,8) = realSessionSN ) then
						dim str_firstclass : str_firstclass = getFormatDateTime(getFormatDateTime(arr_result(0,int_i),5) & " " &arr_result(1,int_i)&":30",3)
						'找出今日下一堂課程時間 2009120417
						str_attend_date = arr_result(0,int_i)
						str_today_session = arr_result(1,int_i)
						str_session_sn = arr_result(2,int_i)
						if isnull(str_session_sn) then
							str_session_sn = realSessionSN & right("0" & str_today_session,2)
						end if
						str_material = arr_result(3,int_i)
						str_special_sn = arr_result(4,int_i)
						str_attend_consultant = arr_result(5,int_i)	
					else
						str_other_session = getFormatDateTime(arr_result(0,int_i),5) & arr_result(1,int_i)
					end if
				else
					if( str_today_session <> "" and int_i = 1) then
						
						dim str_secondclass : str_secondclass = getFormatDateTime(getFormatDateTime(arr_result(0,int_i),5) & " " &arr_result(1,int_i)&":30",3)
						if(DateDiff("N", str_firstclass, str_secondclass) <= 60 AND ( DateDiff("N", str_firstclass, Now()) >= CONST_DEFAULT_PREVIEW_WORD_BRFORE_CLASS_WHEN_CONTINUOUS_ATTEND_MINUTE )) then
							str_session_sn = arr_result(2,int_i)
							if isnull(str_session_sn) then
								str_session_sn = left(arr_result(2,int_i),8) & right("0" & str_today_session,2)
							end if
							str_material = arr_result(3,int_i)
							str_special_sn = arr_result(4,int_i)
							str_attend_consultant = arr_result(5,int_i)
							isContinuousAttend = true
						else
							str_other_session = getFormatDateTime(arr_result(0,int_i),5) & arr_result(1,int_i)
						end if	
					else
						str_other_session = getFormatDateTime(arr_result(0,int_i),5) & arr_result(1,int_i)
					end if			
				end if
			next
		else
			response.write getWord("class_1") 
		end if
	end if 
	
    if ( isEmptyOrNull(str_session_sn) AND hour(now()) = "0" ) then
        str_get_session_sn = getFormatDateTime(date(),7) & right("0" & hour(now()),2)
        if ( true = bolVIPABCJR ) then
	        arr_result = getClientAttendListJR(g_var_client_sn,str_get_session_sn)
        else
            arr_result = getClientAttendList(g_var_client_sn,str_get_session_sn)
        end if
        if (isSelectQuerySuccess(arr_result)) then
            if (Ubound(arr_result) >= 0) then
	            for int_i = 0  to Ubound(arr_result,2) '只捉二筆
		            if ( int_i = 0 and left(str_get_session_sn,8) = left(arr_result(2,int_i),8) ) then  '捉第一筆
			            '找出今日下一堂課程時間 2009120417
			            str_today_session = arr_result(1,int_i)
			            str_session_sn = arr_result(2,int_i)
			            str_material = arr_result(3,int_i)
			            str_special_sn = arr_result(4,int_i)
			            str_attend_consultant = arr_result(5,int_i)
		            else
			            '找出其他咨訊時間
			            if (int_i = 1) then 
				            str_other_session = getFormatDateTime(arr_result(0,int_i),5) & arr_result(1,int_i)
			            end if 
		            end if 
	            next
            end if
        end if
    end if
'20121128 Freeman 修改單字預習課前一小時出現 END

%>
<%'20110314 tree RD11031101 JR和VIPABC加入預覽教材功能---開始---%>
<script src="/lib/javascript/AC_OETags.js" language="javascript"></script><!--張京-->
<%
Dim strBrowserInfo : strBrowserInfo = Request.ServerVariables("HTTP_USER_AGENT") 
if ( InStr(strBrowserInfo,"MSIE 6") > 0 ) then 
%>
<link type="text/css" href="/lib/javascript/JQuery/css/ui-lightness_white/jquery-ui-1.8.14.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.14.custom.min.js"></script>
<% else %>
<link type="text/css" href="/lib/javascript/JQuery/css/ui-lightness_white/jquery-ui-1.8.7.custom.css" rel="stylesheet" />
<script type="text/javascript" src="/lib/javascript/jquery-ui-1.8.2.custom.min.js"></script>
<% end if %>
<!--<script type="text/javascript" src="/lib/javascript/jquery-1.4.4.min.js"></script>-->

<link href="/css/previewCourses.css" rel="stylesheet" type="text/css" /><!--TREE++-->
<!--[if !IE]><!-->	<!--如果不是IE要修改邊界長度-->
	<!--<style>
		.pop-up .pop-body {  width:1025px; height:700px; background:#fff; text-align:left; font-weight:bold; padding:30px 0 10px 25px; margin-top:-5px;}
	</style>-->
<!--<![endif]-->
<script language="JavaScript" type="text/javascript">
	var intDialogSwitch = 0;//偵測視窗窕整過大小用來判斷現在使用著是開啟燈箱還是關閉燈箱
//20101203 tree 預覽教材--燈箱開始--
	//燈箱開啟
	function dialog_open()
	{	
	//	intDialogSwitch = 1;
		$("#dialog_mtl_review").dialog("open");
        $("#sel_nation_code").hide();
        $("#sel_order_time").hide();
        $("#rad_know_abc").hide();
	}
	//燈箱關閉
	function closeb()
	{
	//	intDialogSwitch = 0;
		$("#dialog_mtl_review").dialog("close");
        $("#sel_nation_code").show();
        $("#sel_order_time").show();
        $("#rad_know_abc").show();
	}
	$(function()
	{		
		$("#dialog_mtl_review").dialog("destroy");
		$("#dialog_mtl_review").dialog({
			autoOpen: false,
			bgiframe : true,
			show: "blind",
			draggable: false,
			resizable: true,
			width: 1100,
		//	height: 600,		
			modal: true,
			overlay :{
				background: '#CCCCCC',
				opacity: 1.1
			}
		});
		//================燈箱 start ========================
		//$("#dialog_mtl_review").dialog("open");
		//***********將 原燈箱UI的 關閉鈕 拿掉 start********
		$("a").removeClass("ui-dialog-titlebar-close ui-corner-all ui-state-focus");
		$("span:contains('close')").removeClass().empty();
		//***********將 原燈箱UI的 關閉鈕 拿掉 end********
		//$("#dialog_mtl_review").dialog("option", "resizable", false);
		//$("#dialog_mtl_review").dialog("option", "draggable", false);
		//================燈箱 end ========================
	});
//20101203 tree 預覽教材--燈箱結束--
//20101203 tree 預覽教材--燈箱結束--
//偵測視窗窕整過大小

	var link　= null;
	var block = null;
	var resizeWindow = function () {
		if ( intDialogSwitch == 1 )//當燈箱開啟時調整視窗大小要關閉燈箱在開啟
		{
			//燈箱關閉
			$("#dialog_mtl_review").dialog("close");
			//在開啟燈箱
			$("#dialog_mtl_review").dialog("open");
		}
	};
	window.onload = function () {
		window.onresize = resizeWindow;
	};
//偵測視窗窕整過大小

//'20110314 tree RD11031101 JR和VIPABC加入預覽教材功能---結束---


//系統時間
var str_now_sys_datetime = '<%=sys_now_datetime%>';  
//20101206 阿捨修正 js月份從0開始 需減1
var str_now_datetime = new Date(str_now_sys_datetime.substr(0,4), parseInt(str_now_sys_datetime.substr(4,2))-1, str_now_sys_datetime.substr(6,2), str_now_sys_datetime.substr(8,2), str_now_sys_datetime.substr(10,2), str_now_sys_datetime.substr(12,2));	

var str_session = '<%=str_today_session%>'; //今日上課時間
var str_specianl = '<%=str_special_sn%>';	//大會堂
	
//現在時間
var str_now_time ;
//載入教材頁面

function loadConBbook()
{
	//改變按鈕css
	$("#mune_class_info a").removeClass(); 
	$("#mune_class_info").addClass("mune_01"); 
	$("#mune_consultant").removeClass(); 
	$("#mune_consultant a").addClass("mune_01"); 
	//載入ajax ajax_con_book.asp 將頁面載入
	$("#div_consultant_or_class_info").load("ajax_con_book.asp", {input_session_sn : "<%=str_session_sn%>", input_material : "<%=str_material%>", input_today_session : str_session}  );
}

//載入consultant頁面
function loadConsultant()
{
	//改變按鈕css
	$("#mune_class_info a").addClass("mune_01"); 
	$("#mune_class_info").removeClass(); 
	$("#mune_consultant").addClass("mune_01"); 
	$("#mune_consultant a").removeClass(); 
	//載入ajax ajax_merge_consultant_info.asp 將頁面載入
	$("#div_consultant_or_class_info").load("/program/consultant/include/ajax_merge_consultant_info.asp", {ndate : new Date(), con_sn : "<%=str_attend_consultant%>"});
}

$(document).ready(function()
{     
 	clockTime();
	//一開始預設下方頁面為教材資訊 載入ajax_con_book.asp
	$("#div_consultant_or_class_info").load("ajax_con_book.asp", {input_session_sn : "<%=str_session_sn%>", input_material : "<%=str_material%>", input_today_session : str_session}  );        
});

//時鐘
function clockTime()
{
    str_now_datetime.setSeconds(str_now_datetime.getSeconds() + 1);
    var str_now_year = str_now_datetime.getFullYear().toString();
    var str_now_month = padLeft((str_now_datetime.getMonth()+1).toString(),2,"0");
    var str_now_date = padLeft(str_now_datetime.getDate().toString(),2,"0");
	var str_now_hours = padLeft(str_now_datetime.getHours().toString(),2,"0");
	var str_now_minutest = padLeft(str_now_datetime.getMinutes().toString(),2,"0");
	var str_now_seconds = padLeft(str_now_datetime.getSeconds().toString(),2,"0");

	var int_start_room; //開始進教室時間
	var int_end_room;	//結束進教室時間

	$("#span_clock").text(str_now_hours + ":" + str_now_minutest + ":" + str_now_seconds);
	<% '20101206 阿捨修正 跨天判斷問題 直接加入日期 及 將所有時間點均設定成常數 %>
	$("#span_in_room").hide();
	if (str_session != "") 
    {
		str_now_time = str_now_year+''+str_now_month+''+str_now_date+''+str_now_hours +''+ str_now_minutest +''+ str_now_seconds;

		//大會堂20進教室，其他為26分以後才能進教室
		if ( str_specianl != "" ) 
        {
            int_room_min = '<%=CONST_DEFAULT_FREELOBBY_COUNTDOWN_CLASS_MINUTE%>';
            int_end_min = '<%=CONST_DEFAULT_FREELOBBY_GOIN_CLASS_MINUTE%>';
		}
        else
        {
            int_room_min = '<%=CONST_DEFAULT_COUNTDOWN_CLASS_MINUTE%>';
            int_end_min = '<%=CONST_DEFAULT_GOIN_CLASS_MINUTE%>';
		}
		
        //跨天需減1天判斷
        if ( parseInt(str_now_minutest) <= 15 && parseInt(str_now_hours) == 0 )
        {
            str_now_date = padLeft(parseInt(str_now_date)-1, 2, "0");
        } 
      
        int_start_room = str_now_year+''+str_now_month+''+str_now_date+''+str_session + "" + int_room_min + "00";
        int_end_room = str_now_year+''+str_now_month+''+str_now_date+''+str_session + "" + int_end_min + "00";
        
        if ( parseInt(str_now_time) >= parseInt(int_start_room) && parseInt(str_now_time) <= parseInt(int_end_room) ) 
        {
			//計算開始
			//計算秒(剩餘秒數)數ex：26分進來，到數為60秒
			var int_remain_sec = 60 - str_now_seconds;
			var int_remain_min = int_end_min - str_now_minutest  ;
			
			var int_remain ;
			if ( int_remain_sec > 0 ) {
				int_remain_min = int_remain_min - 1 ; 
			}
						
			int_remain = ( int_remain_min * 60 ) + int_remain_sec;
			setTimeout("TimeView("+int_remain+")",1000);
		} 
        else if ( parseInt(str_now_time) > parseInt(int_end_room) ) //27~30分(上課期間)	  
        {   
			$("#span_in_room_time").hide();
			$("#span_in_room").show();			
		}	
	}
}

timerID = setInterval("clockTime()",1000);

//倒數計時器
function TimeView(sec)
{
	var showTime = Math.floor(sec / 60) ; //分
	
	var dSecs = sec % 60 ; 
	if (dSecs < 10) { //補0
		dSecs = "0"+dSecs
	}
	
	$("#span_in_room_time").text("<%=getWord("CLASS_2")%>"+showTime+":"+dSecs+"<%=getWord("CLASS_3")%>");
}   

//若str的長度小於length，則在其左邊補sign   
function padLeft(str,length,sign)
{  
    if(str.length>=length)
    {  
        return str;
    }   
    else
    {    
        return padLeft(sign+str,length,sign);
    }
}  	 
</script>

		
       <!--內容start-->
       <div id="class_membox">
       
         <!--收藏課程start-->
         <div class="main_mylist">
           <ul>
            <%
			
		if ( str_today_session <> "") then 
			'表示今日有課
			'g_str_show_today = "今日下一堂课程时间："
			'--------------- 課程為3分鐘內，才開放"進入教室"按鈕 -------------- start -----------
			if ( true=isContinuousAttend ) then 
				str_session_date = str_secondclass
			else
				str_session_date = str_firstclass
			end if
			
			'有教材和顧問，才會顯示下方的資訊
            '20101123 阿捨新增60分內 可單字預習
			if ( str_material <> "" and str_attend_consultant <> "" and DateDiff("n",now(),str_session_date) <= CONST_DEFAULT_PREVIEW_WORD_BRFORE_CLASS_MINUTE ) then 
				bol_session_date = 1 '有課程資料，可開啟下方的課程相關資訊
			end if 
			
			if ( true = isContinuousAttend ) then
				'如果連續訂課，頁面課程資訊為第一堂課程
				str_today_session = getFormatDateTime(str_session_date,6)&"&nbsp;&nbsp;/&nbsp;"&right("0"&str_today_session,2)&":30"
				str_show_today = getWord("now_class")&"："&"<span class='red'>"&str_today_session&"</span> "
			elseif ( DateDiff("n",str_session_date,now()) <= CONST_DEFAULT_ONE_CLASS_MINUTE AND DateDiff("n",str_session_date,now()) >= 0 ) then 
				'上課中 (現在課程)
				str_today_session = getFormatDateTime(str_session_date,6)&"&nbsp;&nbsp;/&nbsp;"&hour(str_session_date)&":30"
				str_show_today = getWord("now_class")&"："&"<span class='red'>"&str_today_session&"</span> "	
			else'if ( int_dcgs_time > int_session ) then 
				'上課前  超過3.5小時
				str_other_session = left(str_today_session,2)&":30"
				str_show_today = "，"&getWord("class_6")&"<span class='red'>"&str_other_session&"</span>"&getWord("class_7")
				
			end if 
			'--------------- 課程為3.5小時內，才開放"進入教室"按鈕 -------------- end -----------
		else
			'今日無課程
			str_show_today = "，"&getWord("class_8")
			'退費中且無堂數，無法預訂課程 0：無 / 1：退費中
			if ( getIsClientMoneyBack(g_var_client_sn) = 1 ) then 
				str_show_today = str_show_today &"<a href='reservation_class/normal_and_vip.asp'>"&getWord("class_9")&"</a>。"
			else
				str_show_today = str_show_today &getWord("class_9")&"。"
			end if 
			
		end if 
		%>
             <li><%=getWord("class_10")%><span id="span_clock" class="red"></span><%=str_show_today%></li>
             
			 <% if (bol_session_date = 1) then %>
             
             	<li>
                <% 
				'if ( int_dcgs_time > int_session  ) then 
				if ( DateDiff("N",now(),str_session_date) <= 210  ) then 
				
				 'dcgs排課完成後
				%>
					<%=getWord("class_11")%><a href="class_vocabulary.asp?attend=yes&mtl=<%=str_material%>"><%=getWord("class_12")%></a>及教材预览，<%=getWord("class_13")%>
                <% end if %>
                <span id="span_in_room_time" class="right bold color_1" ></span>
                <span class="right bold color_1" id="span_settime"></span>
                    <span id="span_in_room" class="right">
                        <input type="submit" name="button2" id="button2"  value="+ <%=getWord("class_14")%>" class="btn_1 m-left5" onclick="javascript:location.href='class_notice.asp';"/>
                    </span>
                     
                     </span>
                </li>
                
            <% end if %>
            
            
          
           </ul>
         </div>
         <!--標題end-->
         <% if (bol_session_date = 1) then %>
             <!--選單-->
                <div class="main_memtitle_line">
                <div id="mune_class_info" class="mune_01"><a id="" href="javascript:void(0);" onclick="javascript:loadConBbook();"><%=getWord("CLASS_15")%></a></div>
                <div id="mune_consultant" ><a href="javascript:void(0);" class="mune_01" onclick="javascript:loadConsultant();"><%=getWord("CLASS_16")%></a></div>
                </div>
                <div class="clear"></div>
             <!--選單end-->
             <!--書本資訊標頭-->
              <%
			 '========================== 教材資訊 ========================= start =====================
			 %>
             <div id="div_consultant_or_class_info" class="con_book">
			 <!--ajax_con_book.asp 載入-->
             </div> 
             <div class="clear"></div>
             <%
			 '========================== 教材資訊 ========================= end =====================
			 %>
              <%'暫時關閉
			  if ( 1 = 1) then %>
             <!--書本資訊標頭end-->
			 <!---ryanlin hide 20100113 start-->
			<%	if (1 = 2) then %>
             <!--個人資料標頭-->
             <div class="con_title">
               <div class="con_header_left"><img src="/images/images_cn/arrow_path.gif" alt="" width="4" height="12" hspace="2" /><%=getWord("CLASS_18")%></div>
               <div class="con_header_right"></div>
             </div>
             <div class="clear"></div>
             <!--   個人資料標頭end    -->
             <!--    選      單  -->
                <div class="main_memtitle">
                <div class="mune_02"><%=getWord("CLASS_19")%></div>
                <div ><a href="#" class="mune_02"><%=getWord("CLASS_20")%></a></div>
                <div ><a href="#" class="mune_02"><%=getWord("CLASS_21")%></a></div>
                <div ><a href="#" class="mune_02"><%=getWord("CLASS_22")%></a></div>
                <div ><a href="#" class="mune_02"><%=getWord("CLASS_23")%></a></div>
                </div>
                <div class="clear"></div>
             <!--選單end-->
            
             <!--    書本列表頁      -->
             
             <div class="book_list">
             <div class="left"><img src="/images/images_cn/book_iocn_02.jpg" alt="" /></div>
             <div class="right">
             <div class="left_1">
               <a href="#">The future of touch screen</a><br />
               The student should learn vocabulary, phrases, andgrammar related to...<br />
               <a href="#">travel,</a> <a href="#">USA,</a> <a href="#">NY,</a>
             </div>
             <div class="right"><span class="color_7"><%=getWord("CLASS_22")%>：</span>Carolina Espiro<br />
               <span class="color_7"><%=getWord("CLASS_23")%>：</span><img src="/images/images_cn/love.jpg" alt="" width="13" height="10" /><span class="color_1"> 31</span><br />
               <span class="color_7"><%=getWord("CLASS_21")%>：</span><img src="/images/images_cn/stars.jpg" alt="" width="13" height="13" /><img src="/images/images_cn/stars.jpg" alt="" width="13" height="13" /><img src="/images/images_cn/stars.jpg" alt="" width="13" height="13" /><img src="/images/images_cn/stars_02.jpg" alt="" width="13" height="13" /><img src="/images/images_cn/stars_03.jpg" alt="" width="13" height="13" /><br />
               <span class="color_7"><%=getWord("CLASS_24")%>：</span><span class="color_1">5</span><br />
               <span class="color_7"><%=getWord("CLASS_25")%></span><span class="color_1"><span class="color_7">：</span>0.5</span></div>
             </div>
             <br class="clear" />
             </div>
             
             <!--   書本列表頁    end -->
             <!--   跳頁         -->
			 <%	end if %>
			 <!---ryanlin hide 20100113 end-->
             <% end if %>
         <% end if %>
         <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
         
       </div>
       
       <!-- 內容end -->

<%'20110314 tree RD11031101 JR和VIPABC加入預覽教材功能---開始---
%>
<!--燈箱--開始---->
<div id="Default_dialog_off" style="z-index:300;">
	<script type="text/javascript">
		document.getElementById("Default_dialog_off").style.display= "none"; 	<%'關閉自己%>
	</script>
	<!--燈箱上半部分--開始---->
	<div id="dialog_mtl_review" class="pop-up" style="z-index:300;">
		<div class="pop-top-right" style="z-index:300;">
			<!--影像地圖的連結-->
			<img src="/images/previewcourses/pop_top.png"  width="1050" height="28" border="0" usemap="#Map" />
			<map name="Map" id="Map">
				<area shape="rect" onclick="closeb(); return false;" coords="1017,7,1043,28" href="javascript:void(0);" />
			</map>
		</div>
	<br class="clear" />
		<div class="pop-body" style="z-index:300;"> 
	<!--燈箱上半部分--結束---->
			<% '20101203 TREE 當"課程預覽按鈕"按下時執行，張金的.swf程式----開始-- %>
			<% '20101203 TREE 建立視窗大小 名稱為 dialog_mtl_review %>
			<%
				Dim TUTORMEET_WEB_HOST : TUTORMEET_WEB_HOST = "www.tutormeet.com"
			%>
			<div id="content" style="z-index:300;">
				<div id="primary" style="z-index:300;">
					<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444560040000"
						id="MaterialViewer" width=dialog_mtl_review_width_2 height=dialog_mtl_review_height_2
						codebase="" style="z-index:301">
						<param name="movie" value="/flash/MaterialViewer.swf" />
						<param name="quality" value="high" />
						<param name="bgcolor" value="#ffffff" />
						<param name="flashVars" value="host=<%=TUTORMEET_WEB_HOST%>&materialId=<%=str_material%>" />
						<param name="allowScriptAccess" value="sameDomain" />
						
						<embed src="/flash/MaterialViewer.swf" quality="high" bgcolor="#ffffff"
							width="1000" height="700" name="MaterialViewer" align="middle"
							play="true"
							loop="false"
							quality="high"
							flashVars="host=<%=TUTORMEET_WEB_HOST%>&materialId=<%=str_material%>"
							allowScriptAccess="sameDomain"
							type="application/x-shockwave-flash"
							pluginspage="http://www.adobe.com/go/getflashplayer"
							wmode="opaque">
						</embed>
					</object></noscript>
					<p>Copyright &#169; <%=Year( Date() )%> TutorGroup. All Rights Reserved.</p>
				</div>
			</div>
			<% '-- 20101203 TREE 當"課程預覽按鈕"按下時執行，張金的.swf程式----結束-- %>
	<!--燈箱下半部分--開始---->
		</div>
		<div class="pop-under-right" style="z-index:300;">
			<img src="/images/previewcourses/pop_under.png" width="1050">
		</div>
		<br class="clear" />
	</div>
	<!--燈箱下半部分--結束---->
</div>
<!--燈箱--結束---->
<%'20110314 tree RD11031101 JR和VIPABC加入預覽教材功能---結束---
%>
