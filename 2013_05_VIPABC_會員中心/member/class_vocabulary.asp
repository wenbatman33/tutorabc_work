<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<script LANGUAGE="JavaScript">
function OpenVocWin2(p_str_voc) 
{
	var con_id = "#voc_sound_"+p_str_voc;
	$.ajax({        
            type:"POST",
            cache:false,
            url:"PlayVoc.asp",
            data:"voc=" + p_str_voc,
		
			success:function(data){
			$(con_id).html(data);
			},
			error:function(){
				alert("Processing Error");
			}
			});
}

function excuteVocOpt(p_str_voc) 
{
	var con_id = "#voc_opt_"+p_str_voc.replace(/ /,"_");
	$.ajax({        
            type:"POST",
            cache:false,
            url:"vocabulary_opt.asp",
            data:"voc_opt_type=" + $(con_id).attr("title") + "&voc=" + encodeURIComponent(p_str_voc) + "&data=<%=g_var_client_sn%>",
		
			success:function(data){
				if (data == "success")
				{
					if ($(con_id).attr("title") === "del")
					{
						$(con_id).text("<%=getWord("class_vocabulary_1")%>");
						$(con_id).attr("title", "add");
					}
					else if ($(con_id).attr("title") === "add")
					{
						$(con_id).text("<%=getWord("class_vocabulary_2")%>");
						$(con_id).attr("title", "del");
					}
				}
				else
				{
					alert(data);
				}
			},
			error:function(){
				alert("Processing Error");
			}
			});
			
}
$(document).ready(function()
{     
 	clockTime();
});
//時鐘
function clockTime()
{
	str_now_datetime = new Date();
	var str_now_hours = padLeft(str_now_datetime.getHours().toString(),2,"0");
	var str_now_minutes = padLeft(str_now_datetime.getMinutes().toString(),2,"0");
	var str_now_seconds = padLeft(str_now_datetime.getSeconds().toString(),2,"0");

	$("#span_clock").text(str_now_hours + ":" + str_now_minutes + ":" + str_now_seconds);
}
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
timerID = setInterval("clockTime()",1000);
</script>
<%
   '取得系統時間
    Dim sys_now_datetime : sys_now_datetime = getFormatDateTime(now(), 8)
	Dim int_return_contract : int_return_contract = 0 '判斷是否要導到 「確認合約頁」
	
	Dim str_sql
	Dim str_attend_sestime : str_attend_sestime = "" 
	Dim int_i 
	Dim str_today_session : str_today_session = "" '今日下堂課時間
	Dim str_other_session : str_other_session = "" '其他咨訊時段
	Dim arr_other_session : arr_other_session = "" '其他咨訊時段陣列
	
	Dim str_show_today	'顯示今日上課狀態
	Dim int_now			'現在的時間+3.5小時
	Dim int_session		'實際上課時間
	Dim bol_session_date : bol_session_date = 0 'false (表沒課，不會開啟教材、顧問相關資料)
	Dim str_session_sn	'session_sn
	Dim str_material : str_material = request("mtl")		'教材編號
	Dim isAttend : isAttend = request("attend")
	Dim str_get_session_sn : str_get_session_sn = getSessionRoomTime()	
	
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
	
'----------- 跨時、跨日問題 ------------- start ------------
'if (cint(hour(now())) = 0 and cint(minute(now())) <= 20) then  '跨日的問題午夜12:20分之前還可進教室
'	str_search_date = getFormatDateTime(DateAdd("d" , -1 , date()),5) 
'	str_search_time = 24
'else
'	str_search_date = getFormatDateTime(date(),5)
'	str_search_time = hour(now())
'end if 
'
'if minute(now)<= 20 then str_settime = -1 else str_settime = 0 end if '跨時 ex 13:30分的課 14:05進入時
'
'str_search_time=str_search_time + str_settime
'----------- 跨時、跨日問題 ------------- end ------------
'response.write str_search_date&"<br>"
%>
<!-- Member Content Start -->
<div class="main">
    <div class="main_con">

       <!--內容start-->

       <div class='page_title_1'><h2 class='page_title_h2'>单词预习</h2></div>
  	   <div class="box_shadow" style="background-position:-80px 0px;"></div>
       <div id="class_membox">
        <!--收藏課程start-->
        <!--收藏課程start-->
<%	
str_material = request("mtl")
if isEmptyOrNull( str_material ) or (isAttend = "yes") then %>
         <div class="main_mylist">
           <ul>
        <%
		'20121128 Freeman 修改單字預習課前一小時出現 START
		
		'進資料庫抓訂課
		'20120319 阿捨新增 VIPABCJR客戶判斷 bolVIPABCJR
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
							if(not isNull(arr_result(3,int_i))) then
								str_material = arr_result(3,int_i)
							end if
							str_special_sn = arr_result(4,int_i)
							str_attend_consultant = arr_result(5,int_i)	
						else
							str_other_session = getFormatDateTime(arr_result(0,int_i),5) & arr_result(1,int_i)
						end if
					else
						if( str_today_session <> "" and int_i = 1) then
							'如果今日有課程，會抓第二堂課，然後判斷是否連續訂課且第一堂課開始十分鐘，為true則把單字預習&教材預覽替換成第二堂課內容
							'如果今日沒課程則跳過所有判斷，直接給str_orther_seseion值
							
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
				if (str_material <> "") then 
					bol_session_date = 1
				end if
			else
				'如果抓不到課頁面最下面會出現錯誤訊息，因此mark掉
				'response.write getWord("class_1") 
			end if
		end if 
	
		if ( str_today_session <> "" AND str_attend_date <> "") then 
		
			int_session = getFormatDateTime(getFormatDateTime(str_attend_date,5) & " " & Right("0"&str_today_session,2) & ":30",3) 
			
			if DateDiff("N",now(),int_session) <= 60 then 
				str_today_session = left(str_today_session,2)&":30"
				str_show_today = "，"&getWord("CLASS_4")&"<span class='red'>"&str_today_session&"</span> "&getWord("CLASS_5")
				bol_session_date = 1 
			else
				str_other_session = left(str_today_session,2)&":30"
				str_show_today = "，"&getWord("CLASS_6")&"<span class='red'>"&str_other_session&"</span>"&getWord("CLASS_7")
				bol_session_date = 0
			end if 	
			
		else
			'今日無課程
			str_show_today = "，"&getWord("CLASS_8")&"<a href='reservation_class/normal_and_vip.asp'>"&getWord("CLASS_9")&"</a>。"
		end if 
		'20121128 Freeman 修改單字預習課前一小時出現 END
		%>
             <li><%=getWord("CLASS_10")%><span id="span_clock" class="red"></span><%=str_show_today%></li>
             <% if (bol_session_date = 1) then%>
             	<li><%=getWord("CLASS_11")%><a href="class_vocabulary.asp?attend=yes"><%=getWord("CLASS_12")%></a>，<%=getWord("CLASS_13")%>
                <span class="right bold color_1" id="span_settime"></span>
                <span class="right"><!--<input type="submit" name="button2" id="button2" value="+ <%=getWord("CLASS_14")%>" class="btn_1 m-left5" onclick="javascript:location.href='class_notice.asp';"/>-->
             	</span></li>
            <% end if %>
           </ul>
         </div>
<%
	else
		bol_session_date = 1
	end if
%>
<% if (bol_session_date = 1) then%>
         <!--標題end-->
         <!--  個人資料標頭    -->
         <div class="con_title">
           <div class="con_header_left"><img src="/images/images_cn/arrow_path.gif" alt="" width="4" height="12" hspace="2" /><%=getWord("CLASS_12")%></div>
		   <% if isEmptyOrNull( str_material ) then %>
           <div class="con_header_right "><span class="t-right"><input type="submit" name="button2" id="button2" value="+ <%=getWord("class_vocabulary_5")%>" class="btn_1 m-top5"onclick="javascript:location.href='class.asp';"/>
           </span></div>
		   <% end if %>
         </div>
         <div class="clear"></div>
         <!--   個人資料標頭end    -->
         <!--  單字查詢   -->
		 <%
			Dim obj_mtl_voc_opt
			Dim arr_voc_front_name_result
			Dim arr_voc_result
			Dim int_col
			Dim str_voc_front_name
				str_voc_front_name = request("voc_front_name")
			if(str_voc_front_name  = "") then
				str_voc_front_name = "*"
			end if
			Set obj_mtl_voc_opt = Server.CreateObject("TutorLib.MaterialOpt.MaterialVocabularyOpt")
		 %>
         <div class="con_word">
		  <ul>
		 <%
		if(not isEmptyOrNull(str_material)) then
				arr_voc_front_name_result = obj_mtl_voc_opt.getMaterialFrontChar(str_material)
			if (obj_mtl_voc_opt.PROCESS_STATE_ID = 1) then 
				Response.write("<li><a href='class_vocabulary.asp?type=*&language=zh-cn&mtl="&str_material&"' class=""ch"">"&getWord("class_vocabulary_6")&"</a></li>")		
				for int_col = 0 to Ubound(arr_voc_front_name_result)
					Response.write("<li><a href='class_vocabulary.asp?voc_front_name="&arr_voc_front_name_result(int_col)&"&language=zh-cn&mtl="&str_material&"' class=""eng"">"&arr_voc_front_name_result(int_col)&"</a></li>")				
				next
			else
			end if
		end if
		 %>
           </ul>  
         
             
         </div>
         <div class="clear"></div>
         <!--  單字查詢  end -->
         <!--  單字內容  -->
         <div class="con_word_list">
            <ul>
              <li><%=getWord("class_vocabulary_7")%></li>
              <li class="li_2"><%=getWord("class_vocabulary_8")%></li>
              <li class="li_3"><%=getWord("class_vocabulary_9")%></li>
            </ul> 
         </div>
		
		 <%
		 Dim str_voc
		 Dim int_total_voc : int_total_voc =0
		 Dim str_voc_opt : str_voc_opt = "add"
         Dim strClass : strClass = "con_word_list"
         dim strSubClass : strSubClass = "color_8"
		 if(not isEmptyOrNull(str_material)) then
			 arr_voc_result = obj_mtl_voc_opt.getMaterialVocabulary(str_material, str_voc_front_name)
			 int_total_voc = Ubound(arr_voc_result)+1
			 for int_col = 0 to Ubound(arr_voc_result)
			 str_voc = arr_voc_result(int_col)
			 str_voc_opt = "add"
				 if ((int_col mod 2)= 0) then
					strClass = "con_word_list_2"
				 else
					strClass = "con_word_list"
				 end if

                 if ( true = bolVIPABCJR ) then
	                strSubClass = "C4_BgGreen"
                else
                    strSubClass = "color_8"
                end if
         %>
		<div class="<%=strClass%>">
			<ul >
                 <li><p class="<%=strSubClass%>"><%=lcase(str_voc)%></p>
				 <%
				
				Dim str_voc_des : str_voc_des = getWord("CLASS_VOCABULARY_1")
				if (obj_mtl_voc_opt.checkIsInVocabularyBank(g_var_client_sn, str_voc)) then
					str_voc_opt = "del"
					str_voc_des = getWord("CLASS_VOCABULARY_2")
				end if
				 %>
                 <a href="javascript:void(0)" onclick="window.open('http://dict.baidu.com/s?wd=<%=lcase(str_voc)%>&tn=dict')"><img src="../../images/vip_voc_01.gif" border=0 alt="" /></a>
                 <br/><br/>
				 <a id ="voc_opt_<%=Replace(str_voc, " ", "_")%>" href="javascript:void(0);" onclick="excuteVocOpt('<%=str_voc%>');" title="<%=str_voc_opt%>"><%=str_voc_des%></a>
				 </li>
                 <li class="li_2"><a  id ="voc_sound_<%=Replace(str_voc, " ", "_")%>" href="javascript:void(0);" onclick="OpenVocWin2('<%=Replace(str_voc, " ", "_")%>');" alt="<%=Replace(str_voc, " ", "_")%>" ><img src="/images/images_cn/arrow_icon.jpg" /></a></li>
                 <li class="li_3"><%
                    dim htmlDescription
                    htmlDescription = obj_mtl_voc_opt.getVocabularyHTMLDescription(Replace(str_voc, " ", "_"), "zh-cn")
                    if right(htmlDescription,37) = "<span class=""bold"">英文释义 :</span><br/>" then
                        dim sqlParameters : sqlParameters = Array(str_material, str_voc)
                        dim sqlStatement : sqlStatement = "SELECT [definition] FROM dbo.material_vocabulary WHERE course=@m AND vocabulary=@v"
                        dim queryResult : queryResult = excuteSqlStatementScalar(sqlStatement, sqlParameters, CONST_VIPABC_RW_CONN)
                        if (isSelectQuerySuccess(queryResult)) then
	                        if (Ubound(queryResult) >= 0) then
		                        htmlDescription = htmlDescription & queryResult(0)
	                        end if
                        end if
                    end if
                    response.Write htmlDescription%></li>
            </ul> 
		</div>
			<%next%>
		<%end if%>
         <!--  單字內容  -->
         <!--  跳頁功能  -->
         <div class="go_page">
         <span class="right_page"><span class="right_text_left"><%=getWord("CLASS_VOCABULARY_10")%> <%=int_total_voc%><%=getWord("CLASS_VOCABULARY_11")%></span> </span>
          <div class="clear"></div></div>
         
         <!--  跳頁功能end  -->
       </div>
       <!-- 內容end -->
    </div>
 <%end if%>
<%
	Set obj_mtl_voc_opt = Nothing
%>
<div class="clear"></div>
<font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>

<!-- Main Content End -->
       
