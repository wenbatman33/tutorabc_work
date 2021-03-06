<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<div id="temp_contant">
    <!--內容start-->
    <div class="main_membox">
    	<div class='page_title_3'><h2 class='page_title_h2'>程度分析</h2></div>
		<div class="box_shadow" style="background-position:-80px 0px;"></div> 
        <!--上版位課程start-->
        <% 
		'如果是從新手上路過來的 就不秀tab的內容 Johnny 2012-08-22
		if not "y" = getRequest("noTab", CONST_DECODE_NO) then %>
        <!--#include virtual="/lib/include/html/include_learning.asp"-->
        <% end if %>
        <!--上版位課程end-->
        <!--程度分析-->
        <%
		redim stre1(5)
		redim stre2(5)

		Dim str_sql : str_sql = ""
		Dim i
		Dim evaluation_type : evaluation_type =""
		Dim client_level : client_level =""

		Dim str_client_fname : str_client_fname = Session("client_fname")
		
		Dim c : c= 0
		Dim gp : gp= 0
		Dim lp : lp= 0
		Dim pp : pp= 0
		Dim sp : sp= 0
		Dim vp : vp= 0
		Dim rp : rp= 0
		Dim placementt : placementt = 0 
		'塞資料的相關設定值
		Dim var_arr 								'傳excuteSqlStatementRead的陣列
		Dim arr_result							'接回來的陣列值
		Dim str_tablecol 							'table欄位
		Dim str_tablevalue 						'table值
		
		'------------- 模擬考部份 -----------------------
		Dim int_act_i
		Dim str_act_inpdate   '輸入時間
		Dim str_act_message3	'結束時間
		Dim str_act_message2	'分數
		Dim str_act_note		'題目
			
		Dim Dinpdate			'模擬考測驗日期
		Dim Hinpdate			'模擬考開始--時
		Dim Minpdate			'模擬考開始--分
		
		Dim HDone				'模擬考結束--時
		Dim MDone				'模擬考結束--分
		Dim Done				'模擬考結束--時:分
		Dim DoneSocre			'模擬考分數
		Dim DDone				'之前沒有紀錄的加45分鐘
		
		if session("productno")=40 then 
			'歌倫比亞客戶
			str_sql = "SELECT sn,listening, speaking, reading, pronunciation, grammar, vocabulary  "
			str_sql = str_sql & " FROM columbia.dbo.client_basic(nolock) "
			str_sql = str_sql & " RIGHT JOIN (SELECT columbia_sn FROM client_basic(nolock) WHERE (Sn = @sn)) DERIVEDTBL"
			str_sql = str_sql & " ON client_basic.sn = DERIVEDTBL.columbia_sn "
				
			var_arr = Array(g_var_client_sn)
			arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
		
			if (isSelectQuerySuccess(arr_result)) then
			
				'--------- 塞適性測驗 ------------ start ------------
				if (Ubound(arr_result) > 0) then 
					'有資料
					
					stre1(0)=arr_result(5,0)	'Grammar
					stre1(1)=arr_result(6,0)  'Vocabulary
					stre1(2)=arr_result(4,0)  'Pronunciation
					stre1(3)=arr_result(2,0)	'Speaking
					stre1(4)=arr_result(3,0)	'Reading
					stre1(5)=arr_result(1,0)	'Listening
					
					placementt=6
					
				else
					'無資料
				end if 
				'--------- 塞適性測驗 ------------ end ------------
			end if    	
		
		else
			'TutorABC客戶
			str_sql = "select lesson_group.client_sn, lesson_group.expr1, isNull(lesson_question.client_level,0) client_level,  "
			str_sql = str_sql & " isNull(lesson_question.evaluation_type,0) evaluation_type  "
			str_sql = str_sql & " from "
			str_sql = str_sql & " 	(select client_sn, min(test_sn) as expr1  "
			str_sql = str_sql & " 	from "
			str_sql = str_sql & " 		(select client_sn, test_sn  "
			str_sql = str_sql & " 		from lesson_question_temp(nolock)  "
			str_sql = str_sql & " 		group by client_sn, test_sn   "
			str_sql = str_sql & " 		having (count(sn) >= 5) and client_sn=@client_sn  "
			str_sql = str_sql & " 	) as lesson_temp  "
			str_sql = str_sql & " 	group by client_sn having (client_sn = @client_sn_1)  "
			str_sql = str_sql & " 	) as lesson_group   "
			str_sql = str_sql & " 	left join lesson_question_temp(nolock) as lesson_question   "
			str_sql = str_sql & " 	on lesson_group.expr1 = lesson_question.test_sn  "
		
			var_arr = Array(g_var_client_sn,g_var_client_sn)
		
			arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
		
		
			if (isSelectQuerySuccess(arr_result)) then
		
				'--------- 塞適性測驗 ------------ start ------------
				if (Ubound(arr_result) > 0) then 
					'有資料
					placementt = Ubound(arr_result,2) + 1
					placementt = 1
					for i = 0 to Ubound(arr_result,2) 
					
						evaluation_type = arr_result(3,i)
						client_level = arr_result(2,i)
						
						if evaluation_type=1 then stre1(0)=client_level	 'Grammar
						if evaluation_type=2 then stre1(1)=client_level  'Vocabulary
						if evaluation_type=3 then stre1(2)=client_level  'Pronunciation
						if evaluation_type=4 then stre1(3)=client_level	 'Speaking
						if evaluation_type=5 then stre1(4)=client_level	 'Reading
						if evaluation_type=6 then stre1(5)=client_level	 'Listening
		
					next
					
				else
					'無資料
				end if 
				'--------- 塞適性測驗 ------------ end ------------
			end if    	
		
		end if 	
		
		'若沒值，表示此客戶並未做過適性測驗
		if stre1(0) = "" then stre1(0) = 0
		if stre1(1) = "" then stre1(1) = 0
		if stre1(2) = "" then stre1(2) = 0
		if stre1(3) = "" then stre1(3) = 0
		if stre1(4) = "" then stre1(4) = 0
		if stre1(5) = "" then stre1(5) = 0	
		
		'---------- 客戶實際上課的教材比例值 (目前學習比重) --------------- start ---------------------
		str_sql = "SELECT client_attend_list.client_sn, client_attend_list.Attend_mtl_1, material.FG, material.FV,   "
		str_sql = str_sql & " material.FP, material.FS, material.FL,  material.FR   "
		str_sql = str_sql & " from "
		str_sql = str_sql & " 	( SELECT client_sn, Attend_mtl_1   "
		str_sql = str_sql & " 		FROM client_attend_list(nolock) "
		str_sql = str_sql & " 		WHERE (client_sn = @client_sn) AND (Attend_mtl_1 IS NOT NULL) AND (Valid = 1) AND (Session_State IN ('1', '2')) "
        str_sql = str_sql & "       UNION     "
        str_sql = str_sql & "     SELECT client_sn, ClassMaterialNumber AS Attend_mtl_1 "
        str_sql = str_sql & " 		FROM member.ClassRecord WITH (NOLOCK) "
        str_sql = str_sql & " 		WHERE (client_sn = @client_sn) AND (ClassMaterialNumber IS NOT NULL) AND (ClassRecordValid = 1) AND (ClassStatusId IN ('2','3','4')) "
		str_sql = str_sql & " 	) AS client_attend_list   "
		str_sql = str_sql & " 	LEFT JOIN material ON client_attend_list.Attend_mtl_1 = material.Course   "
		
		var_arr = Array(g_var_client_sn,g_var_client_sn)
		arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)	
			
		if (isSelectQuerySuccess(arr_result)) then
			
			'--------- 塞適性測驗 ------------ start ------------
			if (Ubound(arr_result) > 0) then 
				'有資料
				'if isnull(rs("fg")) then rs("fg")=0
				c = c + 1
				gp = gp + arr_result(2,0)
				lp = lp + arr_result(6,0)
				pp = pp + arr_result(4,0)
				sp = sp + arr_result(5,0)
				vp = vp + arr_result(3,0)
				rp = rp + arr_result(7,0)
			else
				'無資料
				
			end if 
		end if 	
		
		if(c = 0) then c = 1
		
		stre2(0) = round(gp/c)'Grammar
		stre2(1) = round(vp/c)'Vocabulary
		stre2(2) = round(pp/c)'Pronunciation
		stre2(3) = round(sp/c)'Speaking
		stre2(4) = round(rp/c)'Reading
		stre2(5) = round(lp/c) 'Listening
		'---------- 客戶實際上課的教材比例值 (目前學習比重) --------------- end ------------------

		'-------------------- 尚未完成適性測驗的時候 --------------- start -------------
		'response.write placementt
		if (placementt < 6) then
	   	
        %>
        <!--英語能力測試內容start-->
        <div class="main_classcontbox ">
            <!--英語能力測試過程列start-->
            <div class="main_classtest_processbar">
                <%=getWord("know_level")%><br />
                <ul class="process1">
                    <li>
                        <%=getWord("word")%></li>
                    <li>
                        <%=getWord("DCGS_9")%></li>
                    <li>
                        <%=getWord("grammar")%></li>
                    <li>
                        <%=getWord("DCGS_7")%></li>
                    <li>
                        <%=getWord("CLASS_VOCABULARY_8")%></li>
                    <li>
                        <%=getWord("finish")%></li>
                </ul>
                <div class="clear">
                </div>
            </div>
            <div style='width:100%;position:relative;float:left;'>
            <a class='online_evaluation_btn' href="online_evaluation.asp"><%=getWord("test_start")%></a>
            </div>
            
            <!--英語能力測試過程列end-->
            <!--英語能力測試前言內容start-->
            <p>
                <%=getWord("test_10_30_min")%><br />
                <%=getWord("test_horn")%>
                <!--<a href="#"><img src="../images_cn/icon_work.gif" border="0" /></a>-->
                <div id="flash_ability_playbtn" style="float: left;">
                    <%=getWord("please_used_flash_player_10_after")%></div>
                <script type="text/javascript">
                    var fo = new FlashObject("/flash/ability_playbtn.swf", "playbtn", "40", "40", "7", "#000000");
                    fo.addVariable("theFile", "/audio/list0qu1.mp3");
                    fo.addParam("wmode", "transparent");
                    fo.addParam("allowFullscreen", "true");
                    fo.write("flash_ability_playbtn");
                </script>
                <br />
                <%=getWord("no_horn")%>
                <a href="../contact/contact_us.asp"><span class="color_3">
                    <%=getWord("BUY_ACCESS_8")%></span></a></p>
            <div class="testgo">
                <!--<input type="submit" name="button" id="button" value="+ <%=getWord("BUY_ACCESS_8")%>" class="btn_1"/>-->
            </div>
            <!--英語能力測試前言內容end-->
        </div>
        <!--英語能力測試內容end-->
        </form>
        <%
		'-------------------- 尚未完成適性測驗的時候 --------------- end --------------
		else
        %>
        <ul class="class_list5">
            <li>
                <p class="red_title">
                    <%=getWord("o_level")%></p>
            </li>
            <%
		Dim arr_leveltype
		Dim int_j
		arr_leveltype = Array(getWord("DCGS_8"),getWord("CLASS_VOCABULARY_7"),getWord("CLASS_VOCABULARY_8"),getWord("DCGS_6"),getWord("DCGS_9"),getWord("DCGS_7"))
		
		for int_j = 1 to ubound(arr_leveltype)
            %>
            <li>
                <h1>
                    <%=arr_leveltype(int_j)%></h1>
                <h2>
                    <%
			Dim str_class : str_class = "green_ballNum2"
			for i = 1 to 12 
				if (i = stre1(int_j)) then 
					str_class =  "green_ballNum_go" 
				elseif ( i >= stre1(int_j)) then
					str_class =  "gray_ballNum2" 
				end if 
                    %>
                    <p class="<%=str_class%>">
                        <% if (i >= stre1(int_j)) then Response.Write i else Response.Write "" end if %>
                    </p>
                    <%next%>
                </h2>
                <%next%>
            </li>
        </ul>
        <ul class="class_list6">
            <li>
                <p class="red_title">
                    <%=getWord("learn_rate")%></p>
            </li>
            <%
		for int_j = 1 to ubound(arr_leveltype)
            %>
            <li>
                <h1>
                    <%=arr_leveltype(int_j)%></h1>
                <h2>
                    <%
			str_class = "green_ballNum2"
			for i = 1 to 12 
				if (i = stre2(int_j)) then 
					str_class =  "green_ballNum_go" 
				elseif ( i >= stre2(int_j)) then
					str_class =  "gray_ballNum2" 
				end if 
                    %>
                    <p class="<%=str_class%>">
                        <% if (i >= stre2(int_j)) then Response.Write i else Response.Write "" end if %>
                    </p>
                    <%next%>
                </h2>
                <%next%>
            </li>
        </ul>
        <div class="red_title">
            <%=getWord("test_record")%></div>
        <div>
            <table width="100%" border="0" cellpadding="0" cellspacing="0" class="bought_bd">
                <tr>
                    <th class="width_175">
                        <%=getWord("test_date")%>
                    </th>
                    <th class="width_175">
                        <%=getWord("start_time")%>
                    </th>
                    <th class="width_175">
                        <%=getWord("finish_time")%>
                    </th>
                    <th class="rightLine">
                        <%=getWord("score")%>
                    </th>
                </tr>
                <%
		str_sql = "SELECT inpdate,message3,message2,note "
		str_sql = str_sql & " from action(nolock) "
		str_sql = str_sql & " where client_sn = @client_sn and action_kind=@action_kind and valid=@valid  "
		str_sql = str_sql & " order by sn desc"

			var_arr = Array(g_var_client_sn,"TestQuestions",1)
			arr_result = excuteSqlStatementRead(str_sql, var_arr, CONST_VIPABC_RW_CONN)
		
			if (isSelectQuerySuccess(arr_result)) then
			
				'--------- 塞適性測驗 ------------ start ------------
				if (Ubound(arr_result) > 0) then 
					'有資料
					for int_act_i = 0 to ubound(arr_result,2)
						str_act_inpdate = Trim(arr_result(0,int_act_i))
						str_act_message3  = Trim(arr_result(1,int_act_i))
						str_act_message2  = Trim(arr_result(2,int_act_i))
						str_act_note  = Trim(arr_result(3,int_act_i))
						'處理模擬考日期、開始時間
						Dinpdate=DateValue(str_act_inpdate)
						Hinpdate=Right("0"&Hour(str_act_inpdate),2)
						Minpdate=Right("0"&Minute(str_act_inpdate),2)
												
						if (not isEmptyOrNull(str_act_message3)) then 
							'結束時間
							HDone=Right("0"&Hour(str_act_message3),2)
							MDone=Right("0"&Minute(str_act_message3),2)
							
							Done = HDone &":"& MDone
							DoneSocre = str_act_message2
						else
						
							Done = getWord("unfinish")		
							if (not isEmptyOrNull(str_act_message2) and str_act_message2 <> 0) or (not isEmptyOrNull(str_act_note) and str_act_message2 <> 0) then 
								 '為了之前沒記錄的資料加的判別
								Done = ""		'為了之前沒記錄的資料加的修改  
								DDone=DateAdd("n",45,str_act_inpdate) 'Patrick 指示 之前沒有紀錄的加45分鐘
								HDone=Right("0"&Hour(DDone),2)
								MDone=Right("0"&Minute(DDone),2)
								
								Done = HDone &":"& MDone
								DoneSocre = str_act_message2
							else
								DoneSocre = getWord("unfinish")	
							end if 
							
						end if						
                %>
                <tr>
                    <td>
                        <%=Dinpdate%>
                    </td>
                    <td>
                        <%=Hinpdate&":"&Minpdate%>
                    </td>
                    <td>
                        <%=Done%>
                    </td>
                    <td class="rightLine">
                        <%=DoneSocre%>
                    </td>
                </tr>
                <%
					next 
				else
				end if 
                %>
            </table>
            <%
			end if 
            %>
        </div>
        <div class="top_block">
        </div>
        <!--程度分析end-->
    </div>
    <!--內容end-->
    <%
        end if 
    %>
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>