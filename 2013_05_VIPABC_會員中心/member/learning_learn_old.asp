<!--#include virtual="/lib/include/html/template/template_change.asp"-->
<%

	'TODO: for test
	'g_var_client_sn = 147535	'100412

	'課堂購買明細
	'購買日期
	'購買課時數
	'贈送課時數
	'總課時數新增
	'合約
	Dim str_startdate, str_buyClass, str_freeClassNum, str_totalClass
	'沒有買過東西則將起始日期設定為2004/01/01
	str_startdate = "2004-1-1"
	str_buyClass = 0
	str_freeClassNum = 0
	str_totalClass = 0
	'合約
	'Dim str_contract : str_contract = "<a href=""view_contract.asp?csn=" & g_var_client_sn & """ target=""_blank"">详细内容</a>"
	Dim str_contract : str_contract = ""
	Dim int_payment : int_payment = 0
	Dim str_pro_sdate : str_pro_sdate = ""	'產品開始日期
	Dim str_pro_edate : str_pro_edate = ""	'產品結束日期

	'課堂使用紀錄
	'可使用總課時數
	'已使用課時數(含已預約)
	'目前剩餘堂數
	Dim str_available_hour, str_hour_used, str_remain_hour
	str_available_hour = 0
	str_hour_used = 0
	str_remain_hour = 0


	'課堂使用紀錄
	'可使用總錄影贈送堂數
	'已使用錄影贈送堂數
	'目前剩餘錄影贈送堂數
	Dim str_available_record_hour, str_record_used, str_remain_rec_hour
	str_available_record_hour = 0
	str_record_used = 0
	str_remain_rec_hour = 0


	'日期
	'時間
	'詳細說明
	'贈送錄影使用堂數
	Dim str_record_view_day, str_record_view_time, str_record_des, str_record_hour
	str_record_view_day = ""
	str_record_view_time = ""
	str_record_des = ""
	str_record_hour = 0


	'日期
	'時間
	'詳細說明
	'使用堂數
	Dim str_view_date, str_viewtime_start, str_viewtime_end, str_des, str_hour, str_viewtime
	str_view_date = ""
	str_viewtime = ""
	str_viewtime_start = ""
	str_viewtime_end = ""
	str_des = ""
	str_hour = 0

	'sql
	Dim str_sqlstr, arr_sql_query_input, arr_sql_query_result, arr_sql_query_result_2, int_sql_result_index
	
	Dim str_month, str_day
	
	Dim int_index_substring
	Dim int_index_for
	Dim date_today : date_today = Date()
%>
<!--內容start-->
<div class="main_membox">
	<div class='page_title_5'><h2 class='page_title_h2'>账户资讯</h2></div>
	<div class="box_shadow" style="background-position:-80px 0px;"></div> 
    <!--上版位課程start-->
    <!--#include virtual="/lib/include/html/include_learning_old.asp"-->
    <!--上版位課程end-->
    <div class="arrow_title">
        <%= getWord("class_buy_detail")%></div>
    <div>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="history_bd">
            <tr>
                <th>
                    <%= getWord("class_buy_day")%>
                </th>
                <th>
                    <%= getWord("class_buy_hour")%>
                </th>
                <th>
                    <%= getWord("class_gift_hour")%>
                </th>
                <th>
                    <%= getWord("class_add_hour")%>
                </th>
                <th>
                    <%= getWord("class_contact")%>
                </th>
                <th class="rightLine">
                    <%= getWord("NOTES")%>
                </th>
                <!-- 备注 -->
            </tr>
            <%
			'課堂購買明細
			'購買日期 '購買課時數 '贈送課時數 '總課時數新增 '合約
				
			'有合約才顯示購買明細, 在client_temporal_contract裡有記錄就是有合約, client_purchase裡的contract_sn有值就是在client_temporal_contract裡有記錄.
			'client_temporal_contract裡的cdata內記錄合約的內容
			str_sqlstr = "SELECT "
			str_sqlstr = str_sqlstr & "	contract_sn, appdate, free_point, add_points, total, serdate, enddate, pay_money, client_purchase.valid, product_sn, "
			str_sqlstr = str_sqlstr & "	client_temporal_contract.cdata "
			str_sqlstr = str_sqlstr & "From "
			str_sqlstr = str_sqlstr & "(	SELECT "
			str_sqlstr = str_sqlstr & "			contract_sn, "
			str_sqlstr = str_sqlstr & "			CONVERT(varchar, appdate, 111) AS appdate, "
			str_sqlstr = str_sqlstr & "			cast(access_points/65 as int) AS free_point, "
			str_sqlstr = str_sqlstr & "			cast(add_points/65 as int) as add_points, "
			str_sqlstr = str_sqlstr & "			cast((access_points + add_points)/65 as int) AS total, "
			str_sqlstr = str_sqlstr & "			CONVERT(varchar, product_sdate, 111) AS serdate, "
			str_sqlstr = str_sqlstr & "			CONVERT(varchar, product_edate, 111) AS enddate, "
			str_sqlstr = str_sqlstr & "			pay_money, valid, product_sn "
			str_sqlstr = str_sqlstr & "		FROM client_purchase "
			str_sqlstr = str_sqlstr & "		WHERE (client_sn = @client_sn) AND (Valid = 1) "
			str_sqlstr = str_sqlstr & ") client_purchase "
			str_sqlstr = str_sqlstr & " left outer Join "
			str_sqlstr = str_sqlstr & "(	Select sn, cdata, valid "
			str_sqlstr = str_sqlstr & "	From client_temporal_contract "
			str_sqlstr = str_sqlstr & ") client_temporal_contract On client_temporal_contract.sn = client_purchase.contract_sn "
			
			arr_sql_query_input = Array(g_var_client_sn)
			arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
			if (isSelectQuerySuccess(arr_sql_query_result) and  Ubound(arr_sql_query_result)>0) then
			
				For int_index_for = 0 to (Ubound(arr_sql_query_result,2))
					'取得client_purchase內的contract_sn值，這代表了是否有可用的合約
					If (Not isEmptyOrNull(arr_sql_query_result(0, int_index_for))) Then 
						str_contract =  arr_sql_query_result(0, int_index_for)
					Else
						str_contract = ""
					End If	
					
					If Not(isnull(str_contract) and str_contract<>"") Then
					
						if (Not isEmptyOrNull(arr_sql_query_result(1, int_index_for))) then 
							str_startdate =  arr_sql_query_result(1, int_index_for)
						Else
							str_startdate = ""
						end if
						
						If (IsDate(str_startdate)) Then
							str_startdate = Left(str_startdate, 4) & "-" & Mid(str_startdate, 6, 2) & "-" & Mid(str_startdate, 9, 2)
						Else
							str_startdate = ""
						End If
						
						if (Not isEmptyOrNull(arr_sql_query_result(3, int_index_for))) then 
							str_buyClass =  arr_sql_query_result(3, int_index_for)
						Else
							str_buyClass = ""
						end if
						
						if (Not isEmptyOrNull(arr_sql_query_result(2, int_index_for))) then 
							str_freeClassNum =  arr_sql_query_result(2, int_index_for)
						Else
							str_freeClassNum = ""
						end if
						
						if (Not isEmptyOrNull(arr_sql_query_result(4, int_index_for))) then 
							str_totalClass =  arr_sql_query_result(4, int_index_for)
						Else
							str_totalClass = ""
						end if	
						
						'堂數費用
						Dim int_payment_temp
						int_payment_temp = arr_sql_query_result(7, int_index_for)
						int_payment_temp = split(int_payment_temp,"|")
						Dim int_product_sn : int_product_sn = arr_sql_query_result(9, int_index_for)
						
						If (int_payment_temp(0) > 0) Then
							if(int_product_sn >= 41 and int_product_sn <= 51) then
								int_payment = Change_Money_To_Int(int_payment_temp(1))
							else
								int_payment = Change_Money_To_Int(int_payment_temp(1) - int_payment_temp(0))
							end if
						else
							int_payment = Change_Money_To_Int(int_payment_temp(1))
						end if
						
						if (Not isEmptyOrNull(arr_sql_query_result(5, int_index_for))) then 
							str_pro_sdate =  arr_sql_query_result(5, int_index_for)	'產品開始日期
						Else
							str_pro_sdate = ""
						end if	
						if (Not isEmptyOrNull(arr_sql_query_result(6, int_index_for))) then 
							str_pro_edate =  arr_sql_query_result(6, int_index_for)	'產品結束日期
						Else
							str_pro_edate = ""
						end if	
						
						
						int_product_sn = arr_sql_query_result(9, int_index_for)
					
					
						str_sqlstr = "SELECT "
						str_sqlstr = str_sqlstr & " sn, 1 "
						str_sqlstr = str_sqlstr & " FROM client_temporal_contract "
						str_sqlstr = str_sqlstr & " WHERE (client_sn = @client_sn) "
						str_sqlstr = str_sqlstr & "	AND (prodbuy = @prodbuy) "
						str_sqlstr = str_sqlstr & "	AND (cagree = 1) AND (valid = 1) "
						str_sqlstr = str_sqlstr & "	AND (cdata is not null)"
						
						arr_sql_query_input = Array(g_var_client_sn, int_product_sn)
						arr_sql_query_result_2 = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
						'response.Write g_str_sql_statement_for_debug
						If ( isSelectQuerySuccess(arr_sql_query_result_2) and  (Ubound(arr_sql_query_result_2)>0) ) Then
							str_contract = "<a href='contract.asp?contract_sn=" & arr_sql_query_result_2(0,0) & "'>" & getWord("PREVIEW_CONTRACT") & "</a>"
						Else
							str_contract = ""
						End If
						
                        '20120316 阿捨新增 無限卡產品判斷	
                        dim bolUnlimited : bolUnlimited = false
                        dim strUnlimited : strUnlimited = ""
                        if ( Instr(CONST_NO_LIMIT_PRODUCT_SN, (int_product_sn & ",")) > 0 ) then
                            bolUnlimited = true
                            strUnlimited = "unlimited"
                        end if

                        '20121025 阿捨新增 終身無限卡產品判斷	
                        Dim bolLifeUnlimited : bolLifeUnlimited = false
                        dim strLifeUnlimited : strLifeUnlimited = ""
                        if ( Instr(CONST_NO_LIMIT_FOREVER_PRODUCT_SN, (int_product_sn & ",")) > 0 ) then
                            bolLifeUnlimited = true
                            strLifeUnlimited = "終身"
                        end if

						'20120606 VM12053006 學習儲值卡官網+IMS功能 Penny
						Dim bolLearningCard : bolLearningCard = false
						if ( Not isEmptyOrNull(int_product_sn) ) then
							if ( Instr(CONST_VIPABC_LEARNING_CARD_PRODUCT_SN, ("," & int_product_sn & ",")) > 0 ) then
								bolLearningCard = true
							end if
						end if	

						'20120719 Glen新增团购卡(5堂/效期7天)產品判斷 product sn=493        
						Dim bolDianPingGroupbuy : bolDianPingGroupbuy = false
						if ( Not isEmptyOrNull(int_product_sn) ) then						
							if ( Instr(CONST_DIANPING_GROUPBUY_PRODUCT_SN, (int_product_sn & ",")) > 0 ) then
								bolDianPingGroupbuy = true
							end if		
						end if	
            %>
            <tr>
                <td>
                    <%=str_startdate%>
                </td>
                <td>
					<% 
                        'if ( true = bolLearningCard OR true = bolDianPingGroupbuy ) then 
						    'response.Write str_buyClass 
					    'else
                        if ( true = bolUnlimited ) then
						    response.Write strUnlimited
                        else 
					        response.Write str_buyClass
                        end if
					%>
                </td>
                <td>
                    <%=str_freeClassNum%>
                </td>
                <td>
					<%  
                    'if ( true = bolLearningCard OR true = bolDianPingGroupbuy ) then 
						'response.Write str_totalClass 
					'else
                    if ( true = bolUnlimited ) then 
				        response.Write strUnlimited
                    else
                        response.Write str_totalClass
					end if
					%>
                </td>
                <td>
                    <%=str_contract%>
                </td>
                <td class="rightLine">
                    <!-- 课时数费用 -->
                    <!-- RMB -->
                    <!--<%= getWord("PAYMENT_OF_CLASS")%> <%= getWord("RMB")%> $<%=int_payment%><br />	
					-->
                    <!--產品期限-->
                    <%= getWord("PRODUCT_PERIOD")%><%=str_pro_sdate%>
                    -
                    <%
                    if ( true = bolLifeUnlimited ) then
                        response.Write strLifeUnlimited
                    else
                        response.Write str_pro_edate
                    end if
                     %>
                </td>
            </tr>
            <%
					end if
				next
			else
            %>
            <!-- <tr><td colspan="6" align="center">No Data</td></tr> -->
            <%
			end if
            %>
        </table>
        <div class="clear">
        </div>
    </div>
    <div class="arrow_title">
        <%= getWord("class_used_record")%></div>
    <div>
        <%
			'課堂使用紀錄
	        '可使用總課時數, 以使用課時數(含已預約), 目前剩餘堂數
			'取得可使用總課時數
			str_sqlstr = 	"SELECT cast((fp + rp)/65 as int) AS tp " &_
							" FROM " &_
							"	(SELECT SUM(add_points) AS fp, SUM(access_points) AS rp, " &_
							"		min(product_sdate) AS sdate, max(product_edate) AS edate " &_
							"	FROM client_purchase " &_
							"	WHERE (client_sn = @Sn ) " &_
							"		AND (Valid = 1) " &_
							"		AND (product_sdate <= @date1 ) " &_
							"		AND (product_edate >= @date2 )" &_
							"	) DERIVEDTBL"

			arr_sql_query_input = Array(g_var_client_sn, date_today, date_today)
			arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
			if ( isSelectQuerySuccess(arr_sql_query_result) ) then
				if ( Ubound(arr_sql_query_result) >= 0 ) then 
					'有資料
					if ( Not isEmptyOrNull(arr_sql_query_result(0, 0)) ) then
						str_available_hour =  arr_sql_query_result(0, 0)
					end if
				else
					'無資料
				end if 
			end if

			'取得已使用總課時數
			str_sqlstr = 	"SELECT SUM(pused) / 65 AS sused " &_
							" FROM " &_
							"	(SELECT (CASE dbo.client_attend_list.refund WHEN 1 THEN 0 ELSE ((isnull(client_attend_list.svalue,0)+0.00)*65) END) AS pused " &_
							"	FROM client_attend_list " &_
							"	LEFT OUTER JOIN cfg_live_session_points " &_
							"		ON client_attend_list.attend_livesession_types = cfg_live_session_points.sn " &_
							"	RIGHT OUTER JOIN " &_
							"		(SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate " &_
							"		FROM client_purchase " &_
							"		WHERE (client_sn = @Sn1 ) " &_
							"			AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, GETDATE(), 111)) " &_
							"			AND  (CONVERT(varchar, product_edate, 111) >= CONVERT(varchar, GETDATE(), 111)) " &_
							"			AND (Valid = 1) GROUP BY client_sn" &_
							"		) AS derivedtbl_1 " &_
							"	ON CONVERT(varchar, client_attend_list.attend_date, 111) <= derivedtbl_1.edate " &_
							"	AND CONVERT(varchar, client_attend_list.attend_date, 111) >= derivedtbl_1.sdate " &_
							"	AND client_attend_list.client_sn = derivedtbl_1.client_sn  " &_
							"	WHERE (client_attend_list.valid = 1) " &_
							"		AND (client_attend_list.refund IN (0, 1))  " &_
							"	UNION ALL " &_
							"		SELECT sessionrecord_view_list.view_cost AS pused " &_
							"		FROM sessionrecord_view_list " &_
							"		RIGHT OUTER JOIN " &_
							"			(SELECT  client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate " &_
							"			FROM client_purchase  " &_
							"			WHERE (client_sn =  @Sn2 ) " &_
							"				AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, GETDATE(), 111)) " &_
							"				AND (CONVERT(varchar, product_edate, 111) >= CONVERT(varchar, GETDATE(), 111)) " &_
							"				AND (Valid = 1) GROUP BY client_sn" &_
							"			) AS derivedtbl_1_1 " &_
							"		ON CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) <= derivedtbl_1_1.edate " &_
							"			AND CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) >= derivedtbl_1_1.sdate " &_
							"			AND sessionrecord_view_list.client_sn = derivedtbl_1_1.client_sn  " &_
							"		WHERE   (sessionrecord_view_list.client_sn IS NOT NULL)" &_
							"	) AS a "
			arr_sql_query_input = Array(g_var_client_sn, g_var_client_sn)
			arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
			if (isSelectQuerySuccess(arr_sql_query_result)) then
				if (Ubound(arr_sql_query_result) >= 0) then 
					'有資料
					if (Not isEmptyOrNull(arr_sql_query_result(0, 0))) then 
						str_hour_used =  arr_sql_query_result(0, 0)
					end if
				else
					'無資料
				end if 
			end if

			str_remain_hour = str_available_hour - str_hour_used
        %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="bought_bd">
            <tr>
                <th>
                    <%= getWord("class_use_hour")%>
                </th>
                <th rowspan="2" class="width_30">
                    -
                </th>
                <th>
                    <%= getWord("class_used_hour")%>
                </th>
                <th rowspan="2" class="width_30">
                    =
                </th>
                <th class="rightLine">
                    <%= getWord("class_remnant_hour")%>
                </th>
            </tr>
            <tr>
                <td>
					<%
                    'if ( true = bolLearningCard or true = bolDianPingGroupbuy ) then 
						'response.Write str_available_hour 
			        'else
                    if ( true = bolUnlimited ) then 
				        response.Write strUnlimited
                    else
                        response.Write str_available_hour
		            end if
					%>
                </td>
                <td>
                    <%=str_hour_used%>
                </td>
                <td class="rightLine">
					<%
                        'if ( true = bolLearningCard or true = bolDianPingGroupbuy ) then 
						    'response.Write str_remain_hour 
					   'else
                        if ( true = bolUnlimited ) then 
						    response.Write strUnlimited
                        else
                            response.Write str_remain_hour
					    end if
					%>
                </td>
            </tr>
        </table>
        <div class="clear">
        </div>
    </div>
    <div>
        <%
			'課堂使用紀錄
				'可使用總錄影贈送堂數, 已使用錄影贈送堂數, 目前剩餘錄影贈送堂數

			str_sqlstr = 	" SELECT SUM(free_video_point) AS total_free_point, SUM(used_video_point) AS last_free_point " &_
							" FROM client_free_video_point WHERE (client_sn =  @Sn ) "
			arr_sql_query_input = Array(g_var_client_sn)
			arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
			if (isSelectQuerySuccess(arr_sql_query_result)) then
				if (Ubound(arr_sql_query_result) >= 0) then 
					'有資料
					if (Not isEmptyOrNull(arr_sql_query_result(0, 0))) then 
						str_available_record_hour =  arr_sql_query_result(0, 0)
					end if
					if (Not isEmptyOrNull(arr_sql_query_result(1, 0))) then 
						str_remain_rec_hour =  arr_sql_query_result(1, 0)
					end if
				else
					'無資料
				end if 
			end if
            '阿捨修正 贈送錄影檔剩餘堂數
            str_remain_rec_hour = sCDbl(isHavingFreeRecord(g_var_client_sn, 2, 0, CONST_VIPABC_RW_CONN)) / 65
			str_record_used =  sCDbl(str_available_record_hour) - sCDbl(str_remain_rec_hour)
        %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="bought_bd">
            <tr>
                <th>
                    <%= getWord("class_gift_video_hour")%>
                </th>
                <th rowspan="2" class="width_30">
                    -
                </th>
                <th>
                    <%= getWord("class_used_video_hour")%>
                </th>
                <th rowspan="2" class="width_30">
                    =
                </th>
                <th class="rightLine">
                    <%= getWord("class_remnant_video_hour")%>
                </th>
            </tr>
            <tr>
                <td>
					<%
                        'if ( true = bolLearningCard or true = bolDianPingGroupbuy ) then 
						    'response.Write str_available_record_hour 
					    'else
                        if ( true = bolUnlimited ) then
				            response.Write strUnlimited 
					    else
                            response.Write str_available_record_hour
                        end if
					%>
                </td>
                <td>
                    <%=str_record_used%>
                </td>
                <td class="rightLine">
					<% 
                        'if ( true = bolLearningCard or true = bolDianPingGroupbuy ) then 
						    'response.Write str_remain_rec_hour 
					    'else
                        if ( true = bolUnlimited ) then 
						    response.Write strUnlimited 
                        else
                            response.Write str_remain_rec_hour
					    end if
					%>
                </td>
            </tr>
        </table>
        <div class="clear">
        </div>
    </div>
    <div>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="bought_bd">
            <tr>
                <th class="width_100">
                    <%= getWord("class_date")%>
                </th>
                <th class="width_100">
                    <%= getWord("class_time")%>
                </th>
                <th>
                    <%= getWord("DETAIL_DESCRIPTION")%>
                </th>
                <!--
                <th class="rightLine">
                    <%= getWord("class_gift_video")%><br />
                    <%= getWord("BUY_1_18")%>
                </th>
                -->
            </tr>
            <%
				'日期 '時間 '詳細說明 '贈送錄影使用堂數

				str_sqlstr =	" SELECT " &_
								"	sessionrecord_view_used_free_video_point_list.view_datetime, " &_
								"	material.ltitle AS session_plant, " &_
								"	sessionrecord_view_used_free_video_point_list.view_cost / 65 AS view_cost " &_
								" FROM sessionrecord_view_used_free_video_point_list " &_
								" LEFT OUTER JOIN material ON sessionrecord_view_used_free_video_point_list.Course = material.course " &_
								" WHERE (sessionrecord_view_used_free_video_point_list.client_sn = @Sn ) " &_
								"	AND (sessionrecord_view_used_free_video_point_list.view_state = '1') "

				int_sql_result_index = 0
				arr_sql_query_input = Array(g_var_client_sn)
				arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)
				if (isSelectQuerySuccess(arr_sql_query_result)) then
					if (Ubound(arr_sql_query_result) >= 0) then 
						'有資料
						for int_index_for = 0 to Ubound(arr_sql_query_result,2) 
							if (Not isEmptyOrNull(arr_sql_query_result(0, int_index_for))) then 
								str_record_view_day = arr_sql_query_result(0, int_index_for)
								
								'2009-12-08 11:14:18.000 / 2009-12-8 -> 2009-12-08
								if (IsDate(str_record_view_day)) then
									str_record_view_day = CDate(str_record_view_day)
									str_month = CStr(Month(str_record_view_day))
									str_day = CStr(Day(str_record_view_day))
									
									if (Len(str_month)=1) then
										str_month = "0" & str_month
									end if
									if (Len(str_day)=1) then
										str_day = "0" & str_day
									end if
									str_record_view_day = Year(str_record_view_day) & "-" & str_month & "-" & str_day
								end if
								
								str_record_view_time = FormatDateTime(arr_sql_query_result(0, int_index_for),4)
							end if
							
							if (Not isEmptyOrNull(arr_sql_query_result(1, int_index_for))) then 
								str_record_des = arr_sql_query_result(1, int_index_for)
							end if
							
							if (Not isEmptyOrNull(arr_sql_query_result(2, int_index_for))) then 
								str_record_hour = arr_sql_query_result(2, int_index_for)
							end if
							
							int_sql_result_index = int_sql_result_index + 1
            %>
            <tr>
                <td>
                    <%=str_record_view_day%>
                </td>
                <td>
                    <%=str_record_view_time%>
                </td>
                <td class="main_ct">
                    <%=str_record_des%>
                </td>
                <!--
                <td class="rightLine">
                    <%=str_record_hour%><%= getWord("class")%>
                </td>
                -->
            </tr>
            <%
						next
					else
						'無資料
					end if
				end if

				if (int_sql_result_index = 0) then
            %>
            <tr>
                <td colspan="4" align="center">
                    No Data
                </td>
            </tr>
            <%
				end if
            %>
        </table>
        <div class="clear">
        </div>
    </div>
    <div>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="bought_bd">
            <tr>
                <th class="width_100">
                    <%= getWord("CLASS_DATE")%>
                </th>
                <th class="width_100">
                    <%= getWord("CLASS_TIME")%>
                </th>
                <th>
                    <%= getWord("DETAIL_DESCRIPTION")%>
                </th>
                <!--
                <th class="rightLine">
                    <%= getWord("BUY_1_18")%>
                </th>
                -->
            </tr>
            <%
				'日期 '時間 '詳細說明 '使用堂數

				'str_sqlstr="SELECT regular.client_sn, regular.sdate, regular.stime, regular.pused * ISNULL(derivedtbl_2.discount, 1) AS pused, regular.lessonplan, regular.session_sn, left(convert(varchar,regular.attend_datetime,108),5) attend_datetime FROM (SELECT client_sn, sdate, stime, pused, lessonplan, session_sn, YEAR(sdate) AS yreal, MONTH(sdate) AS mreal, attend_datetime FROM (SELECT client_attend_list.client_sn, CONVERT(varchar, client_attend_list.attend_date, 111) AS sdate, client_attend_list.attend_sestime AS stime, (CASE dbo.client_attend_list.refund WHEN 1 THEN 0 ELSE ((isnull(client_attend_list.svalue,0)+0.00)*65) END) AS pused, material.ltitle AS lessonplan, client_attend_list.session_sn, client_attend_list.attend_datetime FROM client_attend_list LEFT OUTER JOIN material ON client_attend_list.attend_mtl_1 = material.course LEFT OUTER JOIN cfg_live_session_points ON client_attend_list.attend_livesession_types = cfg_live_session_points.sn RIGHT OUTER JOIN (SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate FROM client_purchase WHERE (client_sn = @Sn1) /*--Jessica delete by CS08122601-- AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, GETDATE(), 111)) AND (CONVERT(varchar, product_edate, 111) >= CONVERT(varchar, GETDATE(), 111))*/ AND (Valid = 1) GROUP BY client_sn) AS derivedtbl_1 ON CONVERT(varchar, client_attend_list.attend_date, 111) <= derivedtbl_1.edate AND CONVERT(varchar, client_attend_list.attend_date, 111) >= derivedtbl_1.sdate AND client_attend_list.client_sn = derivedtbl_1.client_sn WHERE (client_attend_list.valid = 1) AND (client_attend_list.refund IN (0, 1)) UNION ALL SELECT sessionrecord_view_list.client_sn, CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) AS sdate, LEFT(CONVERT(varchar, sessionrecord_view_list.view_datetime, 114), 2) AS stime, sessionrecord_view_list.view_cost AS pused, material_1.ltitle AS lessonplan, sessionrecord_view_list.view_session_recordfiles_sn AS session_sn, sessionrecord_view_list.view_datetime FROM material AS material_1 RIGHT OUTER JOIN sessionrecord_view_list ON material_1.course = sessionrecord_view_list.Course RIGHT OUTER JOIN (SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate FROM client_purchase AS client_purchase_1 WHERE (client_sn = @Sn2) /*--Jessica delete by CS08122601-- AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, GETDATE(), 111)) AND (CONVERT(varchar, product_edate, 111) >= CONVERT(varchar, GETDATE(), 111))*/ AND (Valid = 1) GROUP BY client_sn) AS derivedtbl_1_1 ON CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) <= derivedtbl_1_1.edate AND CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) >= derivedtbl_1_1.sdate AND sessionrecord_view_list.client_sn = derivedtbl_1_1.client_sn WHERE (sessionrecord_view_list.client_sn IS NOT NULL)) AS a) AS regular LEFT OUTER JOIN (SELECT sn, client_sn, dyear, dmonth, discount, valid FROM client_points_discount WHERE (valid = 1)) AS derivedtbl_2 ON regular.yreal = derivedtbl_2.dyear AND regular.mreal = derivedtbl_2.dmonth AND regular.client_sn = derivedtbl_2.client_sn ORDER BY regular.sdate DESC, regular.stime DESC, attend_datetime desc"
				str_sqlstr = 			"SELECT "
				str_sqlstr = str_sqlstr & "regular.client_sn, regular.sdate, regular.stime, regular.pused * ISNULL(derivedtbl_2.discount, 1) AS pused, regular.lessonplan, regular.session_sn, left(convert(varchar,regular.attend_datetime,108),5) attend_datetimem, regular.ClientAttendListSn, sst_number "
				str_sqlstr = str_sqlstr & "FROM "
				str_sqlstr = str_sqlstr & "(	SELECT "
				str_sqlstr = str_sqlstr & "			ClientAttendListSn, client_sn, sdate, stime, pused, lessonplan, session_sn, YEAR(sdate) AS yreal, MONTH(sdate) AS mreal, attend_datetime, "
				'這堂課可以退定的最後時間
				str_sqlstr = str_sqlstr & "			DATEADD(n, 30, (DATEADD(hh, (stime-4), sdate)) ) As deadline, sst_number "
				str_sqlstr = str_sqlstr & "		FROM "
				str_sqlstr = str_sqlstr & "		(	SELECT client_attend_list.sn AS ClientAttendListSn, client_attend_list.client_sn, CONVERT(varchar, client_attend_list.attend_date, 111) AS sdate, client_attend_list.attend_sestime AS stime, (CASE dbo.client_attend_list.refund WHEN 1 THEN 0 ELSE ((isnull(client_attend_list.svalue,0)+0.00)*65) END) AS pused, material.ltitle AS lessonplan, client_attend_list.session_sn, client_attend_list.attend_datetime, sst_number "
				str_sqlstr = str_sqlstr & "			FROM client_attend_list "
				str_sqlstr = str_sqlstr & "			LEFT OUTER JOIN material ON client_attend_list.attend_mtl_1 = material.course "
				'str_sqlstr = str_sqlstr & "			LEFT OUTER JOIN cfg_live_session_points ON client_attend_list.attend_livesession_types = cfg_live_session_points.sn "
				str_sqlstr = str_sqlstr & "			RIGHT OUTER JOIN "
				str_sqlstr = str_sqlstr & "			(	SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate "
				str_sqlstr = str_sqlstr & "				FROM client_purchase "
				str_sqlstr = str_sqlstr & "				WHERE (client_sn = @Sn1) "
				str_sqlstr = str_sqlstr & "					AND (Valid = 1) "
				str_sqlstr = str_sqlstr & "				GROUP BY client_sn "
				str_sqlstr = str_sqlstr & "			) AS client_purchase_a "
				str_sqlstr = str_sqlstr & "				ON CONVERT(varchar, client_attend_list.attend_date, 111) <= client_purchase_a.edate "
				str_sqlstr = str_sqlstr & "					AND CONVERT(varchar, client_attend_list.attend_date, 111) >= client_purchase_a.sdate "
				str_sqlstr = str_sqlstr & "					AND client_attend_list.client_sn = client_purchase_a.client_sn "
				str_sqlstr = str_sqlstr & "			LEFT OUTER JOIN cfgSessionType ON (		/*大會堂判斷*/ "
				str_sqlstr = str_sqlstr & "				case when client_attend_list.attend_livesession_types = 4 "
				str_sqlstr = str_sqlstr & "				AND client_attend_list.special_sn is NOT NULL "
				str_sqlstr = str_sqlstr & "				THEN 99 "
				str_sqlstr = str_sqlstr & "				ELSE  client_attend_list.attend_livesession_types end "
				str_sqlstr = str_sqlstr & "			)= cfgSessionType.sst_number AND cfgSessionType.web_site = 'C' "
				str_sqlstr = str_sqlstr & "			WHERE (client_attend_list.valid = 1) AND (client_attend_list.refund IN (0, 1)) "
				str_sqlstr = str_sqlstr & "			UNION ALL "
				str_sqlstr = str_sqlstr & "				SELECT "
				str_sqlstr = str_sqlstr & "					0 AS ClientAttendListSn, sessionrecord_view_list.client_sn, CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) AS sdate, LEFT(CONVERT(varchar, sessionrecord_view_list.view_datetime, 114), 2) AS stime, sessionrecord_view_list.view_cost AS pused, material_1.ltitle AS lessonplan, sessionrecord_view_list.view_session_recordfiles_sn AS session_sn, sessionrecord_view_list.view_datetime, '100' AS 'sst_number' "
				str_sqlstr = str_sqlstr & "				FROM material AS material_1 "
				str_sqlstr = str_sqlstr & "				RIGHT OUTER JOIN sessionrecord_view_list ON material_1.course = sessionrecord_view_list.Course "
				str_sqlstr = str_sqlstr & "				RIGHT OUTER JOIN "
				str_sqlstr = str_sqlstr & "				(	SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate "
				str_sqlstr = str_sqlstr & "					FROM client_purchase "
				str_sqlstr = str_sqlstr & "					WHERE (client_sn = @Sn2) AND (Valid = 1) "
				str_sqlstr = str_sqlstr & "					GROUP BY client_sn "
				str_sqlstr = str_sqlstr & "				) AS client_purchase_r "
				str_sqlstr = str_sqlstr & "					ON CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) <= client_purchase_r.edate "
				str_sqlstr = str_sqlstr & "						AND CONVERT(varchar, sessionrecord_view_list.view_datetime, 111) >= client_purchase_r.sdate "
				str_sqlstr = str_sqlstr & "						AND sessionrecord_view_list.client_sn = client_purchase_r.client_sn "
				str_sqlstr = str_sqlstr & "				WHERE (sessionrecord_view_list.client_sn IS NOT NULL) "
				str_sqlstr = str_sqlstr & "		) AS regular "
				str_sqlstr = str_sqlstr & ") AS regular "
				str_sqlstr = str_sqlstr & "LEFT OUTER JOIN "
				str_sqlstr = str_sqlstr & "(	SELECT sn, client_sn, dyear, dmonth, discount, valid "
				str_sqlstr = str_sqlstr & "		FROM client_points_discount "
				str_sqlstr = str_sqlstr & "		WHERE (valid = 1) "
				str_sqlstr = str_sqlstr & ") AS derivedtbl_2 "
				str_sqlstr = str_sqlstr & "	ON regular.yreal = derivedtbl_2.dyear "
				str_sqlstr = str_sqlstr & "		AND regular.mreal = derivedtbl_2.dmonth "
				str_sqlstr = str_sqlstr & "		AND regular.client_sn = derivedtbl_2.client_sn "
				'str_sqlstr = str_sqlstr & "Where GETDATE() > deadline "
				str_sqlstr = str_sqlstr & "ORDER BY regular.sdate DESC, regular.stime DESC, attend_datetime desc "

				'response.Write(g_var_client_sn & "<br>")
				'response.Write("abc" & "<br>" & str_sqlstr & "<br>")
				'Response.End
				
				arr_sql_query_input = Array(g_var_client_sn, g_var_client_sn)
				arr_sql_query_result = excuteSqlStatementRead( str_sqlstr, arr_sql_query_input, CONST_VIPABC_RW_CONN)

				if (isSelectQuerySuccess(arr_sql_query_result)) then
					if (Ubound(arr_sql_query_result) >= 0) then 
						'有資料
						for int_index_for = 0 to Ubound(arr_sql_query_result,2) 
							'2009/12/07  -> 2009-12-07
							if (Not isEmptyOrNull(arr_sql_query_result(0, int_index_for))) then 
								str_view_date = arr_sql_query_result(1, int_index_for)
								str_view_date = Left(str_view_date,4) & "-" & Mid(str_view_date,6,2) & "-" & Right(str_view_date,2)
							end if
							
							if("100" = arr_sql_query_result(8, int_index_for)) then
								'錄影檔
								str_viewtime = arr_sql_query_result(6, int_index_for)
							else
								str_viewtime = getSessionDateTime(arr_sql_query_result(7, int_index_for), arr_sql_query_result(1, int_index_for), arr_sql_query_result(2, int_index_for), arr_sql_query_result(5, int_index_for), 6, 0, CONST_VIPABC_RW_CONN)
							end if
							if (Not isEmptyOrNull(arr_sql_query_result(6, int_index_for))) and (Not isEmptyOrNull(arr_sql_query_result(0, int_index_for))) then 
								if len(arr_sql_query_result(2, int_index_for))=1 then
									str_viewtime_start = FormatDateTime(Cdate(arr_sql_query_result(1, int_index_for) & " " & "0" & arr_sql_query_result(2, int_index_for) & ":30:00"), 4)
								else
									str_viewtime_start = FormatDateTime(Cdate(arr_sql_query_result(1, int_index_for) & " " & arr_sql_query_result(2, int_index_for) & ":30:00"), 4)
								end if
								str_viewtime_end = FormatDateTime(DateAdd("n",45,str_viewtime_start),4)
							end if
							
							'取得課程描述
							if (Not isEmptyOrNull(arr_sql_query_result(4, int_index_for))) then
								str_des = arr_sql_query_result(4, int_index_for)
							else
								str_des = "&nbsp;"
							end if

							'看狀況、加上link
							if (Len(str_viewtime_start)>0) and (DateDiff("n", (Cdate(arr_sql_query_result(1, int_index_for) & " " & arr_sql_query_result(2, int_index_for) & ":30:00")),Now)>10) then
								if (Not isEmptyOrNull(arr_sql_query_result(3, int_index_for))) and (Not isEmptyOrNull(arr_sql_query_result(5, int_index_for))) and (arr_sql_query_result(3, int_index_for)>=35) and (Len(arr_sql_query_result(5, int_index_for))=13) then
									str_des = "<a href=""learning_record_detail.asp?session_sn=" & arr_sql_query_result(5, int_index_for) & "&read_only=y"""&_
												" class=""account_info"">" & arr_sql_query_result(4, int_index_for) & "</a>"
								end if
							else
								if (Not isEmptyOrNull(arr_sql_query_result(3, int_index_for))) and (arr_sql_query_result(3, int_index_for)<>"20") then
										str_des = "<b>session time has not yet arrived.</b>"
								end if
							end if							
							
							if (Not isEmptyOrNull(arr_sql_query_result(3, int_index_for))) then 
								str_hour = Round(Round(arr_sql_query_result(3, int_index_for), 4)/65 , 2)

								if str_hour=0.0 then 
									str_hour =  getWord("no_deduct_class")
								else
									str_hour = str_hour & "&nbsp; "&getWord("class")
								end if
							end if
            %>
            <tr>
                <!-- 日期 '時間 '詳細說明 '使用堂數 -->
                <td>
                    <%=str_view_date%>&nbsp;
                </td>
                <td>
                    <%=str_viewtime%>&nbsp;
                </td>
                <td class="main_ct">
                    &nbsp;<%=str_des%>&nbsp;
                </td>
                <!--
                <td class="rightLine">
                    <%=str_hour%>&nbsp;
                </td>
                -->
            </tr>
            <%
						next
					else
            %>
            <tr>
                <td colspan="4" align="center">
                    No Data
                </td>
            </tr>
            <%
					end if
				end if
            %>
        </table>
        <div class="clear">
        </div>
    </div>
    <!--	
	<div class="page" style="display:none;">
		<a href="#"><%=getWord("last_page")%></a>
		<a href="#">1</a>
		<a  href="#">2</a>
		<a  href="#">3</a>
		<a  href="#">4</a>
		<a  href="#">5</a>
		<a  href="#">6</a>
		<a  href="#">7</a>
		<a  href="#">8</a>
		<a  href="#"><%=getWord("next_page")%></a>
	</div>
-->
    <font id="<%=CONST_HTML_MAIN_CONTENTS_DIV_ID%>"></font>
</div>
<!--內容end-->
<%
Function Change_Money_To_Int(ByVal InputStr)
	IF InputStr="" Then
		Change_Money_To_Int = ""
		Exit Function
	End IF

	On Error Resume Next 
	Dim str_result
	Dim i
	str_result = ""
	IF ( (Len(Cstr(InputStr))\3 >= 1) and  (Len(Cstr(InputStr))<>3) ) Then
	   For i= 1 To Len(Cstr(InputStr))\3
		   str_result = "," & Right(InputStr,3) & str_result
		   InputStr = Left(Inputstr,Len(Inputstr)-3)
	   Next
	   str_result = InputStr & str_result
	Else
	   str_result = InputStr
	End IF
	IF (Left(str_result,1)=",") Then
	   str_result = Mid(str_result,2,Len(str_result))
	End IF

	Change_Money_To_Int = str_result
End Function
%>