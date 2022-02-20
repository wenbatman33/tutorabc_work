<!--#include virtual="/lib/functions/functions_member.asp"-->
<%




function getCustomerBookingTime(ByVal p_str_order_class_date, ByVal p_str_client_sn, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：則檢查該日期客戶是否已有訂課 並取得訂課的時間
'傳入參數  ：[p_str_order_class_date:String] 客戶訂課日期
'            [p_str_client_sn:Integer] 客戶編號
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：訂課的時間 (Format: ",6,16,15,")
'牽涉數據表：
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/19 Created
'*****************************************************************************************/
    Dim str_sql, arr_param_value, arr_attend_list_data, i, str_booking_time

    str_booking_time = ""

    '若有傳入日期 則檢查該日期客戶是否已有訂課
    if (Not isEmptyOrNull(p_str_order_class_date) and IsDate(p_str_order_class_date) and Not isEmptyOrNull(p_str_client_sn) and Not isEmptyOrNull(p_int_connect_type)) then            
        str_sql = " SELECT attend_sestime FROM client_attend_list " &_
                  " WHERE (valid = 1) AND (CONVERT(varchar, client_attend_list.attend_date, 111) = CONVERT(varchar, @sdate, 111)) " &_
                  " AND (client_attend_list.client_sn = @client_sn) "
        arr_param_value = Array(p_str_order_class_date, p_str_client_sn)
        arr_attend_list_data = excuteSqlStatementRead(str_sql, arr_param_value, p_int_connect_type)
        if (isSelectQuerySuccess(arr_attend_list_data)) then
            if (Ubound(arr_attend_list_data) >= 0) then
                for i = 0 to Ubound(arr_attend_list_data, 2)
                    str_booking_time = str_booking_time & arr_attend_list_data(0, i) & ","
                next
                str_booking_time = "," & str_booking_time
            end if
        end if
    end if
    getCustomerBookingTime = str_booking_time
end function

Function getCustomerOneManyBookingCount(ByVal p_str_client_sn, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：取得客戶可以訂位(一對多)的堂數
'傳入參數  ：[p_str_client_sn:Integer] 客戶編號
'回傳值    ：可以訂課的堂數 (Integer)
'牽涉數據表：client_purchase
'備註      ：因為開放六天前就可以訂位(合約開始那一日的課)，所以這邊查詢的時間區間(合約)為六天後之前開始 或 今天之後結束
'歷程(作者/日期/原因) ：[Lucy Chen] 2010/02/22 Created
'*****************************************************************************************/
	Dim str_sql, arr_param_value, arr_attend_list_data

	If (Not isEmptyOrNull(p_str_client_sn) and IsNumeric(p_str_client_sn)) Then
		str_sql = "		  SELECT "
		str_sql = str_sql & "	sum(access_points) access_points, "
		str_sql = str_sql & "	sum(add_points) add_points "
		str_sql = str_sql & " FROM dbo.client_purchase "
		str_sql = str_sql & " WHERE (client_sn = @client_sn) "
		str_sql = str_sql & "	AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, DATEADD(day,6,GETDATE()), 111)) "
		str_sql = str_sql & "	AND (CONVERT(varchar, GETDATE(), 111) <= CONVERT(varchar, product_edate, 111)) "
		str_sql = str_sql & "	AND (Valid = 1) "
		str_sql = str_sql & " group by client_sn "
		
		arr_param_value = Array(p_str_client_sn)
        arr_attend_list_data = excuteSqlStatementRead(str_sql, arr_param_value, p_int_connect_type)
		
		If (isSelectQuerySuccess(arr_attend_list_data)) Then
			If (Ubound(arr_attend_list_data) >= 0) then
				'取得所有可使用的點數
				getCustomerOneManyBookingCount = CInt(arr_attend_list_data(0, 0))
				getCustomerOneManyBookingCount = getCustomerOneManyBookingCount + CInt(arr_attend_list_data(1, 0))
			Else
				getCustomerOneManyBookingCount = 0
			End IF
		Else
			getCustomerOneManyBookingCount = 0
		End IF
		
		'取得期限在這六天內的合約，的已使用點數
		str_sql = " 		  SELECT SUM(pused) as pused "
		str_sql = str_sql & " FROM "
		str_sql = str_sql & " (	SELECT "
		str_sql = str_sql & " 		(CASE dbo.client_attend_list.refund WHEN 1 THEN 0 ELSE ((isnull(client_attend_list.svalue,0)+0.00)*65) END) AS pused "
		str_sql = str_sql & " 	FROM  client_attend_list "
		str_sql = str_sql & " 	RIGHT OUTER JOIN "
		str_sql = str_sql & " 	(	SELECT client_sn, CONVERT(varchar, MIN(product_sdate), 111) AS sdate, CONVERT(varchar, MAX(product_edate), 111) AS edate "
		str_sql = str_sql & " 		FROM client_purchase "
		str_sql = str_sql & " 		WHERE (client_sn = @client_sn1) "
		str_sql = str_sql & " 			AND (CONVERT(varchar, product_sdate, 111) <= CONVERT(varchar, DATEADD(day,6,GETDATE()), 111)) "
		str_sql = str_sql & " 			AND  (CONVERT(varchar, GETDATE(), 111) <= CONVERT(varchar, product_edate, 111)) "
		str_sql = str_sql & " 			AND (Valid = 1) "
		str_sql = str_sql & " 		GROUP BY client_sn "
		str_sql = str_sql & " 	) AS client_purchase "
		str_sql = str_sql & " 		ON CONVERT(varchar, client_attend_list.attend_date, 111) <= client_purchase.edate "
		str_sql = str_sql & " 			AND CONVERT(varchar, client_attend_list.attend_date, 111) >= client_purchase.sdate "
		str_sql = str_sql & " 			AND client_attend_list.client_sn = client_purchase.client_sn  "
		str_sql = str_sql & " 	WHERE (client_attend_list.valid = 1) AND (client_attend_list.refund IN (0, 1)) "
		str_sql = str_sql & " ) As client_attend_list "
		arr_param_value = Array(p_str_client_sn)
        arr_attend_list_data = excuteSqlStatementRead(str_sql, arr_param_value, p_int_connect_type)
		
		If (isSelectQuerySuccess(arr_attend_list_data)) Then
			If (Ubound(arr_attend_list_data) >= 0) Then
				'總共使用的點數(含1on1 & Normal Session & View Session recording file)
				getCustomerOneManyBookingCount = getCustomerOneManyBookingCount - CInt(arr_attend_list_data(0, 0))
			End IF
		End IF
	Else
		getCustomerOneManyBookingCount = 0
	End If
	
	getCustomerOneManyBookingCount = getCustomerOneManyBookingCount / 65
End Function

%>