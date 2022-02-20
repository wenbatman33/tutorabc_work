<%
function updateClientPurchaseDeductPoints(ByVal p_int_client_purchase_sn, ByVal p_int_deduct_points, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：更新 需扣除的點數: 大陸定時定額產品，在結算週期內未使用完規定的堂數時，
'            需從合約扣除，此欄位就是記錄要扣除的點數。
'傳入參數  ：[p_int_client_purchase_sn:Integer] client_purchase.account_sn
'            [p_int_deduct_points:Integer] client_purchase.deduct_points 要扣除的點數(加上原本的)
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：true or false
'牽涉數據表：client_purchase, client_purchase_log
'備註      ：已不在使用
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/23 Created
'*****************************************************************************************/
    Dim str_sql, arr_param_value, int_sql_write_result, bol_result

    bol_result = false

    if (Not isEmptyOrNull(p_int_client_purchase_sn)) then
        str_sql = " UPDATE client_purchase SET deduct_points = @deduct_points WHERE (account_sn = @account_sn)"
        arr_param_value = Array(p_int_deduct_points, p_int_client_purchase_sn)
        int_sql_write_result = excuteSqlStatementWrite(str_sql, arr_param_value, p_int_connect_type)
        if (int_sql_write_result >= 0) then
            bol_result = true
        end if
    end if
    updateClientPurchaseDeductPoints = bol_result
end function


function insertClientPurchaseDeductPointsLog(ByVal p_arr_data, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：更新 需扣除的點數: 大陸定時定額產品，在結算週期內未使用完規定的堂數時，
'            需從合約扣除，此欄位就是記錄要扣除的點數。
'傳入參數  ：[p_arr_data:Array] 新增的資料
'               p_arr_data[0]: client_sn 客戶編號 client_basic.sn
'               p_arr_data[1]: purchase_sn client_purchase.account_sn
'               p_arr_data[2]: account_sdate 結算日 開始日期
'               p_arr_data[3]: account_edate 結算日 結束日期
'               p_arr_data[4]: booking_count 結算日內上課筆數(排除return)
'               p_arr_data[5]: return_count 因客戶端原因而return的筆數
'               p_arr_data[6]: account_cycle_max_session 結算週期內要上完的堂數 = 產品規定的堂數 + 因客戶端原因return的筆數
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：true or false
'牽涉數據表：client_purchase_deduct_points_log
'備註      ：已不在使用
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/24 Created
'*****************************************************************************************/
    Dim str_sql, int_sql_write_result, bol_result

    bol_result = false

    if (IsArray(p_arr_data)) then
        str_sql = " INSERT INTO client_purchase_deduct_points_log(client_sn, " &_
                  " purchase_sn, account_sdate, account_edate, booking_count, " &_
                  " return_count, account_cycle_max_session) " &_
                  " VALUES(@client_sn, @purchase_sn, @account_sdate, " &_ 
                  " @account_edate, @booking_count, @return_count, @account_cycle_max_session)"
        int_sql_write_result = excuteSqlStatementWrite(str_sql, p_arr_data, p_int_connect_type)
        if (int_sql_write_result >= 0) then
            bol_result = true
        end if
    end if
    insertClientPurchaseDeductPointsLog = bol_result
end function


function checkExtendPurchaseServiceDate(ByVal p_int_client_sn, ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：檢查因系統或服務原因而return的筆數，是否達到延長合約服務結束日期的條件
'            若符合則依產品的結算週期延長服務結束日
'傳入參數  ：[p_int_client_sn:Integer] 客戶編號
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：0:  不符合條件 不需延長
'            1:  符合條件並成功延長服務結束日
'            -1: 符合條件但在更新client_purchase時失敗
'            -2: 符合條件但在紀錄log client_purchase_log 時失敗
'牽涉數據表：client_session_refund, session_refund
'備註      ：已不在使用
'歷程(作者/日期/原因) ：[Ray Chien] 2010/02/23 Created
'*****************************************************************************************/
    Dim arr_session_refund, str_sql, arr_param_value, int_now_use_product_sn, int_return_count
    Dim arr_cfg_product, arr_client_purchase
    Dim obj_member_opt, int_member_result
    Dim str_service_start_date, int_sql_write_result, int_exec_result
    
    '因服務或系統return時才會呼叫此函式，這時還沒將此筆return update 進db，所以預設是1
    int_return_count = 1

    int_exec_result = 0

    Set obj_member_opt = Server.CreateObject("TutorLib.MemberOperation.ClientDataOpt")
    int_member_result = obj_member_opt.prepareData(p_int_client_sn, p_int_connect_type)
    if (int_member_result = CONST_FUNC_EXE_SUCCESS) then
        int_now_use_product_sn = obj_member_opt.getData(CONST_CONTRACT_PRODUCT_SN)  '目前使用的產品編號
        if (Not isEmptyOrNull(int_now_use_product_sn)) then
            '取得 大陸定時定額產品的結算週期、每個結算週期內至少(>=)要上滿多少堂
            arr_cfg_product = getProductInfo(int_now_use_product_sn, p_int_connect_type)
            if (IsArray(arr_cfg_product)) then
                '取得 合約服務日期已開始還未到期 的客戶的合約資料 應該只會有一筆
                arr_client_purchase = getClientPurchaseServiceStartInfo(p_int_client_sn, int_now_use_product_sn, "", p_int_connect_type)
                if (IsArray(arr_client_purchase)) then
                    str_service_start_date = Year(arr_client_purchase(2, 0)) & Right("0"&Month(arr_client_purchase(2, 0)), 2) & Right("0"&Day(arr_client_purchase(2, 0)), 2) & "00000"

                    '--合約服務開始日之後因服務或系統的原因return
                    '--取得客戶因服務或系統而return的筆數
                    str_sql = " SELECT count(*) AS total FROM client_session_refund "
                    str_sql = str_sql & " INNER JOIN session_refund ON client_session_refund.srsn = session_refund.sn "
                    str_sql = str_sql & " AND (session_refund.session_sn >= '"&str_service_start_date&"') "
                    str_sql = str_sql & " AND (session_refund.valid = 4) AND (session_refund.refund_status IN(1,3)) "
                    str_sql = str_sql & " WHERE (client_session_refund.client_sn) = @client_sn "
                    arr_param_value = Array(p_int_client_sn)    
                    arr_session_refund = excuteSqlStatementRead(str_sql, arr_param_value, p_int_connect_type)
                    if (isSelectQuerySuccess(arr_session_refund)) then
                        if (Ubound(arr_session_refund) >= 0) then
                            int_return_count = int_return_count + CInt(arr_session_refund(0, 0))
                            '若return筆數 mod 結算週期內需上滿堂數 = 1，則代表超過一個結算週期的基數 需延長服務結束日
                            if (int_return_count mod arr_cfg_product(3, 0) = 1) then
                                '紀錄log
                                if (recordClientPurchaseLog(arr_client_purchase(0,0), p_int_connect_type)) then
                                    '更新 client_purchase 
                                    str_sql = "UPDATE client_purchase SET product_edate = @product_edate WHERE account_sn = @account_sn"
                                    arr_param_value = Array(DateAdd("d", arr_cfg_product(2, 0), arr_client_purchase(3, 0)), arr_client_purchase(0,0))
                                    int_sql_write_result = excuteSqlStatementWrite(str_sql, arr_param_value, p_int_connect_type)
                                    if (int_sql_write_result >= 0) then
                                        int_exec_result = CONST_FUNC_EXE_SUCCESS
                                    else
                                        int_exec_result = -1
                                    end if
                                else
                                    int_exec_result = -2
                                end if
                            end if
                        end if
                    end if
                end if
            end if
        end if
    end if
    Set obj_member_opt = Nothing

    checkExtendPurchaseServiceDate = int_exec_result
end function


function countAfterCycleEnd(ByVal p_int_client_purchase_sn, ByVal p_int_client_sn, _
                            ByVal p_dat_account_end_date, ByVal p_flt_total_points, _
                            ByVal p_int_deduct_session, ByVal p_int_csp_sn, _
                            ByVal p_flt_available_points, ByVal p_int_source_website, _
                            ByVal p_str_source_code, ByVal p_int_service_week, _
                            ByVal p_dat_contract_start_date, ByVal p_dat_contract_end_date, _
                            ByVal p_dat_service_start_date, ByVal p_dat_service_end_date, _
                            ByVal p_int_allow_reservation_week, ByVal p_int_account_cycle, _
                            ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：週期末結算，也是下週的週期初
'            每週期期初計算：每個週期期初，先比對該週小錢包堂數是否大於等於每週應扣除堂數(eg. 當第15週起始時，比對第15週之小錢包堂數)
'            1.	若小錢包堂數=0，則從大錢包撥每週應扣除堂數到小錢包。
'            2.	若小錢包堂數 < 每週應扣除堂數，則從大錢包撥不足的堂數到小錢包。
'            3.	若小錢包堂數 >= 每週應扣除堂數，則不從大錢包中扣撥每週應扣除堂數到小錢包。
'傳入參數  ：[p_int_client_purchase_sn:Integer] client_purchase.account_sn
'            [p_int_client_sn:Integer] 客戶編號 client_basic.sn
'            [p_dat_account_end_date:Integer] 週期末結算日
'            [p_flt_total_points:Integer] 合約的總點數
'            [p_int_deduct_session:Integer] 每週期(期末)應扣除堂數
'            [p_int_csp_sn:Integer] customer_session_points.csp_sn
'            [p_flt_available_points:Float] 大錢包 : 可用點數  = 已使用堂數(client_attend_list.valid=1) + return - 錄影檔使用
'            [p_int_source_website:Integer] 來源網站:
'               1 = TutorABC, 2 = TutorABCJr,  3 = TutorMing 
'               4 = NTUtorMing,  5 = Columbia,  6 = VIPABC, 7=TutorABCIMS, 
'               8 = TUTORABCJR_IMS, 9=TUTORMING_IMS, 10 = NTUTORMING_IMS,  11 = CLUE_IMS, 12 = VIPABC_IMS
'            [p_str_source_code:String] 來源程式位置
'            [p_int_service_week:String] 目前是客戶合約的服務開始日第幾週
'            [p_dat_contract_start_date:String] 合約開始日
'            [p_dat_contract_end_date:String] 合約結束日
'            [p_dat_service_start_date:String] 服務開始日
'            [p_dat_service_end_date:String] 服務結束日
'            [p_int_allow_reservation_week:String] 開放預訂幾週的課程
'            [p_int_account_cycle:Integer] 結算週期
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：true or false
'牽涉數據表：client_purchase, client_purchase_log
'備註      ：
'歷程(作者/日期/原因) ：[Ray Chien] 2010/04/06 Created
'*****************************************************************************************/
    Dim int_client_purchase_sn, int_client_sn, arr_client_attend_list, arr_scycle_data, arr_update_data
    Dim int_csp_sn, dat_next_account_start_date, dat_next_account_end_date
    Dim int_service_week, int_source_website, str_source_code, arr_log_data, bol_result, str_note
    

    '上課次數(排除return)
    Dim int_booking_count : int_booking_count = 0
                            
    '小錢包 : 本週剩餘可用的點數 
    Dim flt_week_points : flt_week_points = 0

    '大錢包 : 可用點數  = 已使用堂數(client_attend_list.valid=1) + return - 錄影檔使用
    Dim flt_available_points : flt_available_points = p_flt_available_points

    '合約總點數  client_purchase.add_points(產品總點數) + client_purchase.access_points(贈送點數)
    Dim flt_total_points : flt_total_points = p_flt_total_points

    '每週期(期末)應扣除堂數
    Dim int_deduct_session : int_deduct_session = CInt(p_int_deduct_session)

    int_client_sn = p_int_client_sn
    int_client_purchase_sn = p_int_client_purchase_sn
    int_csp_sn = p_int_csp_sn
    int_service_week = p_int_service_week
    int_source_website = p_int_source_website
    str_source_code = p_str_source_code

    '下個週期初 開始日 = 週期末結算日 + 1天
    dat_next_account_start_date = DateValue(DateAdd("d", 1, p_dat_account_end_date))

    '下個週期初 結束日
    dat_next_account_end_date = DateValue(DateAdd("d", 6, dat_next_account_start_date))
    
    '**** 週期末試算服務結束日 START ****
    Call calculateServiceEndDate(int_client_purchase_sn, int_client_sn, flt_available_points, _
                                 dat_next_account_start_date, int_deduct_session, _
                                 p_int_allow_reservation_week, p_dat_contract_start_date, p_dat_contract_end_date, _
                                 p_dat_service_start_date, p_dat_service_end_date, p_int_account_cycle, _
                                 int_source_website, str_source_code, p_int_connect_type)
    '**** 週期末試算服務結束日 END ****



    '**** 週期初撥堂數計算 START ****

    '取出下個週期內的上課筆數(排除return)
    arr_client_attend_list = getCustomerBookingCount(int_client_sn, dat_next_account_start_date, dat_next_account_end_date, p_int_connect_type)            
    int_booking_count = CInt(arr_client_attend_list(0))

    '一堂課都沒預訂，代表小錢包堂數=0，則從大錢包撥每週應扣除堂數到小錢包
    if (int_booking_count = 0) then
        flt_week_points = int_deduct_session * CONST_SESSION_TRANSFORM_POINT

    '預訂筆數 >= 每週期(期末)應扣除堂數，代表小錢包堂數 >= 每週應扣除堂數，則不從大錢包中扣撥每週應扣除堂數到小錢包
    elseif (int_booking_count >= int_deduct_session) then
        flt_week_points = 0

    '預訂筆數 < 每週期(期末)應扣除堂數，代表小錢包堂數 < 每週應扣除堂數，則從大錢包撥不足的堂數到小錢包
    elseif (int_booking_count < int_deduct_session) then
        flt_week_points = (int_deduct_session - int_booking_count) * CONST_SESSION_TRANSFORM_POINT
    end if

    '從大錢包撥放堂數到小錢包
    flt_available_points = Round(flt_available_points - flt_week_points, 2)

'               p_arr_data[0] csp_sn: customer_session_points.csp_sn
'               p_arr_data[1] csh_client_sn: client_sn 客戶編號 client_basic.sn
'               p_arr_data[2] csh_purchase_sn: client_purchase.account_sn
'               p_arr_data[3] csh_total_points: 大錢包 = 已使用堂數(client_attend_list.valid=1) + return - 錄影檔使用 - 當週小錢包
'               p_arr_data[4] csh_week_points: 小錢包 = 本週剩餘可用的點數
'               p_arr_data[5] csh_type: 新增此筆資料的類型 
'               p_arr_data[6] csh_booking_sn: 訂課TABLE的流水號 client_attend_list.sn若是系統排程insert資料的話，不會有值 
'               p_arr_data[7] csh_staff_sn: 員工編號 staff_basic.sn  從IMS訂課才會有 
'               p_arr_data[8] csh_note: 備註 
'               p_arr_data[9] csh_cycle_week: 目前是客戶合約的服務開始日第幾週 
'               p_arr_data[10] csh_account_sdate: 週期初開始日期 
'               p_arr_data[11] csh_account_edate: 週期末結束日期 (結算日) 
'               p_arr_data[12] csh_source_website: 來源網站: 1 = TutorABC, 2 = TutorABCJr,  3 = TutorMing 4 = NTUtorMing,  5 = Columbia,  6 = VIPABC 
'               p_arr_data[13] csh_source_code: 來源程式位置 
    str_note = "週期初撥堂數計算，從大錢包撥" & flt_week_points & "堂到小錢包"
    arr_log_data = Array(int_csp_sn, int_client_sn, int_client_purchase_sn, _
                         flt_available_points, flt_week_points, 2, _
                         "", "", str_note,  _
                         int_service_week, dat_next_account_start_date, dat_next_account_end_date, _
                         int_source_website, str_source_code)
    '新增一筆資料到 大小錢包異動紀錄資料表 總剩餘堂數和每週使用堂數異動紀錄
    bol_result = insertCustomerSessionPointsLog(arr_log_data, p_int_connect_type)

    if (bol_result) then
        arr_update_data = Array(int_csp_sn, int_client_sn, int_client_purchase_sn, _
                                flt_available_points, flt_week_points, flt_week_points)
        bol_result = updateCustomerSessionPoints(arr_update_data, p_int_connect_type)
    end if
    '**** 週期初撥堂數計算 END ****

    countAfterCycleEnd = bol_result
end function


function calculateServiceEndDate(ByVal p_int_client_purchase_sn, ByVal p_int_client_sn, ByVal p_flt_available_points, _
                                 ByVal p_dat_next_account_start_date, ByVal p_int_deduct_session, _
                                 ByVal p_int_calculate_number, ByVal p_dat_contract_start_date, _
                                 ByVal p_dat_contract_end_date, ByVal p_dat_service_start_date, _
                                 ByVal p_dat_service_end_date, ByVal p_int_account_cycle, _
                                 ByVal p_int_source_website, ByVal p_str_source_code, _
                                 ByVal p_int_connect_type)
'/****************************************************************************************
'描述      ：試算課時數結束日
'傳入參數  ：[p_int_client_purchase_sn:Integer] client_purchase.account_sn
'            [p_int_client_sn:Integer] 客戶編號 client_basic.sn
'            [p_flt_available_points:Float] 大錢包
'            [p_dat_next_account_start_date:DateTime] 下個週期初 開始日
'            [p_int_deduct_session:Integer] 每週期(期末)應扣除堂數
'            [p_int_calculate_number:Integer] 試算幾週
'            [p_dat_contract_start_date:String] 合約開始日
'            [p_dat_contract_end_date:String] 合約結束日
'            [p_dat_service_start_date:String] 服務開始日
'            [p_dat_service_end_date:String] 服務結束日
'            [p_int_account_cycle:Integer] 結算週期
'            [p_int_source_website:Integer] 來源網站:
'               1 = TutorABC, 2 = TutorABCJr,  3 = TutorMing 
'               4 = NTUtorMing,  5 = Columbia,  6 = VIPABC, 7=TutorABCIMS, 
'               8 = TUTORABCJR_IMS, 9=TUTORMING_IMS, 10 = NTUTORMING_IMS,  11 = CLUE_IMS, 12 = VIPABC_IMS
'            [p_str_source_code:String] 來源程式位置
'            [p_int_connect_type:Integer] 連線的主機型態
'回傳值    ：成功: 試算後的日期 , 失敗: 空字串
'牽涉數據表：
'備註      ：若「試算課時數結束日」大於「合約結束日」，則「服務結束日」 = 「試算課時數結束日」
'            試算課時數結束日公式 (單位:週，無條件進位至整數) =
'            未來4周的第四週週數 + 
'            ((總剩餘堂數 –試算每周總應扣除堂數) / 產品定義的每周應扣除堂數)
'歷程(作者/日期/原因) ：[Ray Chien] 2010/04/09 Created
'*****************************************************************************************/
    Dim i, arr_client_attend_list, int_client_sn, int_client_purchase_sn
    Dim dat_next_account_start_date, dat_next_account_end_date
    Dim int_contract_week, int_contract_cycle_day, int_contract_cycle_week
    Dim str_sql, arr_param_value, int_sql_write_result
    Dim dat_new_service_end_date, int_source_website, str_source_code

    int_client_sn = p_int_client_sn
    int_client_purchase_sn = p_int_client_purchase_sn
    int_source_website = p_int_source_website
    str_source_code = p_str_source_code

    Dim bol_calculate_end : bol_calculate_end = false

    '試算試算課時數結束日(週)
    Dim int_calculate_week : int_calculate_week = 0

    '可用的總堂數
    Dim int_available_session : int_available_session = p_flt_available_points \ CONST_SESSION_TRANSFORM_POINT

    '試算每周總應扣除堂數 總和
    Dim int_total_deduct : int_total_deduct = 0

    '訂課筆數
    Dim int_booking_count : int_booking_count = 0

    '每週期(期末)應扣除堂數
    Dim int_deduct_session : int_deduct_session = CInt(p_int_deduct_session)

    '下個週期初 開始日
    dat_next_account_start_date = p_dat_next_account_start_date

    '下個週期初 結束日
    dat_next_account_end_date = DateValue(DateAdd("d", 6, dat_next_account_start_date))

    Response.Write "<br>client_purchase.account_sn=" & int_client_purchase_sn & "<br>"
    Response.Write "<br>合約結束日=" & p_dat_contract_end_date & "<br>"
    Response.Write "<br>服務結束日=" & p_dat_service_end_date & "<br>"
    Response.Write "<br>試算幾週=" & p_int_calculate_number & "<br>"
    Response.Write "<br>每週期(期末)應扣除堂數=" & p_int_deduct_session & "<br>"

    '取得客戶合約開始日 到 合約結束日共幾週
    int_contract_cycle_day = DateDiff("d", DateValue(p_dat_contract_start_date), DateValue(p_dat_contract_end_date)) mod arr_cfg_product(2, pindex)
    int_contract_cycle_week = DateDiff("d", DateValue(p_dat_contract_start_date), DateValue(p_dat_contract_end_date)) \ arr_cfg_product(2, pindex)
    int_contract_week = int_contract_cycle_week
    if (int_contract_cycle_day > 0) then
        int_contract_week = int_contract_week + 1
    end if

    for i = 1 to p_int_calculate_number-1

    'Response.Write "<br>下個週期初 開始日=" & dat_next_account_start_date & "<br>"
    'Response.Write "<br>下個週期初 結束日=" & dat_next_account_end_date & "<br>"

        '判斷 若下週的結束日 <= 結束日期，則不需在運算其他週
        if (DateDiff("d", DateValue(dat_next_account_end_date), DateValue(p_dat_service_end_date)) <= 0) then            
            dat_next_account_end_date = p_dat_service_end_date
            bol_calculate_end = true
        end if

        '取出下個週期內的上課筆數(排除return)
        arr_client_attend_list = getCustomerBookingCount(int_client_sn, dat_next_account_start_date, dat_next_account_end_date, p_int_connect_type)            
        int_booking_count = CInt(arr_client_attend_list(0))
        '未來週 預訂的課程 不超過每週期(期末)應扣除堂數，相減後的數字就是試算每周總應扣除堂數
        if (int_booking_count < int_deduct_session) then
            int_total_deduct = int_total_deduct + (int_deduct_session - int_booking_count)
        end if

        if (bol_calculate_end) then
            exit for
        else
            '下個週期初 開始日
            dat_next_account_start_date = DateValue(DateAdd("d", 1, dat_next_account_end_date))

            '下個週期初 結束日
            dat_next_account_end_date = DateValue(DateAdd("d", 6, dat_next_account_start_date))
        end if
    next

    '((總剩餘堂數 –試算每周總應扣除堂數) / 產品定義的每周應扣除堂數)    
    int_calculate_week = (int_available_session - int_total_deduct) \ int_deduct_session
    if (((int_available_session - int_total_deduct) mod int_deduct_session) > 0) then
        int_calculate_week = int_calculate_week + 1
    end if

    '判斷若「試算課時數結束日」大於「合約結束日」，則「服務結束日」 = 「試算課時數結束日」
    if (CInt(int_calculate_week) > CInt(int_contract_week)) then
        '新的服務結束日 = 合約結束日 + 延長的週數
        dat_new_service_end_date = DateValue(DateAdd("d", ((int_calculate_week - int_contract_week) * 7), p_dat_contract_end_date))

        '更新服務結束日
        str_sql = " UPDATE client_purchase SET product_edate = @product_edate WHERE account_sn = @account_sn "
        arr_param_value = Array(dat_new_service_end_date, int_client_purchase_sn)
        int_sql_write_result = excuteSqlStatementWrite(str_sql, arr_param_value, p_int_connect_type)
        if (int_sql_write_result >= 0) then
            bol_result = true
            '新增一筆到 合約紀錄 log 表
            Call insertClientPurchaseLog(int_client_purchase_sn, "", int_source_website, str_source_code, p_int_connect_type)

        end if

        Response.Write "<br>新的服務結束日=" & dat_new_service_end_date & "<br>"

        Response.Write "<br>試算每周總應扣除堂數=" & int_calculate_week & "<br>"


        Response.Write "<br>客戶合約的服務開始日 到 服務結束日共幾週=" & int_contract_week & "<br>"

    end if
end function

%>