<!--#include virtual="/lib/include/global.inc"-->
<!--#include file="functions/functions_schedule_account_session.asp"-->
<%
    Dim bol_is_debug_mode : bol_is_debug_mode = true

    '當週結算日
    '也是下週週期初

    '每天跑的排程，計算一個週期的上課次數，若未達到限制的堂數則要扣回來

    Dim dat_sys_datetime : dat_sys_datetime = now()
    Dim dat_account_start_date, dat_account_end_date
    Dim arr_cfg_product, pindex
    Dim arr_client_purchase, purindex
    Dim arr_client_attend_list
    Dim int_account_number  '結算人數
    Dim int_account_no_deduct_number        '不需扣除未上完的堂數 人數
    Dim int_account_deduct_success_number   '需扣除未上完的堂數 成功的人數
    Dim int_account_deduct_fail_number      '需扣除未上完的堂數 失敗的人數
    Dim int_account_cycle_max_session_total '結算週期內至少(>=)要上滿多少堂
    Dim int_account_cycle_booking_count     '結算週期內的上課筆數 = 上課筆數(排除return) + 因服務端或系統端return筆數
    Dim int_deduct_points
    Dim arr_customer_session_points, bol_result
    Dim int_service_cycle_day, int_service_cycle_week, int_service_week

    '合約總點數  client_purchase.add_points(產品總點數) + client_purchase.access_points(贈送點數)
    Dim flt_total_points : flt_total_points = 0

    '來源網站: 1 = TutorABC, 2 = TutorABCJr,  3 = TutorMing 4 = NTUtorMing,  5 = Columbia,  6 = VIPABC
    Dim int_source_website : int_source_website = CONST_VIPABC_WEBSITE

    '來源程式位置
    Dim str_source_code : str_source_code = Request.ServerVariables("PATH_INFO")

    '要連到哪個DB
    Dim int_connect_db : int_connect_db = CONST_VIPABC_RW_CONN

    '開放預訂幾週的課程
    Dim int_allow_reservation_week : int_allow_reservation_week = CONST_VIPABC_ALLOW_RESERVATION_WEEK

'Response.Write  checkExtendPurchaseServiceDate(102222, int_connect_db)
'call  recordClientPurchaseLog(115557, int_connect_db)
'Response.Write isSessionBetweenAccountDate(102222, "2010022407036", int_connect_db)
'Response.Write updateClientPurchaseDeductPoints(113106, 0, int_connect_db)
'Response.Write insertClientPurchaseDeductPointsLog(Array(162579, 115557, "2010/2/18", "2010/2/24", 4, 1, 3), int_connect_db)
'Response.end

    'TODO: 做法
    '1. 只抓大陸的產品、產品的結算週期、每個結算週期內至少(>=)要上滿多少堂
    '回傳值    ：二維陣列
    '            [0]: 產品編號
    '            [1]: 適用年齡
    '            [2]: 結算週期
    '            [3]: 每個結算週期內至少(>=)要上滿多少堂 , 每週期(期末)應扣除堂數
    arr_cfg_product = getProductInfo("", int_connect_db)
    if (IsArray(arr_cfg_product)) then

        for pindex = 0 to Ubound(arr_cfg_product, 2)

            '2. 取得 合約服務日期已開始還未到期 的客戶的合約資料 且 購買的產品是定時定額
            '回傳值    ：二維陣列
            '            [0]: purchase 編號
            '            [1]: 客戶編號 client_sn
            '            [2]: 服務開始日
            '            [3]: 服務結束日
            '            [4]: contract_sn
            '            [5]: 總點數 add_points
            '            [6]: pay_money
            '            [7]: purchase_status
            '            [8]: 產品編號 product_sn
            '            [9]: 贈送點數 access_points
            '            [10]: 合約開始日 contract_sdate
            '            [11]: 合約結束日 contract_edate
            arr_client_purchase = getClientPurchaseServiceStartInfo("", arr_cfg_product(0, pindex), "", int_connect_db)
            if (IsArray(arr_client_purchase)) then
                if (bol_is_debug_mode) then
                    Response.Write "產品編號: " & arr_cfg_product(0, pindex) & "<br/>"
                end if

                if (bol_is_debug_mode) then
                    Response.Write "總共有: " & Ubound(arr_client_purchase, 2) & "人需判斷是否要結算<br>"
                end if

                '結算日 Start
                dat_account_start_date = DateValue(DateAdd("d", -(CInt(arr_cfg_product(2, pindex))-1), dat_sys_datetime))

                '結算日 END
                dat_account_end_date = DateValue(dat_sys_datetime)

                if (bol_is_debug_mode) then
                    Response.Write "結算日:" & dat_account_start_date & "(" & WeekDayName(Weekday(DateValue(dat_account_start_date))) & ") ~ " &_
                                   dat_account_end_date & "(" & WeekDayName(Weekday(DateValue(dat_account_end_date))) & ")<br>"
                end if

                int_account_number = 0
                int_account_no_deduct_number = 0
                int_account_deduct_success_number = 0
                int_account_deduct_fail_number = 0
                for purindex = 0 to Ubound(arr_client_purchase, 2)
                    int_deduct_points = 0

                    int_service_cycle_day = DateDiff("d", DateValue(arr_client_purchase(2, purindex)), DateValue(dat_sys_datetime)) mod arr_cfg_product(2, pindex)
                    int_service_cycle_week = DateDiff("d", DateValue(arr_client_purchase(2, purindex)), DateValue(dat_sys_datetime)) \ arr_cfg_product(2, pindex)

                    '取得目前是客戶合約的服務開始日第幾週
                    int_service_week = int_service_cycle_week
                    if (int_service_cycle_day > 0) then
                        int_service_week = int_service_week + 1
                    end if

                    '合約總點數  client_purchase.add_points(產品總點數) + client_purchase.access_points(贈送點數)
                    flt_total_points = arr_client_purchase(5, purindex) + arr_client_purchase(9, purindex)


                    '**** 檢查 客戶剩餘點數總表 和 大小錢包異動紀錄資料表 是否有資料 ****
                    arr_customer_session_points = getCustomerSessionPoints(arr_client_purchase(0, purindex), arr_client_purchase(1, purindex), int_connect_db)
                    if (Not IsArray(arr_customer_session_points)) then
                        '可能是首次上課，所以沒有先前的資料，理論上在帳號開通時就會有資料在customer_session_points
                        '這裡只是在做一次check
                        '初始化 客戶剩餘點數總表 和 大小錢包異動紀錄資料表
                        'TODO: 錯誤回報
						bol_result = initCustomerSessionPoints(arr_client_purchase(0, purindex), _
                                                               arr_client_purchase(1, purindex), _
                                                               flt_total_points, arr_cfg_product(3, pindex), _
                                                               2, int_source_website, str_source_code, _
                                                               arr_cfg_product(2, pindex), int_connect_db)

                    else                        

                        '3. DateDiff("d", DateValue(str_service_start_date), DateValue(dat_sys_datetime)) mod 結算週期 = 6 的人才需要結算
                        '是 當週週期末結算日
                        '也是 下週週期初的開始                
                        if (DateDiff("d", DateValue(arr_client_purchase(2, purindex)), DateValue(dat_sys_datetime)) mod arr_cfg_product(2, pindex) = 6) then

                            bol_result = countAfterCycleEnd(arr_client_purchase(0, purindex), arr_client_purchase(1, purindex), _
                                                            dat_account_end_date, flt_total_points, _
                                                            arr_cfg_product(3, pindex), arr_customer_session_points(0), _
                                                            arr_customer_session_points(2), int_source_website, _
                                                            str_source_code, int_service_week, arr_client_purchase(10, pindex), _
                                                            arr_client_purchase(11, pindex), arr_client_purchase(2, pindex), _
                                                            arr_client_purchase(3, pindex), int_allow_reservation_week, _
                                                            arr_cfg_product(3, pindex), int_connect_db)                                                                          
                        end if

                    end if
                    
                next
                if (bol_is_debug_mode) then
                    Response.Write "總共結算人數:" & int_account_number & "<br>"
                    Response.Write "不需扣除未上完的堂數:" & int_account_no_deduct_number & "<br>"
                    Response.Write "需扣除未上完的堂數 成功的人數:" & int_account_deduct_success_number & "<br>"
                    Response.Write "需扣除未上完的堂數 失敗的人數:" & int_account_deduct_fail_number & "<br>"
                    Response.Write "---------------------------------<br><br>"
                end if
            end if            
        next
    end if

    '1. 只抓大陸的產品、產品的結算週期、每個結算週期內至少(>=)要上滿多少堂
    '2. 取得 合約服務日期已開始還未到期 的客戶的合約資料
    '3. DateDiff("d", DateValue(服務開始日), DateValue(now())) mod 結算週期 = 0 的人才需要結算
    '4. 若結算週期內沒有上滿限制的堂數，要扣回來
    '   結算週期內要上完的堂數 = 產品規定的堂數 + 因客戶端原因return的筆數
    '   若要扣，則更新 client_purchase.deduct_points
 


    '1. 抓合約服務開始日期 和 產品的結算週期 和 每個結算週期內至少(>=)要上滿多少堂

    '2. DateDiff("d", DateValue(str_service_start_date), DateValue(now())) mod 結算週期 = 0 的人才需要結算

    '3. now-結算週期 to now 這週期內之前有refund過嗎 有的話判斷是客戶原因或系統原因 
    '   客戶原因的話 看幾堂 這週期要上滿的堂數限制就加幾堂
    '   應該要紀錄refund的原因 和是否被拿來使用在限制堂數上

    '4. 抓該週期內的上課次數 client_attend_list sdate bewteen now-7 to now
    '   客戶原因refund的要算嗎? 系統原因refund的要算嗎?

    '5. 判斷沒上滿的要扣回設定的堂數
%>