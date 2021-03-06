var g_obj_reservation_class_ctrl = {

    dat_sys_datetime: new Date(g_obj_class_data.str_now_sys_datetime.substr(0, 4), parseInt(g_obj_class_data.str_now_sys_datetime.substr(4, 2), 10) - 1, g_obj_class_data.str_now_sys_datetime.substr(6, 2), g_obj_class_data.str_now_sys_datetime.substr(8, 2), g_obj_class_data.str_now_sys_datetime.substr(10, 2), g_obj_class_data.str_now_sys_datetime.substr(12, 2)),

    // 預定或要取消的大會堂編號
    strReservationLobbySn: "",

    // 已預定課程 總共要扣的課時數
    flt_total_order_session: 0.0,

    // 取得客戶總剩餘堂數
    str_client_total_session: g_obj_class_data.str_client_total_session,

    // 課程類型 一對一 OR 一對多 OR 大會堂 OR 一對三
    arr_class_type: {
        "normal": g_obj_class_data.arr_class_type.normal,
        "vip": g_obj_class_data.arr_class_type.vip,
        "lobby": g_obj_class_data.arr_class_type.lobby,
        "one_on_three": g_obj_class_data.arr_class_type.one_on_three
    },

    // 錯誤代碼
    arr_err_code: {
        "same_class": -1,                  // 新增到同一堂課
        "class_no_open_in_4_hour": -2,     // 選到沒有開課的諮詢時間 (例如四小時內)
        "class_no_open_in_24_hour": -3,    // 24小時內臨時訂課僅開放一對多
        "wrong_class_dt": -4,              // 錯誤的訂課日期和時間
        "class_no_open_in_5_min": -5,      // 大會堂 只開放五分鐘前可以訂課
        "not_enough_session": -9,          // 合约剩余课时数不足，请至「学习记录」查询详细资讯。
        "select_other_class_date": -99     // 请选择其他上课日期
    },

    // ajax or post or 其他if判斷式要用到的 的頁面URL
    arr_page_url: {
        "ajax_class_info": "/program/member/reservation_class/include/ajax_class_info.asp",                    // ajax 取值用
        "normal_and_vip_near_order": "/program/member/reservation_class/normal_and_vip_near_order.asp",        // if 判斷式用
        "normal_and_vip_near_cancel": "/program/member/reservation_class/normal_and_vip_near_cancel.asp",      // if 判斷式用
        "reservation_class_exe": "/program/member/reservation_class/reservation_class_exe.asp",                // 點擊 近期預定課程頁面 "確認訂課" 按鈕 要POST FROM的頁面
        "ajax_cancel_class": "/program/member/reservation_class/include/ajax_cancel_class.asp"                 // ajax 取消課程
    },

    // 取得可供客戶訂課的時間
    getHtmlSelectClassTime: function (event) {
        var str_class_type = $("#sel_class_type").val();
        var str_class_date = $("#txt_class_date").val();
        var int_clear_old_date = event.data.clear_old_date;

        if (int_clear_old_date === 1) {
            $("#hdn_prev_select_class_date").val("");
        }

        if (str_class_date.trim() != "" && str_class_date.trim() != g_obj_class_data.arr_msg.no_sel_class_date && str_class_date.trim() != $("#hdn_prev_select_class_date").val()) {
            // 儲存已選擇的訂課日期
            $("#hdn_prev_select_class_date").val(str_class_date);

            // 載入可供訂課的諮詢時間
            $("#li_reservation_class_time").load(g_obj_reservation_class_ctrl.arr_page_url.ajax_class_info, { ndate: new Date(), act: 4, rcdate: str_class_date, rctype: str_class_type });
        }
    },

    // Button 新增 事件
    addNewReservationClass: function () {
        var str_class_type = $("#sel_class_type").val();
        var str_class_date = $("#txt_class_date").val();
        var str_class_time = $("#sel_class_time").val();
        var str_now_browse_program_name = $("#hdn_now_browse_program_name").val();
        var int_add_result;
        var str_contract_start_date_yyyymmdd = g_obj_class_data.str_service_start_date.replace(new RegExp("/", "gm"), ""); //  服務開始日 YYYYMMDDHH
        var str_contract_end_date_yyyymmdd = g_obj_class_data.str_service_end_date.replace(new RegExp("/", "gm"), "");   //  服務結束日 YYYYMMDDHH
        var str_class_date_yyyymmdd = str_class_date.replace(new RegExp("/", "gm"), ""); //  訂課日期 YYYYMMDDHH

        var dat_contract_start_date = new Date(str_contract_start_date_yyyymmdd.substr(0, 4),
                                               parseInt(str_contract_start_date_yyyymmdd.substr(4, 2), 10) - 1,
                                               str_contract_start_date_yyyymmdd.substr(6, 2));

        var dat_contract_end_date = new Date(str_contract_end_date_yyyymmdd.substr(0, 4),
                                               parseInt(str_contract_end_date_yyyymmdd.substr(4, 2), 10) - 1,
                                               str_contract_end_date_yyyymmdd.substr(6, 2));

        var dat_class_date = new Date(str_class_date_yyyymmdd.substr(0, 4),
                                               parseInt(str_class_date_yyyymmdd.substr(4, 2), 10) - 1,
                                               str_class_date_yyyymmdd.substr(6, 2));

        // 合約已到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
        if (str_contract_end_date_yyyymmdd == "") {
            // 合約已到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
            alert(g_obj_class_data.arr_msg.contract_edate_invalid_0 + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
            return false;
        }

        // 檢查是否選到合約開始日之前的日期
        if (dat_contract_start_date > dat_class_date) {
            // 請預訂YYYY/MM/DD(合約開始)之後的課程。
            alert(g_obj_class_data.arr_msg.reservation_after_contract_start);
            return false;
        }

        // 檢查是否選到合約結束日之後的日期
        if (dat_class_date > dat_contract_end_date) {
            // 合約于YYYY/MM/DD到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
            alert(g_obj_class_data.arr_msg.contract_edate_invalid_1 + g_obj_class_data.str_service_end_date + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
            return false;
        }

        // 檢查合約是否未到期
        if (!g_obj_class_data.bol_client_contract_edate_valid) {
            // 合約于YYYY/MM/DD到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
            alert(g_obj_class_data.arr_msg.contract_edate_invalid_1 + g_obj_class_data.str_service_end_date + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
            return false;
        }

        // 總剩餘堂數 = 剩餘堂數
        var flt_available_points = parseFloat(g_obj_class_data.str_client_total_session);

        // 檢查是否還有足夠的堂數
        // 判斷購買合約的產品類型 1:一般產品 2:定時定額
        // 若是定時定額的話且預訂當週的課程，必須加上當週剩餘堂數
        if (g_obj_class_data.str_client_purchase_product_type == "2") {
            // 預訂本周的課
            if (g_obj_reservation_class_ctrl.isSessionInThisWeek(str_class_date, g_obj_class_data.dat_now_week_account_start_date, g_obj_class_data.dat_now_week_account_end_date)) {
                // 總剩餘堂數 = 剩餘堂數 + 當週剩餘堂數
                flt_available_points = parseFloat(g_obj_class_data.str_client_total_session) + parseFloat(g_obj_class_data.str_client_week_session);
            }
        }

        // 檢查是否還有足夠的堂數
        if (flt_available_points <= 0) {
            // 合约剩余课时数不足，请至「学习记录」查询详细资讯。
            alert(g_obj_class_data.arr_msg.contract_edate_invalid_4 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
            return false;
        }

        // 檢查課程類型
        if (str_class_type.trim() === "") {
            alert(g_obj_class_data.arr_msg.no_sel_class_type);
            $("#sel_class_type").focus();
            return false;
        }

        // 檢查是否符合日期格式
        if (str_class_date.trim() === "" || str_class_date.trim() === g_obj_class_data.arr_msg.no_sel_class_date) {
            alert(g_obj_class_data.arr_msg.no_sel_class_date);
            $("#txt_class_date").focus();
            return false;
        }

        // 檢查諮詢時間是否有開課
        if (str_class_time.trim() === "") {
            alert(g_obj_class_data.arr_msg.no_sel_class_time);
            $("#sel_class_time").focus();
            return false;
        }
        // 24小時內臨時訂課僅開放一對多
        else if (str_class_time.trim() === g_obj_reservation_class_ctrl.arr_err_code.class_no_open_in_24_hour.toString()) {
            alert(g_obj_class_data.arr_msg.class_no_open_in_24_hour);
            return false;
        }
        // 選到沒有開課的諮詢時間 (例如四小時內)
        else if (str_class_time.trim() === g_obj_reservation_class_ctrl.arr_err_code.class_no_open_in_4_hour.toString()) {
            // 一對三訂課顯示
            if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three) {
                alert(g_obj_class_data.arr_msg_class_no_order_in_hour.one_on_three);
            }
            // 一對一訂課顯示
            else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
                alert(g_obj_class_data.arr_msg_class_no_order_in_hour.vip);
            }
            // 大會堂訂課顯示
            else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
                alert(g_obj_class_data.arr_msg_class_no_order_in_hour.lobby);
            }
            return false;
        }
        // 選到 人事部門規定之不開放上課之日期 或 選到的日期並無可供選擇的時段
        else if (str_class_time.trim() === g_obj_reservation_class_ctrl.arr_err_code.select_other_class_date.toString()) {
            alert(g_obj_class_data.arr_msg.select_other_class_date);
            return false;
        }

        // 若是首次按下則顯示提醒訊息
        if ($("#hdn_show_reservation_class_remind_msg").val() === "") {
            alert(g_obj_class_data.arr_msg.normal_remind);
        }
        $("#hdn_show_reservation_class_remind_msg").val("1");

        // 從預訂課程頁面訂課後轉址到 此次預訂課程頁面
        if (str_now_browse_program_name !== g_obj_reservation_class_ctrl.arr_page_url.normal_and_vip_near_order && str_now_browse_program_name !== g_obj_reservation_class_ctrl.arr_page_url.normal_and_vip_near_cancel) {
            $("#form_reservation_class_module").submit();
            //window.location.href = "normal_and_vip_near_order";
        }
        // 新增一筆訂課資訊
        else {
            int_add_result = g_obj_reservation_class_ctrl.addNewClassInfo($("#sel_class_type").val(), $("#txt_class_date").val(), $("#sel_class_time").val());
            // 新增失敗
            if (parseFloat(int_add_result) < 0) {
                // 顯示錯誤訊息
                alert(g_obj_reservation_class_ctrl.getReservationFailMessage(int_add_result.toString()));
                return false;
            }
        }
        return true;
    },

    // 新增一筆訂課到列表
    addNewClassInfo: function (p_str_class_type, p_str_class_date, p_str_class_time, p_str_con_ename, p_str_topic, p_str_lobby_session_sn) {
        var str_class_type = p_str_class_type;
        var str_class_date = p_str_class_date;
        var str_order_class_datetime = str_class_date.replace(new RegExp("/", "gm"), "") + p_str_class_time; //  訂課日期時間 YYYYMMDDHH
        var str_append_class_info = "", str_img_title = "";
        var flt_subtract_session = -1.0;    // 要扣的堂數
        var str_ball_class_name = "";
        var str_reservation_class_count = $("#hdn_reservation_class_count").val();
        var int_client_total_session = g_obj_reservation_class_ctrl.str_client_total_session;
        var int_client_week_session = g_obj_reservation_class_ctrl.str_client_week_session;
        var str_con_ename = p_str_con_ename;
        if (str_con_ename != null) {
            str_con_ename = str_con_ename.replace(" ", "<br/>");
        }

        // 檢查選到的時間和時段是否新增過
        if ($("#li_" + str_order_class_datetime).length > 0) {
            return g_obj_reservation_class_ctrl.arr_err_code.same_class;
        }
        else {
            // 設定選取到的大會堂編號
            if (null != p_str_lobby_session_sn) {
                g_obj_reservation_class_ctrl.strReservationLobbySn = p_str_lobby_session_sn;
            }
            flt_subtract_session = g_obj_reservation_class_ctrl.getSubtractSession(str_order_class_datetime, str_class_type);

            // 檢查剩餘課時數是否足夠
            if (flt_subtract_session >= 0 && ((int_client_total_session + int_client_week_session) - (g_obj_reservation_class_ctrl.flt_total_order_session + flt_subtract_session)).toFixed(2) < 0) {
                return g_obj_reservation_class_ctrl.arr_err_code.not_enough_session;
            }
            else if (flt_subtract_session >= 0) {
                // Get 一對三或一對一訂課顯示 紅色或綠色或藍色圓點
                str_ball_class_name = g_obj_reservation_class_ctrl.getClassImageFileName(str_class_type);

                // 一對三訂課顯示 img title
                if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three) {
                    str_img_title = g_obj_class_data.arr_msg.font_one_on_three_class;
                }
                // 一對一訂課顯示 img title
                else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
                    str_img_title = g_obj_class_data.arr_msg.font_vip_class;
                }
                // 大會堂訂課顯示 img title
                else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
                    str_img_title = g_obj_class_data.arr_msg.font_lobby_class;
                }

                // 一對一和一對三
                if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three || str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
                    str_append_class_info = "<li id='li_" + str_order_class_datetime + "'><h1 id='h1_" + str_order_class_datetime + "'><font id='font_" + str_order_class_datetime + "'></font><img width='23' height='16' border='0' src='/images/reservation_class/" + str_ball_class_name + "' title='" + str_img_title + "'/></h1>" +
                                                    "<h2>" + str_class_date + "</h2>" +
                                                    "<h3>" + p_str_class_time + ":30</h3>" +
                                                    "<h4 id='subtract_session_" + str_order_class_datetime + "'><span class='color_1'>" + flt_subtract_session + "</span>" + g_obj_class_data.arr_msg.font_list_1 + "</h4>" +
                                                    "<h6 id='h6_order_button_" + str_order_class_datetime + "'><input type='button' id='btn_" + str_order_class_datetime + "' value='-" + g_obj_class_data.arr_msg.btn_cancel_class + "' class='btn_cancel' /></h6>" +
                                                    "<input type='hidden' id='hdn_enabled_" + str_order_class_datetime + "' value='1' /><input type='hidden' id='hdn_class_type_" + str_order_class_datetime + "' value='" + str_class_type + "' />" +
                                                    "<input type='hidden' id='hdn_no_cancel_in_four_" + str_order_class_datetime + "' value='0' /></li>";
                }
                // 大會堂
                else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
                    str_append_class_info = "<li id='li_" + str_order_class_datetime + "'><h1 id='h1_" + str_order_class_datetime + "'><font id='font_" + str_order_class_datetime + "'></font><img width='23' height='16' border='0' src='/images/reservation_class/" + str_ball_class_name + "' title='" + str_img_title + "'/></h1>" +
                                                    "<h2>" + str_class_date + "</h2>" +
                                                    "<h3>" + p_str_class_time + ":30</h3>" +
                                                    "<h4><span id='span_con_ename_" + str_order_class_datetime + "'>" + str_con_ename + "</span></h4>" +
                                                    "<h5>" + p_str_topic + "</h5>" +
                                                    "<h2 id='subtract_session_" + str_order_class_datetime + "' class='shit'><span class='color_1'>" + flt_subtract_session + "</span>" + g_obj_class_data.arr_msg.font_list_1 + "</h2>" +
                                                    "<h6 id='h6_order_button_" + str_order_class_datetime + "'><input type='button' id='btn_" + str_order_class_datetime + "' value='-" + g_obj_class_data.arr_msg.btn_cancel_class + "' class='btn_cancel' /></h6>" +
                                                    "<input type='hidden' id='hdn_enabled_" + str_order_class_datetime + "' value='1' /><input type='hidden' id='hdn_class_type_" + str_order_class_datetime + "' value='" + str_class_type + "' />" +
                                                    "<input type='hidden' id='hdn_lobby_session_sn_" + str_order_class_datetime + "' value='" + p_str_lobby_session_sn + "' />" +
                                                    "<input type='hidden' id='hdn_no_cancel_in_four_" + str_order_class_datetime + "' value='0' /></li>";
                }

                $("#ul_reservation_class_info").append(str_append_class_info);

                // 註冊點擊事件 取消課程按鈕
                $("#btn_" + str_order_class_datetime).bind("click", { class_dt: str_order_class_datetime, class_type: str_class_type }, g_obj_reservation_class_ctrl.cancelClassTemporary);

                // 訂課筆數 + 1
                (str_reservation_class_count.trim() != "") ? $("#hdn_reservation_class_count").val((parseInt(str_reservation_class_count, 10) + 1).toString()) : $("#hdn_reservation_class_count").val("1");

                // 更新 預定課程筆數、總預計扣除課時數、目前剩餘的堂數
                g_obj_reservation_class_ctrl.updateReservationClassInfo(flt_subtract_session, "+");

                if ($("#ul_reservation_class_info").css("display") == "none") {
                    if (document.getElementById("div_near_order_class") != null && $("#hdn_reservation_class_count").val() !== "0") {
                        // 一對一和一對三
                        if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three || str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
                            $("#div_near_order_class").text(g_obj_class_data.arr_msg.a_have_order_class);
                            $("#div_no_reservation_class_for_normal").hide();
                            $("#div_no_reservation_class_remind_msg").hide(); // 隱藏 找不到或沒有資料的共用訊息框
                        }
                        // 大會堂
                        else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
                            $("#div_near_order_class").text(g_obj_class_data.arr_msg.a_have_order_class);
                            $("#div_no_reservation_class_for_lobby").hide();
                            $("#div_no_reservation_class_remind_msg").hide(); // 隱藏 找不到或沒有資料的共用訊息框
                        }
                    }
                    //阿捨移除
                    //$("#ul_reservation_class_info").show("fast");     // 顯示 訂課資訊列表
                    //$("#div_reservation_class_result").show("fast");  // 顯示 訂課結果及確認送出按鈕
                    $("#remind_contract_start_msg").hide();           // 隱藏 提示合約開始日訊息
                }

                // 設定 Timer 每隔一分鐘 檢查訂的課程是否在四小時內
                g_obj_reservation_class_ctrl.checkClassNoCancel();
                setTimeout(g_obj_reservation_class_ctrl.checkClassNoCancel, 60000);

                // 更新 已預定課程 總共要扣的課時數
                g_obj_reservation_class_ctrl.flt_total_order_session += flt_subtract_session;

                return 1;
            }
            else {
                return flt_subtract_session;
            }
        }
    },

    // 取消訂課事件
    cancelClassTemporary: function (event) {
        var str_class_type = event.data.class_type;
        var str_order_class_datetime = event.data.class_dt;
        var str_reservation_class_count = $("#hdn_reservation_class_count").val();

        // 加上 class "no"
        $("#li_" + str_order_class_datetime).addClass("no");

        // 加上"已取消"字樣
        $("#font_" + str_order_class_datetime).text(g_obj_class_data.arr_msg.font_already_cancel_class);

        // 清空預計扣除課時數
        $("#subtract_session_" + str_order_class_datetime).empty();
        $("#subtract_session_" + str_order_class_datetime).append("<font>&nbsp;</font>");

        // 替換 "取消此課程" button 為 "重訂此課程"
        $("#btn_" + str_order_class_datetime).val("+" + g_obj_class_data.arr_msg.btn_reorder_class).unbind("click").bind("click", { class_dt: str_order_class_datetime, class_type: str_class_type }, g_obj_reservation_class_ctrl.ajaxRegetSubtractSession);

        // Enabled 此筆課程設定 (按確認送出後是以此來判斷是否有效 1:有效 0:已取消)
        $("#hdn_enabled_" + str_order_class_datetime).val("0");

        // 訂課筆數 - 1
        (str_reservation_class_count.trim() != "") ? $("#hdn_reservation_class_count").val((parseInt(str_reservation_class_count, 10) - 1).toString()) : $("#hdn_reservation_class_count").val("1");

        // 更新 預定課程筆數、總預計扣除課時數、目前剩餘的堂數
        g_obj_reservation_class_ctrl.updateReservationClassInfo(g_obj_reservation_class_ctrl.getSubtractSession(str_order_class_datetime, str_class_type), "-");
    },

    // 重新訂課事件
    reOrderClass: function (p_str_order_class_datetime, p_flt_subtract_session, p_str_class_type) {
        var str_order_class_datetime = p_str_order_class_datetime;  // 訂課日期時間 YYYYMMDDHHmm
        var flt_subtract_session = p_flt_subtract_session;          // 重新取得需扣的課時數
        var str_reservation_class_count = $("#hdn_reservation_class_count").val();

        // 刪除 class "no"
        $("#li_" + str_order_class_datetime).removeClass("no");

        // 刪除"已取消"字樣
        $("#font_" + str_order_class_datetime).text("");

        // 重新顯示預計扣除課時數
        $("#subtract_session_" + str_order_class_datetime).append("<span class='color_1'>" + flt_subtract_session + "</span>" + g_obj_class_data.arr_msg.font_list_1);

        // 替換 "取消此課程" button 為 "重訂此課程"
        $("#btn_" + str_order_class_datetime).val("-" + g_obj_class_data.arr_msg.btn_cancel_class).unbind("click").bind("click", { class_dt: str_order_class_datetime, class_type: p_str_class_type }, g_obj_reservation_class_ctrl.cancelClassTemporary);

        // Disabled 此筆課程設定 (按確認送出後是以此來判斷是否有效 1:有效 0:已取消)
        $("#hdn_enabled_" + str_order_class_datetime).val("1");

        // 訂課筆數 + 1
        (str_reservation_class_count.trim() != "") ? $("#hdn_reservation_class_count").val((parseInt(str_reservation_class_count, 10) + 1).toString()) : $("#hdn_reservation_class_count").val("1");

        // 更新 預定課程筆數、總預計扣除課時數、目前剩餘的堂數
        g_obj_reservation_class_ctrl.updateReservationClassInfo(flt_subtract_session, "+");
    },

    // 取得預計扣除課時數顯示文字
    getSubtractSessionText: function (p_flt_subtract_session, p_str_order_class_datetime, p_str_class_type) {
        var str_text = g_obj_class_data.arr_msg.font_list_1;
        var str_mark_1 = g_obj_class_data.arr_msg.font_list_before_24_hour;
        var str_mark_2 = g_obj_class_data.arr_msg.font_list_between_12_hour;
        var str_mark_3 = g_obj_class_data.arr_msg.font_list_between_4_hour;
        var str_result = str_text;
        var str_diff_minute;

        if (p_flt_subtract_session >= 0) {
            // 取得目前系統時間和選到的課程時間差(minute)
            str_diff_minute = g_obj_reservation_class_ctrl.getDiffSystemAndClassMinute(p_str_order_class_datetime);

            // 24小時前
            if (parseInt(str_diff_minute, 10) > 1440) {
                str_result = str_text + str_mark_1;
            }
            // 24 ~ 12 小時
            else if (parseInt(str_diff_minute, 10) <= 1440 && parseInt(str_diff_minute, 10) > 720) {
                str_result = str_text + str_mark_2;
            }
            // 12 ~ 4 小時
            else if (parseInt(str_diff_minute, 10) <= 720 && parseInt(str_diff_minute, 10) > 240) {
                str_result = str_text + str_mark_3;
            }
        }
        return str_result;
    },

    // 取得目前系統時間和選到的課程時間差(minute)
    getDiffSystemAndClassMinute: function (p_str_order_class_datetime) {
        var str_order_class_datetime;

        // 預訂課程的時間
        str_order_class_datetime = new Date(p_str_order_class_datetime.substr(0, 4),
                                                parseInt(p_str_order_class_datetime.substr(4, 2), 10) - 1,
                                                p_str_order_class_datetime.substr(6, 2),
                                                p_str_order_class_datetime.substr(8, 2), 30);

        return (str_order_class_datetime.getTime() / (1000 * 60)) - (g_obj_reservation_class_ctrl.dat_sys_datetime.getTime() / (1000 * 60));
    },

    // 取得預計扣除課時數
    getSubtractSession: function (p_str_order_class_datetime, p_str_class_type) {
        // 訂課日期和時間 YYYY/MM/DD HH:mm
        var str_order_class_datetime_format = p_str_order_class_datetime.substr(0, 4) + "/"
                                                  + p_str_order_class_datetime.substr(4, 2) + "/"
                                                  + p_str_order_class_datetime.substr(6, 2) + " "
                                                  + p_str_order_class_datetime.substr(8, 2) + ":30";

        var str_post_data = "act=3&rcdatetime=" + escape(str_order_class_datetime_format) + "&rctype=" + p_str_class_type + "&lobbysn=" + g_obj_reservation_class_ctrl.strReservationLobbySn;
        var flt_subtract_session;
        var flt_subtract_session = $.ajax({
            type: "POST",
            cache: false,
            url: g_obj_reservation_class_ctrl.arr_page_url.ajax_class_info,
            data: str_post_data,
            async: false
        }).responseText;
        if (flt_subtract_session !== "") {
            flt_subtract_session = parseFloat(flt_subtract_session);
            if (flt_subtract_session >= 0) {
                return flt_subtract_session;
            }
            else {
                return flt_subtract_session.toString();
            }
        }
        else {
            return g_obj_class_data.arr_msg.reservation_class_fail_contact_cs;
        }
    },

    // 重新取得預計扣除的課時數
    ajaxRegetSubtractSession: function (event) {
        var str_class_type = event.data.class_type;

        // 訂課日期時間 YYYYMMDDHHmm
        var str_order_class_datetime = event.data.class_dt;

        // 訂課日期和時間 YYYY/MM/DD HH:mm
        var str_order_class_datetime_format = str_order_class_datetime.substr(0, 4) + "/"
                                                  + str_order_class_datetime.substr(4, 2) + "/"
                                                  + str_order_class_datetime.substr(6, 2) + " "
                                                  + str_order_class_datetime.substr(8, 2) + ":30";
        var str_post_data = "act=3&rcdatetime=" + escape(str_order_class_datetime_format) + "&rctype=" + str_class_type;
        var flt_subtract_session;
        $.ajax({
            type: "POST",
            cache: false,
            url: g_obj_reservation_class_ctrl.arr_page_url.ajax_class_info,
            data: str_post_data,
            success: function (data) {
                flt_subtract_session = parseFloat(data);
                if (flt_subtract_session >= 0) {
                    g_obj_reservation_class_ctrl.reOrderClass(str_order_class_datetime, flt_subtract_session, str_class_type);
                }
                else {
                    // 重訂失敗 顯示錯誤訊息
                    alert(g_obj_reservation_class_ctrl.getReservationFailMessage(flt_subtract_session.toString()));
                    //alert(g_obj_class_data.arr_msg.reorder_class_fault);
                }
            },
            error: function () {
                alert(g_obj_class_data.arr_msg.reorder_class_fault);
                //alert("Ajax Error!");
            }
        });
    },

    // 取得一對三或一對一要顯示的圖片檔名
    getClassImageFileName: function (p_str_class_type) {
        // 一對三訂課顯示
        if (p_str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three) {
            return "one_on_three.gif";
        }
        // 一對一訂課顯示
        else if (p_str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
            return "one_on_one.gif";
        }
        // 大會堂訂課顯示
        else if (p_str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
            return "lobby.gif";
        }
    },

    // 取得一對三或一對一要顯示的ball class name
    getBallClassName: function (p_str_class_type) {
        // 一對三訂課顯示 紅色圓點
        if (p_str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three) {
            return "ball_red";
        }
        // 一對一訂課顯示 綠色圓點
        else if (p_str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
            return "ball_green";
        }
        // 大會堂訂課顯示 藍色圓點
        else if (p_str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
            return "ball_blue";
        }
    },

    // 更新總預定課程筆數
    updateTotalClassCount: function () {
        $("#span_reservation_class_count").text($("#hdn_reservation_class_count").val());
    },

    // 更新總預計扣除課時數
    updateTotalSubtractSession: function (p_flt_subtract_session, p_str_opt) {
        var int_total_subtract_session = $("#span_reservation_class_total_subtract_session").text();

        (int_total_subtract_session.trim() == "") ? int_total_subtract_session = 0.0 : int_total_subtract_session = parseFloat(int_total_subtract_session);

        if (p_str_opt === "+") {
            $("#span_reservation_class_total_subtract_session").text((int_total_subtract_session + p_flt_subtract_session).toFixed(2).toString());
        }
        else if (p_str_opt === "-") {
            $("#span_reservation_class_total_subtract_session").text((int_total_subtract_session - p_flt_subtract_session).toFixed(2).toString());
        }
    },

    // 更新客戶目前剩餘的堂數(課時數)，已扣掉預訂的
    updateClientTotalSession: function () {
        var int_total_subtract_session = $("#span_reservation_class_total_subtract_session").text();
        var int_client_total_session = g_obj_reservation_class_ctrl.str_client_total_session;

        (int_total_subtract_session.trim() == "") ? int_total_subtract_session = 0 : int_total_subtract_session = parseFloat(int_total_subtract_session);
        (int_client_total_session.trim() == "") ? int_client_total_session = 0 : int_client_total_session = parseFloat(int_client_total_session);

        if (int_client_total_session != 0) {
            $("#span_client_total_session").text((int_client_total_session - int_total_subtract_session).toFixed(2).toString());
        }
    },

    // 更新 預定課程筆數、總預計扣除課時數、目前剩餘的堂數
    updateReservationClassInfo: function (p_flt_subtract_session, p_str_opt) {
        // 改變總預定課程筆數
        g_obj_reservation_class_ctrl.updateTotalClassCount();

        // 改變總預計扣除課時數 加或減
        g_obj_reservation_class_ctrl.updateTotalSubtractSession(p_flt_subtract_session, p_str_opt);

        // 改變客戶目前剩餘的堂數(課時數)，已扣掉預訂的
        g_obj_reservation_class_ctrl.updateClientTotalSession();
    },

    // parse訂課資訊列表 設定目前已預訂的課程
    setTotalReservationClass: function () {
        var arr_ul_child_len = $("#ul_reservation_class_info").children();  // 取得訂課列表上的每筆資料
        var str_order_class_datetime = "";                         // 所有 客戶已預訂的課程日期和時間 YYYYMMDDHH
        var str_all_order_class_dt = "";
        for (var i = 0; i < arr_ul_child_len.length; i++) {
            if (arr_ul_child_len[i].id != "") {
                str_order_class_datetime = arr_ul_child_len[i].id.substring(3, arr_ul_child_len[i].id.length);
                // 檢查是否還沒取消
                if ($("#hdn_enabled_" + str_order_class_datetime).val() == "1") {
                    // 大會堂
                    if ($("#hdn_class_type_" + str_order_class_datetime).val() === g_obj_class_data.arr_class_type.lobby) {
                        (str_all_order_class_dt != "") ? str_all_order_class_dt = str_all_order_class_dt + "," + $("#hdn_class_type_" + str_order_class_datetime).val() + str_order_class_datetime + $("#hdn_lobby_session_sn_" + str_order_class_datetime).val() : str_all_order_class_dt = $("#hdn_class_type_" + str_order_class_datetime).val() + str_order_class_datetime + $("#hdn_lobby_session_sn_" + str_order_class_datetime).val();
                    }
                    // 一對一和一對多
                    else {
                        (str_all_order_class_dt != "") ? str_all_order_class_dt = str_all_order_class_dt + "," + $("#hdn_class_type_" + str_order_class_datetime).val() + str_order_class_datetime : str_all_order_class_dt = $("#hdn_class_type_" + str_order_class_datetime).val() + str_order_class_datetime;
                    }
                }
            }
        }
        document.getElementById("hdn_total_reservation_class_dt").value = str_all_order_class_dt;
        $("#hdn_total_reservation_class_dt").val(str_all_order_class_dt);
    },

    // 從課程列表移除一筆 (大會堂用)
    removeClassInfo: function (p_str_order_class_datetime, p_str_class_type) {
        var str_reservation_class_count = $("#hdn_reservation_class_count").val();

        // 檢查是否課程資訊列表是否已存在
        if (document.getElementById("li_" + p_str_order_class_datetime) != null) {
            // 檢查此筆課程設定是否已 "取消" (1:有效 0:已取消)
            if ($("#hdn_enabled_" + p_str_order_class_datetime).val() == "1") {
                // 訂課筆數 - 1
                (str_reservation_class_count.trim() != "") ? $("#hdn_reservation_class_count").val((parseInt(str_reservation_class_count, 10) - 1).toString()) : $("#hdn_reservation_class_count").val("1");

                // 更新 預定課程筆數、總預計扣除課時數、目前剩餘的堂數
                g_obj_reservation_class_ctrl.updateReservationClassInfo(g_obj_reservation_class_ctrl.getSubtractSession(p_str_order_class_datetime, p_str_class_type), "-");
            }

            // 移除
            $("#li_" + p_str_order_class_datetime).remove();

            // 沒有預定大會堂
            if ($("#span_reservation_class_count").text() === "0") {
                $("#div_near_order_class").text(g_obj_class_data.arr_msg.a_no_have_order_class);
                $("#ul_reservation_class_info").hide();     // 隱藏 訂課資訊列表
                $("#div_reservation_class_result").hide();  // 隱藏 訂課結果及確認送出按鈕
                $("#div_no_reservation_class_for_lobby").show();
                $("#remind_contract_start_msg").show();     // 顯示 提示合約開始日訊息
                $("#div_no_reservation_class_remind_msg").show(); // 隱藏 找不到或沒有資料的共用訊息框
            }
        }
    },

    // 點擊 近期預定課程頁面 "確認訂課" 按鈕 事件
    // 點擊 近期取消課程頁面 "確認訂課" 按鈕 事件
    // 點擊 近期预订大会堂课程 "確認訂課" 按鈕 事件
    // 點擊 近期取消大会堂课程 "確認訂課" 按鈕 事件
    submitReservationClass: function () {
        // 設定目前已預訂的課程
        g_obj_reservation_class_ctrl.setTotalReservationClass();

        if (($("#hdn_total_reservation_class_dt").val()).trim() === "") {
            // 您目前尚未預訂課程
            alert(g_obj_class_data.arr_msg.no_add_class);
            return false;
        }
        return true;
    },

    // 每隔一分鐘 檢查訂的課程是否在四小時內
    checkClassNoCancel: function () {
        var arr_ul_child_len = $("#ul_reservation_class_info").children(); // 取得訂課列表上的每筆資料
        var str_order_class_datetime = "";                                 // 所有 客戶已預訂的課程日期和時間 YYYYMMDDHH
        var int_no_cancel_in_minute = 0;                                   // 幾分鐘內無法取消課程 minute
        var str_msg_no_cancel_in_minute = "";                              // 幾分鐘內無法取消課程 訊息
        var str_class_type = "";

        // 將系統時間 + 60秒
        g_obj_reservation_class_ctrl.dat_sys_datetime.setSeconds(g_obj_reservation_class_ctrl.dat_sys_datetime.getSeconds() + 60);

        for (var i = 0; i < arr_ul_child_len.length; i++) {
            if (arr_ul_child_len[i].id != "") {
                str_order_class_datetime = arr_ul_child_len[i].id.substring(3, arr_ul_child_len[i].id.length);

                // 檢查 是否還沒取消  且 還未轉換成 "開課前四小時內無法取消" 且 訂課時間在四小時內
                if ($("#hdn_enabled_" + str_order_class_datetime).val() === "1" && $("#hdn_no_cancel_in_four_" + str_order_class_datetime).val() === "0") {
                    str_class_type = $("#hdn_class_type_" + str_order_class_datetime).val();

                    // 一對三訂課
                    if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three) {
                        int_no_cancel_in_minute = g_obj_class_data.arr_class_no_cancel_in_minute.one_on_three;
                        str_msg_no_cancel_in_minute = g_obj_class_data.arr_msg_class_no_cancel_in_hour.one_on_three;
                    }
                    // 一對一訂課
                    else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
                        int_no_cancel_in_minute = g_obj_class_data.arr_class_no_cancel_in_minute.vip;
                        str_msg_no_cancel_in_minute = g_obj_class_data.arr_msg_class_no_cancel_in_hour.vip;
                    }
                    // 大會堂訂課
                    else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
                        int_no_cancel_in_minute = g_obj_class_data.arr_class_no_cancel_in_minute.lobby;
                        str_msg_no_cancel_in_minute = g_obj_class_data.arr_msg_class_no_cancel_in_hour.lobby;
                    }

                    if (g_obj_reservation_class_ctrl.getDiffSystemAndClassMinute(str_order_class_datetime) <= parseInt(int_no_cancel_in_minute, 10)) {
                        $("#hdn_no_cancel_in_four_" + str_order_class_datetime).val("1");
                        $("#h6_order_button_" + str_order_class_datetime).text(str_msg_no_cancel_in_minute);
                    }
                }
            }
        }

        // 設定 Timer 每隔一分鐘 檢查訂的課程是否在四小時內
        setTimeout(g_obj_reservation_class_ctrl.checkClassNoCancel, 60000);
    },

    // 取得訂課失敗的錯誤訊息 (要取得扣堂數發生錯誤時)
    getReservationFailMessage: function (p_str_err_code) {
        var str_err_msg = "";

        // 新增失敗: 新增到同一堂課
        if (p_str_err_code == g_obj_reservation_class_ctrl.arr_err_code.same_class) {
            str_err_msg = g_obj_class_data.arr_msg.sel_same_class;
        }
        // 新增失敗: 選到沒有開課的諮詢時間 4小時內
        else if (p_str_err_code == g_obj_reservation_class_ctrl.arr_err_code.class_no_open_in_4_hour) {
            // 一對三訂課顯示
            if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.one_on_three) {
                str_err_msg = g_obj_class_data.arr_msg_class_no_order_in_hour.one_on_three;
            }
            // 一對一訂課顯示
            else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.vip) {
                str_err_msg = g_obj_class_data.arr_msg_class_no_order_in_hour.vip;
            }
            // 大會堂訂課顯示
            else if (str_class_type == g_obj_reservation_class_ctrl.arr_class_type.lobby) {
                str_err_msg = g_obj_class_data.arr_msg_class_no_order_in_hour.lobby;
            }
        }
        // 新增失敗: 24小時內臨時訂課僅開放一對三
        else if (p_str_err_code == g_obj_reservation_class_ctrl.arr_err_code.class_no_open_in_24_hour) {
            str_err_msg = g_obj_class_data.arr_msg.class_no_open_in_24_hour;
        }
        // 新增失敗: 大會堂 只開放五分鐘前可以訂課
        else if (p_str_err_code == g_obj_reservation_class_ctrl.arr_err_code.class_no_open_in_5_min) {
            str_err_msg = g_obj_class_data.arr_msg.class_no_open_in_5_min;
        }
        // 新增失敗: 合约剩余课时数不足，请至「学习记录」查询详细资讯。
        else if (p_str_err_code == g_obj_reservation_class_ctrl.arr_err_code.not_enough_session) {
            // 合约剩余课时数不足，请至「学习记录」查询详细资讯。
            str_err_msg = g_obj_class_data.arr_msg.contract_edate_invalid_4 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3;
        }
        // 新增失敗: 訂位失敗，請與客服人員連絡：02-3365-9999 (因為錯誤的訂課日期或時間)
        else if (p_str_err_code == g_obj_reservation_class_ctrl.arr_err_code.wrong_class_dt) {
            // 訂位失敗，請與客服人員連絡：02-3365-9999
            str_err_msg = g_obj_class_data.arr_msg.reservation_class_fail_contact_cs;
        }
        else {
            str_err_msg = g_obj_class_data.arr_msg.reservation_class_fail_contact_cs;
        }

        return str_err_msg;
    },

    // 檢查訂的課程是否在本週的結算週期內
    isSessionInThisWeek: function (p_dat_session_date, p_dat_account_sdate, p_dat_account_edate) {
        var dat_session_date_yyyymmdd = p_dat_session_date.replace(new RegExp("/", "gm"), "");    //  課程日期 YYYYMMDDHH
        var dat_account_sdate_yyyymmdd = p_dat_account_sdate.replace(new RegExp("/", "gm"), "");  //  本週的週期初開始日 YYYYMMDDHH
        var dat_account_edate_yyyymmdd = p_dat_account_edate.replace(new RegExp("/", "gm"), "");  //  本週的週期末結束日 (結算日) YYYYMMDDHH

        var dat_session_date = new Date(dat_session_date_yyyymmdd.substr(0, 4),
                                          parseInt(dat_session_date_yyyymmdd.substr(4, 2), 10) - 1,
                                          dat_session_date_yyyymmdd.substr(6, 2));

        var dat_account_sdate = new Date(dat_account_sdate_yyyymmdd.substr(0, 4),
                                           parseInt(dat_account_sdate_yyyymmdd.substr(4, 2), 10) - 1,
                                           dat_account_sdate_yyyymmdd.substr(6, 2));

        var dat_account_edate = new Date(dat_account_edate_yyyymmdd.substr(0, 4),
                                           parseInt(dat_account_edate_yyyymmdd.substr(4, 2), 10) - 1,
                                           dat_account_edate_yyyymmdd.substr(6, 2));
        if (dat_session_date >= dat_account_sdate && dat_session_date <= dat_account_edate)
            return true;
        else
            return false;
    },

    // AJAX 取消課程事件
    // 訂課編號,訂課日期,訂課時間 //多傳了三個變數
    cancelClass: function (p_attend_list_sn, p_attend_date, p_attend_time, intPage, intPageType, strSessionIdName) {
        var str_post_data = "";
        var str_error_msg = "";
        var strSessionId = document.getElementById(strSessionIdName);
        str_post_data = "attend_list_sn=" + p_attend_list_sn +
                                "&attend_list_date=" + p_attend_date +
                                "&attend_list_time=" + p_attend_time;

        //if (confirm(g_obj_class_data.arr_msg.confirm_cancel_class))
        //{
        // 隱藏 "取消" 文字
        //$("#cancel_class_link_" + p_attend_list_sn).hide();

        // 顯示 等待讀取的圖片
        //$("#img_cancel_class_loading_" + p_attend_list_sn).show();
        $.ajax({
            type: "POST",
            cache: false,
            async: false,
            url: g_obj_reservation_class_ctrl.arr_page_url.ajax_cancel_class,
            data: str_post_data,
            success: function (data) {
                // 取消失敗
                if (data === "-1") {
                    str_error_msg = g_obj_class_data.arr_msg.cancel_class_fault;
                    strSessionId.checked = true;  //checkbox勾選
                }
                // 取消失敗 要取消的課程在四小時內 不能取消且checkbox自動打勾
                else if (data === "-2") {
                    str_error_msg = g_obj_class_data.arr_msg.font_no_cancel_in_4_hour;
                    strSessionId.checked = true; //checkbox勾選
                    //$("#" + strSessionIdName + "_IMG").addClass("check_on");
                    //$("#" + strSessionIdName + "_IMG").removeClass("check_off");
                    //$("#" + strSessionId).attr("checked", true);
                    //str_in_four_hours = "true";
                }

                // 取消成功
                if (str_error_msg === "") {
                    // 隱藏 等待讀取的圖片
                    //$("#img_cancel_class_loading_" + p_attend_list_sn).hide();

                    // 刪除訂課資訊
                    //$("#cancel_class_info_" + p_attend_list_sn).remove();
                    if ( document.getElementById(strSessionIdName) ) {
                        strSessionId.checked = false; //checkbox取消勾選
                    }
                    //$("#" + strSessionId).attr("checked", false);
                    alert(data);
                    //strClassIDString = "";
                    ChangePage(intPage, intPageType);
                }
                // 取消失敗
                else {
                    // 顯示 "取消" 文字
                    //$("#cancel_class_link_" + p_attend_list_sn).hide();

                    // 隱藏 等待讀取的圖片
                    //$("#img_cancel_class_loading_" + p_attend_list_sn).hide();

                    alert(str_error_msg);
                }
            },
            error: function () {
                // 顯示 "取消" 文字
                //$("#cancel_class_link_" + p_attend_list_sn).show();

                // 隱藏 等待讀取的圖片
                //$("#img_cancel_class_loading_" + p_attend_list_sn).hide();

                alert(g_obj_class_data.arr_msg.cancel_class_fault);
                //alert("Ajax Error");
            }
        });
        //}
    } // end of cancelClass
};