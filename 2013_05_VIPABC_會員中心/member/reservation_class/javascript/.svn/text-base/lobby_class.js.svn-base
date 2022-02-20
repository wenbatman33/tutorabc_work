var g_obj_reservation_lobby_class_ctrl = {
    // ajax or post or 其他if判斷式要用到的 的頁面URL
    arr_page_url: {
        "ajax_lobby_class_info": "/program/member/reservation_class/include/ajax_lobby_class_info.asp",        // ajax 動態載入 根據搜尋條件的大會堂資訊
        "ajax_lobby_class_search": "/program/member/reservation_class/include/ajax_lobby_class_search.asp",    // ajax 動態載入 搜尋大會堂開課資訊介面   
        "ajax_cancel_class": "/program/member/reservation_class/include/ajax_cancel_class.asp",                // ajax 取消課程
        "ajax_class_info": "/program/member/reservation_class/include/ajax_class_info.asp"                     // ajax 取值用
    },


    // 載入大會堂開課的資訊
    searchLobbyClassInfo: function () {
        $("#div_lobby_class_info").empty();
        $("#div_lobby_class_info").append("<img id=\"img_ajax_loader\" src=\"/lib/javascript/JQuery/images/ajax-loader.gif\" border=\"0\" />");
        $("#div_lobby_class_info").load(g_obj_reservation_lobby_class_ctrl.arr_page_url.ajax_lobby_class_info,
                                            {
                                                ndate: new Date(),
                                                lobby_topic: $("#sel_search_lobby_topic").val(),
                                                con_sn: $("#sel_search_lobby_consultant").val(),
                                                lobby_suit_lev: $("#sel_level").val(),
                                                lobby_time: $("#sel_search_lobby_time").val(),
                                                page: $("#page").val()
                                            });
    },
    // 新增一筆訂課到列表
    addNewLobbyClass: function (p_str_lobby_session_sn, p_bol_from_other_page, p_str_lobby_date, p_str_lobby_time, p_str_con_ename, p_str_lobby_topic, p_str_lobby_sn) {

//        if ($("#chk_lobby_class_" + p_str_lobby_session_sn).attr('checked')) {
//            //PM11112104 警告大會堂已上過這類課程 			
//            $.ajax({
//                url: 'include/reservation_lobby_check_order_ajax.asp',
//                data: { ele_val_str: p_str_lobby_session_sn },
//                cache: false,
//                async: false,
//                error: function (xhr) {
//                    alert("錯誤" + xhr.responseText)
//                },
//                success: function (result) {
//                    if (result != "0") {
//                        alert(result + "上过本教材");
//                    }
//                }
//            });
//        }

        var obj_chk_id = ""; //頁面上checkbox id
        var str_class_date = ""; //開課日期
        var str_class_time = ""; //開課時間
        var str_con_ename = ""; //顧問名字
        var str_topic = ""; //大會堂主題
        var bol_is_add_new; //true: 新增, false: 刪除
        var int_add_result;

        if (!p_bol_from_other_page) {
            obj_chk_id = "chk_lobby_class_" + p_str_lobby_session_sn; //頁面上checkbox id
            str_class_date = $("#span_lobby_class_date_" + p_str_lobby_session_sn).text(); //開課日期
            str_class_time = $("#span_lobby_class_time_" + p_str_lobby_session_sn).text(); //開課時間
            str_con_ename = $("#hdn_lobby_class_con_ename_" + p_str_lobby_session_sn).val(); //顧問名字
            str_topic = $("#span_lobby_class_topic_" + p_str_lobby_session_sn).text(); //大會堂主題                

            if (document.getElementById(obj_chk_id).checked) {
                bol_is_add_new = true;
            }
            else {
                bol_is_add_new = false;
            }
            //alert(bol_is_add_new);
        }
        else {
            var str_post_data = "act=chkorder&rcsn=" + escape(p_str_lobby_session_sn) + "&rcdate=" + escape(p_str_lobby_date) + "&rctime=" + escape(p_str_lobby_time.substring(0, 2));
            // 檢查客戶是否已經預定了此時段的課程
            var str_chk_already_reservation = $.ajax({
                type: "POST",
                cache: false,
                url: g_obj_reservation_lobby_class_ctrl.arr_page_url.ajax_class_info,
                data: str_post_data,
                async: false
            }).responseText;

            if (str_chk_already_reservation == "have") {
                //alert(g_obj_class_data.arr_msg.this_class_already_order);
                return false;
            }

            str_class_date = p_str_lobby_date;      // 開課日期
            str_class_time = p_str_lobby_time;      // 開課時間
            str_con_ename = p_str_con_ename;        // 顧問名字
            str_topic = p_str_lobby_topic;          // 大會堂主題

            bol_is_add_new = true;

            $("#div_no_reservation_class_for_lobby").hide();
            $("#div_no_reservation_class_remind_msg").hide(); // 隱藏 找不到或沒有資料的共用訊息框
        }
        str_order_class_datetime = str_class_date.replace(new RegExp("/", "gm"), "") + str_class_time.substring(0, 2); //  訂課日期時間 YYYYMMDDHHmm

        if (bol_is_add_new) {
            var str_contract_start_date_yyyymmdd = g_obj_class_data.str_service_start_date.replace(new RegExp("/", "gm"), ""); //  服務開始日 YYYYMMDDHH
            var str_contract_end_date_yyyymmdd = g_obj_class_data.str_service_end_date.replace(new RegExp("/", "gm"), "");  //  服務結束日 YYYYMMDDHH
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
                if (obj_chk_id != "") {
                    document.getElementById(obj_chk_id).checked = false;
                }
                return false;
            }

            // 檢查是否選到合約開始日之前的日期
            if (dat_contract_start_date > dat_class_date) {
                // 請預訂YYYY/MM/DD(合約開始)之後的課程。
                alert(g_obj_class_data.arr_msg.reservation_after_contract_start);
                if (obj_chk_id != "") {
                    document.getElementById(obj_chk_id).checked = false;
                }
                return false;
            }

            // 檢查是否選到合約結束日之後的日期
            if (dat_class_date > dat_contract_end_date) {
                // 合約于YYYY/MM/DD到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
                alert(g_obj_class_data.arr_msg.contract_edate_invalid_1 + g_obj_class_data.str_service_end_date + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
                if (obj_chk_id != "") {
                    document.getElementById(obj_chk_id).checked = false;
                }
                return false;
            }

            // 檢查合約是否未到期
            if (!g_obj_class_data.bol_client_contract_edate_valid) {
                // 合約于YYYY/MM/DD到期，請至 ｢學習紀錄｣ 查詢詳細資訊。
                alert(g_obj_class_data.arr_msg.contract_edate_invalid_1 + g_obj_class_data.str_service_end_date + g_obj_class_data.arr_msg.contract_edate_invalid_2 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
                if (obj_chk_id != "") {
                    document.getElementById(obj_chk_id).checked = false;
                }
                return false;
            }

            // 總剩餘堂數 = 剩餘堂數
            var flt_available_points = parseFloat(g_obj_class_data.str_client_total_session) + parseFloat(g_obj_class_data.str_client_week_session);

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

            if (flt_available_points <= 0) 
            {
                // 合约剩余课时数不足，请至「学习记录」查询详细资讯。
                alert(g_obj_class_data.arr_msg.contract_edate_invalid_4 + "，" + g_obj_class_data.arr_msg.contract_edate_invalid_3);
                if (obj_chk_id != "") 
                {
                    document.getElementById(obj_chk_id).checked = false;
                }
                return false;
            }

            int_add_result = g_obj_reservation_class_ctrl.addNewClassInfo(g_obj_reservation_class_ctrl.arr_class_type.lobby, str_class_date, str_class_time.substring(0, 2), str_con_ename, str_topic, p_str_lobby_session_sn);

            // 新增失敗
            if (parseFloat(int_add_result) < 0) 
            {
                // 顯示錯誤訊息
                alert(g_obj_reservation_class_ctrl.getReservationFailMessage(int_add_result.toString()));
                if (obj_chk_id != "") 
                {
                    document.getElementById(obj_chk_id).checked = false;
                }
                return false;
            }
            $("#btn_search_lobby").focus();
        }
        else {
            g_obj_reservation_class_ctrl.removeClassInfo(str_order_class_datetime, g_obj_reservation_class_ctrl.arr_class_type.lobby);
        }
    }
};