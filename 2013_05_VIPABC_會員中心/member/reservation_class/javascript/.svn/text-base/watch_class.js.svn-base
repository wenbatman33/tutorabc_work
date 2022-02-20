    var g_obj_class_calender = {
        
        // ajax or post or 其他if判斷式要用到的 的頁面URL
        arr_page_url : {
            "ajax_class_list_calender_head" : "/program/member/reservation_class/include/ajax_class_list_calender_head.asp",    // ajax 觀看課表 月曆模式 顯示年和切換月的按鈕
            "ajax_class_list_calender" : "/program/member/reservation_class/include/ajax_class_list_calender.asp",              // ajax 觀看課表 月曆模式 顯示日期
            "ajax_class_list_week_head" : "/program/member/reservation_class/include/ajax_class_list_week_head.asp",            // ajax 觀看課表 列表模式 顯示年和切換月的按鈕
            "ajax_class_list_week" : "/program/member/reservation_class/include/ajax_class_list_week.asp",                      // ajax 觀看課表 列表模式 顯示日期
            "ajax_cancel_class" : "/program/member/reservation_class/include/ajax_cancel_class.asp"                             // ajax 取消課程
        },

        // 月曆模式 切換上下個月時的事件
        refreshCalender : function(p_dat_date, p_dat_month)
        {
            // 載入 行事曆 顯示日期
            $("#div_calender_day").load(g_obj_class_calender.arr_page_url.ajax_class_list_calender, {ndate : new Date(), dat_year : p_dat_date, dat_month : p_dat_month});

            // 載入 顯示年和切換月的按鈕
            $("#div_calender_head").load(g_obj_class_calender.arr_page_url.ajax_class_list_calender_head, {ndate : new Date(), dat_year : p_dat_date, dat_month : p_dat_month});
        },

        // 列表模式 切換上下周時的事件
        refreshWeek : function(p_dat_monday)
        {
            // 載入 行事曆 顯示日期
            $("#div_calender_day").load(g_obj_class_calender.arr_page_url.ajax_class_list_week, {ndate : new Date(), dat_monday : p_dat_monday});

            // 載入 顯示年和切換月的按鈕
            $("#div_calender_head").load(g_obj_class_calender.arr_page_url.ajax_class_list_week_head, {ndate : new Date(), dat_monday : p_dat_monday});
        },

        // 切換行事曆模式 事件
        onchangeSelectCalenderMode : function()
        {
            // 月历模式
            if ($("#sel_change_calender_mode").val() === "1")
            {
                g_obj_class_calender.refreshCalender("", "");
            }
            // 列表模式
            else
            {
                g_obj_class_calender.refreshWeek("");
            }
        }      
    };

    $(document).ready(function(){    
        g_obj_class_calender.refreshCalender("", "");
    });