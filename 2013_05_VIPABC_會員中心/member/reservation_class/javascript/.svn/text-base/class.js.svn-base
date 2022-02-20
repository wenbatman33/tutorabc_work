var g_obj_class = {

    // ajax or post or 其他if判斷式要用到的 的頁面URL
    arr_page_url: {
        "ajax_class_list_week_head": "include/ajax_reservation_class_change_week.asp",    // ajax 顯示年和切換月的按鈕
        "ajax_class_list_week": "include/ajax_reservation_class_chose_time.asp",          // ajax 顯示時間
        "ajax_class_description": "include/ajax_reservation_class_desc.asp"               // ajax 訂課說明
    },

    // 列表模式 切換上下周時的事件
    refreshWeek: function (p_dat_monday, p_str_class_type, p_str_product_type, p_str_point_type)
    {
        /*
        var strAlreadyCancel = document.getElementById("hdn_already_cancle").value;
        var strAlreadyReservation = document.getElementById("hdn_already_reservation").value;
        var bolGoNextWeek = false;

        // 判斷是否有在頁面上取消課程或預訂課程
        // 提示使用者 []
        if ("yes" == strAlreadyCancel || "yes" == strAlreadyReservation)
        {
        }*/

        // 載入 行事曆 顯示日期
        $("#div_chose_time").load(g_obj_class.arr_page_url.ajax_class_list_week,
                                {
                                    ndate: new Date(),
                                    dat_monday: p_dat_monday,
                                    client: g_str_client_sn,
                                    class_type: p_str_class_type,
                                    product_type: p_str_product_type,
                                    product_point_type: p_str_point_type
                                });


        // 載入 顯示年和切換月的按鈕
        $("#div_change_week").load(g_obj_class.arr_page_url.ajax_class_list_week_head,
                                {
                                    ndate: new Date(),
                                    dat_monday: p_dat_monday,
                                    class_type: p_str_class_type,
                                    product_type: p_str_product_type,
                                    product_point_type: p_str_point_type
                                });
    },

    // 選擇課程類型
    changeClassType: function (p_str_class_type, p_str_product_type, p_str_point_type) 
    {
        // 載入訂課說明
        $("#div_class_description").load(g_obj_class.arr_page_url.ajax_class_description,
                                        {
                                            ndate: new Date(),
                                            client: g_str_client_sn,
                                            product_sn: g_str_now_use_product_sn,
                                            class_type: p_str_class_type,
                                            product_type: p_str_product_type,
                                            product_point_type: p_str_point_type
                                        });
        $("#div_change_week").empty();
        $("#div_chose_time").empty();
        $("#div_chose_time").html("<img src='/javascript/JQuery/images/ajax-loader.gif'>");
        g_obj_class.refreshWeek("", p_str_class_type, p_str_product_type, p_str_point_type);

    },

    submitForm: function () {
        document.getElementById("btn_submit").disabled = true;
        return true;
    }
};

$(document).ready(function ()
{
    // 載入訂課說明
    $("#div_class_description").load(g_obj_class.arr_page_url.ajax_class_description,
                                    {
                                        ndate: new Date(), 
                                        client: g_str_client_sn,
                                        product_sn: g_str_now_use_product_sn,
                                        product_type: g_str_product_type,
                                        product_point_type: g_str_product_point_type,
                                        class_type: g_strClassType,
                                        service_start: g_strServiceStart
                                    });


    // 載入下個星期
    $("#div_change_week").load(g_obj_class.arr_page_url.ajax_class_list_week_head,
                                    {
                                        ndate: new Date(),
                                        product_type: g_str_product_type,
                                        product_point_type: g_str_product_point_type,
                                        class_type: g_strClassType
                                    });


    // 載入選擇訂課時間
    $("#div_chose_time").load(g_obj_class.arr_page_url.ajax_class_list_week,
                                    { 
                                        ndate: new Date(), client: g_str_client_sn,
                                        product_type: g_str_product_type,
                                        product_point_type: g_str_product_point_type,
                                        class_type: g_strClassType 
                                    });

});