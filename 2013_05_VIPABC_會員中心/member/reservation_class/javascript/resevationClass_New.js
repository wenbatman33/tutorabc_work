    var g_obj_resevation_class_calender = {
        
        // ajax or post or 其他if判斷式要用到的 的頁面URL
        arr_page_url : {
            "ajax_class_list_calender_head" : "/program/member/reservation_class/include/ajax_resevationClass_New_calender_head.asp",	// ajax 訂課 周曆模式 顯示年和切換月的按鈕
            "ajax_resevationClass_New_calender" : "/program/member/reservation_class/include/ajax_resevationClass_New_calender.asp",	// ajax 訂課 周曆模式 週曆
			"resevationClass_New" : "/program/member/reservation_class/resevationClass_New.asp",		// ajax 訂課 周曆模式 內涵 年和切換月的按鈕 與 週曆
			"reservation_class_exe" : "/program/member/reservation_class/reservation_class_exe.asp"		// 點擊 近期預定課程頁面 "確認訂課" 按鈕 要POST FROM的頁面
        },
		
		//arr_class_lessons : {	//每一種課程，計算堂數的規則
		//	<%=CONST_CLASS_TYPE_ONE_ON_ONE%> : 3,
		//	<%=CONST_CLASS_TYPE_ONE_ON_THREE%> : 
		//	{
		//		"normal" : 1,
		//		"in_24_to_12" : 1.25,
		//		"in_12_to_4" : 1.5,
		//		"in_4_to_2" : 2
		//	},
		//	<%=CONST_CLASS_TYPE_LOBBY%> : "1"
		//},

        // 月曆模式 切換上下個月時的事件
        refreshCalender : function(p_dat_date, p_dat_month, p_dat_day)
        {
			var rs_class_type = $('#hd_reservation_class_type').val();
			
			// 載入 行事曆 顯示日期
            $("#div_calender_day").load(g_obj_resevation_class_calender.arr_page_url.ajax_resevationClass_New_calender, {ndate : new Date(), class_type: rs_class_type, dat_year : p_dat_date, dat_month : p_dat_month, dat_day : p_dat_day});
			
         	// 載入 顯示年和切換月的按鈕
            $("#div_calender_head").load(g_obj_resevation_class_calender.arr_page_url.ajax_class_list_calender_head, {ndate : new Date(), class_type: rs_class_type, dat_year : p_dat_date, dat_month : p_dat_month, dat_day : p_dat_day});
        },
		
		//從訂課型類選擇入口載入定課畫面(週曆)
		showReservatableClass : function(p_class_type)
        {
			$.post(
				g_obj_resevation_class_calender.arr_page_url.resevationClass_New, 
				{class_type: p_class_type},
				function(data) {
					$('#reservation_class_top').html(data);
					g_obj_resevation_class_calender.refreshCalender("", "", "");	//載入定課頁面 週曆型態
				},
				"text"
			);
        },
	
		//訂課事件
		reservateClass: function(p_class_type)
		{	
			return false;
			//TODO: 訂課等後台
			var arr_chk_res_class = $(":checkbox[name='chk_res_class']:checked");
			var chk_number = arr_chk_res_class.size();
			var str_order_class = "";
			
			if(chk_number>0)
			{	
				for (index_rCjs=0; index_rCjs<chk_number; index_rCjs++)
				{
					str_order_class = str_order_class + "," + arr_chk_res_class.eq(index_rCjs).attr("id");
					
					//arr_chk_res_class.eq(index_rCjs).attr("id")
					//年年年年月月日日時時
				}
			}
			//alert(g_obj_resevation_class_calender.arr_class_lessons.p_class_type);
			//getClassMsgForEmail(
			//(ByVal p_int_class_type, p_str_class_date, p_str_class_time, str_opt)
			str_order_class = str_order_class.substr(1, str_order_class.length - 1);
			//if(confirm(before_confirm_order_class + "\n\n"))
			//{
			//	$('#hd_reservation_class_type').val(str_order_class);
			//}
			alert(str_order_class);
			return false;
			
		},
		
        // 取消課程事件
        // 訂課編號,訂課日期,訂課時間
        cancelClass : function(p_attend_list_sn, p_attend_date, p_attend_time)
        {
            //var str_post_data = "";
            //var str_error_msg = "";

            //str_post_data = "attend_list_sn=" + p_attend_list_sn +                            
            //                "&attend_list_date=" + p_attend_date +
            //                "&attend_list_time=" + p_attend_time;

            //if (confirm(g_obj_class_data.arr_msg.confirm_cancel_class))
            //{
            //    // 隱藏 "取消" 文字
            //    $("#cancel_class_link_" + p_attend_list_sn).hide();

                // 顯示 等待讀取的圖片
            //    $("#img_cancel_class_loading_" + p_attend_list_sn).show();

            //    $.ajax({
		    //        type:"POST",
		    //        cache:false,
		    //        url:g_obj_resevation_class_calender.arr_page_url.ajax_cancel_class,
		    //        data:str_post_data,
		    //        success:function(data){
                        // 取消失敗
            //            if (data === "-1")
            //            {
            //                str_error_msg = g_obj_class_data.arr_msg.cancel_class_fault;
            //            }
                        // 取消失敗 要取消的課程在四小時內
            //            else if (data === "-2")
            //            {
            //                str_error_msg = g_obj_class_data.arr_msg.font_no_cancel_in_4_hour;
            //            }

                        // 取消成功
            //            if (str_error_msg === "")                        
            //            {
                            // 隱藏 等待讀取的圖片
            //                $("#img_cancel_class_loading_" + p_attend_list_sn).hide();

                            // 刪除訂課資訊
            //                $("#cancel_class_info_" + p_attend_list_sn).remove();

            //                alert(data);
            //            }
            //            // 取消失敗
            //            else
            //            {
                            // 顯示 "取消" 文字
            //                $("#cancel_class_link_" + p_attend_list_sn).hide();

                            // 隱藏 等待讀取的圖片
            //                $("#img_cancel_class_loading_" + p_attend_list_sn).hide();

            //                alert(str_error_msg);
            //            }
		    //        },
		    //        error:function(){
                        // 顯示 "取消" 文字
            //            $("#cancel_class_link_" + p_attend_list_sn).show();

                        // 隱藏 等待讀取的圖片
            //            $("#img_cancel_class_loading_" + p_attend_list_sn).hide();

            //            alert(g_obj_class_data.arr_msg.cancel_class_fault);
                        //alert("Ajax Error");
		    //        }
	        //    });
            //}
        }
    };
