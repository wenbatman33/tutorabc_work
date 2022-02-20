/// <summary>
/// 錯誤代碼和錯誤訊息的定義 處理
/// </summary>
/// <remarks>[RayChien] 2009/12/01 Created</remarks>
//
/// <param name="g_obj_error_handle">全域變數</param>
var g_obj_error_handle = 
{
    success_code : 1,
    error_code : {
        //參數錯誤判斷-1~-100
        "param_is_null" : -1,
        "param_is_undefined" : -2,
        "param_type_is_invalid" : -3,
        "param_is_not_array" : -4,
        //input輸入判斷 -101~-200
        "form_elements_are_empty" : -101,
        "input_data_is_empty" : -102,
        "input_is_not_printable" : -103,
        "input_length_error":-104,
        "input_not_empty":-105,
        "input_range_error":-106,
        //格式錯誤判斷-201~-300
        "number_format_error" : -201,
        "password_format_error" : -202,
        "passwordconfirm_format_error" : -203,
        "age_format_error" : -204,
        "stardate_and_enddate_format_error" : -205,
        "date_format_error" : -206,
        "email_format_error" : -207,
        "msn_id_format_error" : -208,
        "chinese_format_error" : -209,
        "chinese_name_format_error" : -210,
        "english_format_error" : -211,
        "english_name_format_error" : -212,
        "phone_format_error" : -213,
        "qq_id_format_error" : -214,
        "url_format_error" : -215,
        "NumberOrEnglish_format_error":-216
    }
    
}

