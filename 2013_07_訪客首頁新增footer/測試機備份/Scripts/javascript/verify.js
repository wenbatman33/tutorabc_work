///<summary>
/// 資料的正確性、格式的驗證
/// </summary>
/// <remarks>[RayChien] 2009/10/29 Created</remarks>
///
/// <param name="g_obj_verify">全域變數</param>
var g_obj_verify = 
{
    /// <summary>
    /// 驗證 是否null
    /// </summary>
    /// <param name="p_var_input">輸入值</param>
    /// <example>
    /// var p_var_input;
    /// g_obj_verify.isNull(p_var_input); // true
    /// var p_var_input=123;
    /// g_obj_verify.isNull(p_var_input); // false
    /// </example>
    /// <returns>
    /// 等於null: true
    /// 不等於null: false
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isNull : function (p_var_input)    //check OK!
    {
        return !p_var_input;
    },
    /// <summary>
    /// 驗證 是否為Array
    /// </summary>
    /// <param name="p_var_input">輸入值</param>
    /// <example>
    /// var p_var_input = new Array(4);
    /// g_obj_verify.isArray(p_var_input);  // return 1
    /// var p_var_input = "notarray";
    /// g_obj_verify.isArray(p_var_input);  // return -5
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isArray : function (p_var_input)  //check OK!
    {
        try
        {
            if (g_obj_verify.isNull(p_var_input))
            {
                //回傳參數為null
                return g_obj_error_handle.error_code.param_is_null;
            }
            if(p_var_input.constructor == Array)
            {
                return g_obj_error_handle.success_code;
            }
            //回傳參數不為陣列
            return g_obj_error_handle.error_code.param_is_not_array;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 驗證 只允許數字或英文字母
    /// </summary>
    /// <param name="p_str_input">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_input = "ad1f31badf";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNumberOrEnglish(p_str_input,p_empty_ok); // return 1
    /// var p_str_input = "#%&*()";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNumberOrEnglish(p_str_input,p_empty_ok); // return -24
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isNumberOrEnglish : function (p_str_input, p_empty_ok)   //check OK!
    {
        try
        {
            // 正則表示式: 只允許數字及英文字母
            var str_chk_reg = /^[a-zA-Z0-9]+$/;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_input, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_input = p_str_input.trim();
            //允許不帶值，且p_var_input也真的為空值，則回傳true
            if (p_str_input == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合數字regular expression
			var b_chk_ret = (p_str_input.match(str_chk_reg) == p_str_input);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.NumberOrEnglish_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }

    },

    /// <summary>
    /// 驗證 是否為數字
    /// </summary>
    /// <param name="p_str_number">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_number = "1565161";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNumber(p_str_number,p_empty_ok); // return 1
    /// var p_str_number = "#%&*()";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNumber(p_str_number,p_empty_ok); // return -7
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isNumber : function (p_str_number, p_empty_ok)   //check OK!
    {
        try
        {
            // 正則表示式: 數字以外
            //var str_chk_reg = /\D/; 
			var str_chk_reg = /^\s*\d+\s*$/g;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_number, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_number = p_str_number.trim();
            //允許不帶值，且p_var_input也真的為空值，則回傳true
            if (p_str_number == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合數字regular expression
			var b_chk_ret = (p_str_number.match(str_chk_reg) == p_str_number);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.number_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },

        /// <summary>
    /// 驗證 是否為正負數數字
    /// </summary>
    /// <param name="p_str_number">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_number = "1565161";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNegativeOrNumber(p_str_number,p_empty_ok); // return 1
    /// var p_str_number = "#%&*()";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNegativeOrNumber(p_str_number,p_empty_ok); // return -7
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isNegativeOrNumber : function (p_str_number, p_empty_ok)   //check OK!
    {
        try
        {
            // 正則表示式: 數字以外
            //var str_chk_reg = /\D/; 
			var str_chk_reg = /^\s*[-,.,0-9]+\s*$/g;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_number, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_number = p_str_number.trim();
            //允許不帶值，且p_var_input也真的為空值，則回傳true
            if (p_str_number == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合數字regular expression
			var b_chk_ret = (p_str_number.match(str_chk_reg) == p_str_number);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.number_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },

    /// <summary>
    /// 驗證 是否符合日期格式 (YYYY/MM/DD or YYYY-MM-DD)
    /// </summary>
    /// <param name="p_str_dat">Date</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_dat = "2009/11/09";
    /// var p_empty_ok=false;
    /// g_obj_verify.isDate(p_str_dat,p_empty_ok); // return 1
    /// var g_str_tmp = "2009/13/09";
    /// var p_empty_ok=false;
    /// g_obj_verify.isDate(p_str_dat,p_empty_ok); // return -12
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isDate : function (p_str_dat, p_empty_ok)       //check OK!
    {
		try
        {
            // Date regular expression: 比對字串是否有出現 "2009/10/20" or "2009-10-20" or "2009 10 20"這樣的格式
            var str_chk_reg = /^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/g;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_dat, "string") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //判斷regular expression 格式是否符合

            var b_chk_ret = (p_str_dat.match(str_chk_reg) == p_str_dat);
            //去除字串的前後空白
            p_str_dat = p_str_dat.trim();
            //允許不帶值，且p_str_dat也真的為空值，則回傳true
            if (p_str_dat == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.date_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
	},

		

    /// <summary>
    /// 驗證 Email Address的格式
    /// </summary>
    /// <param name="p_str_email">Email</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_email = "raychien@tutorabc.com";
    /// var p_empty_ok=false;
    /// g_obj_verify.isEmailValid(p_str_email,p_empty_ok);     // return 1
    /// var p_str_email = "ff jon@ utorabc.com";
    /// var p_empty_ok=false;
    /// g_obj_verify.isEmailValid(p_str_email,p_empty_ok);     // return -13
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isEmailAddressValid : function (p_str_email, p_empty_ok)         //check OK!
    {
        try
        {
            // Email regular expression
            var str_chk_reg = /^\w+([\._-]?\w+)*@\w+([\._-]?\w+)*(\.\w{2,3})+$/g;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            // 判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_email, "string") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_email = p_str_email.trim();
            //允許不帶值，且p_str_email也真的為空值，則回傳true
            if (p_str_email == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //若email符合regular expression回傳true
			var b_chk_ret = (p_str_email.match(str_chk_reg) == p_str_email);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.email_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 驗證 密碼
    /// </summary>
    /// <param name="p_str_pwd">密碼</param>
    /// <param name="p_int_len_min">至少要幾位數</param>
    /// <param name="p_int_len_max">至多幾位數</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_pwd="A232*fsadf";
    /// var p_int_len_min=5;
    /// var p_int_len_max=20;
    /// var p_empty_ok=false;
    /// g_obj_verify.isPasswordValid(p_str_pwd,p_int_len_min,p_int_len_max,p_empty_ok); //return 1
    /// var p_str_pwd="A2j中文sadf";
    /// var p_int_len_min=5;
    /// var p_int_len_max=20;
    /// var p_empty_ok=false;
    /// g_obj_verify.isPasswordValid(p_str_pwd,p_int_len_min,p_int_len_max,p_empty_ok); //return -8
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isPasswordValid : function (p_str_pwd, p_int_len_min, p_int_len_max, p_empty_ok)          //check OK!
    {
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_pwd, "string") || !g_obj_verify.isParamTypeValid(p_int_len_min, "number") || !g_obj_verify.isParamTypeValid(p_int_len_max,"number") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //允許不帶值，且p_str_pwd也真的為空值，則回傳true
            if (p_str_pwd == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認字數是否符合要求
            if (!(p_int_len_min <= p_str_pwd.length && p_str_pwd.length <= p_int_len_max))
            {
                return g_obj_error_handle.error_code.input_length_error;
            }
            // 是否符合ascii Printable
            if (g_obj_verify.isPrintable(p_str_pwd))
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.password_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 驗證 密碼確認
    /// </summary>
    /// <param name="p_str_pwd">密碼</param>
    /// <param name="p_str_confirm_pwd">確認密碼</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_pwd="A232*fsadf";
    /// var p_str_confirm_pwd="A232*fsadf";
    /// var p_empty_ok=false;
    /// g_obj_verify.isPasswordConfirmValid(p_str_pwd,p_str_confirm_pwd,p_empty_ok); //return 1
    /// var p_str_pwd="A232*fsadf";
    /// var p_str_confirm_pwd="A1321fsadf";
    /// var p_empty_ok=false;
    /// g_obj_verify.isPasswordConfirmValid(p_str_pwd,p_str_confirm_pwd,p_empty_ok); //return -9
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
    isPasswordConfirmValid : function (p_str_pwd, p_str_confirm_pwd, p_empty_ok)   //check OK!
    {
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_pwd, "string") || !g_obj_verify.isParamTypeValid(p_str_confirm_pwd, "string") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //允許不帶值，且p_str_pwd也真的為空值，則回傳true
            if (p_str_pwd == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //回傳密碼與確認密碼是否相同
            if(p_str_pwd == p_str_confirm_pwd)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.passwordconfirm_format_error; 
        }
        catch(e)
        {
            return e.description;
        }
    },

    /// <summary>
    /// 驗證 中文姓名
    /// </summary>
    /// <param name="p_str_chname">中文姓名</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_chname="汪哥哥";
    /// var p_empty_ok=false;
    /// g_obj_verify.isChineseNameValid(p_str_chname,p_empty_ok); //return 1
    /// var p_str_chname="汪adg哥";
    /// var p_empty_ok=false;
    /// g_obj_verify.isChineseNameValid(p_str_chname,p_empty_ok); //return -16
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/24 Created</remarks>
	isChineseNameValid : function (p_str_chname, p_empty_ok)  //check OK!
    {
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_chname, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_chname = p_str_chname.trim();
            //允許不帶值，且p_str_chname也真的為空值，則回傳true
            if (p_str_chname == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認中文姓名長度是否符合要求
            if(!(1 <= p_str_chname.length && p_str_chname.length <= 10))
            {
                return g_obj_error_handle.error_code.input_length_error;
            }
            //確認中文姓名是否符合中文的 regular expression
            if(g_obj_verify.isChinese(p_str_chname) == g_obj_error_handle.success_code)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.chinese_name_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },

    /// <summary>
    /// 驗證 英文名字
    /// </summary>
    /// <param name="p_str_enname">英文名字</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_enname="Ryan Lin";
    /// var p_empty_ok=false;
    /// g_obj_verify.isEnglishNameValid(p_str_enname,p_empty_ok); //return 1
    /// var p_str_enname="adfabv***n";
    /// var p_empty_ok=false;
    /// g_obj_verify.isEnglishNameValid(p_str_enname,p_empty_ok); //return -18
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RayChien] 2009/10/29 Created</remarks>
    isEnglishNameValid : function (p_str_enname, p_empty_ok)    //check OK!
    {
        try
        {
            //English Name regular expression
            var str_patten = /^[A-Za-z\-\.]+\s?[A-Za-z\-]+$/;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_enname, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_enname = p_str_enname.trim();
            //允許不帶值，且p_str_enname也真的為空值，則回傳true
            if (p_str_enname == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認英文姓名字數是否符合要求
            if(!(1 <= p_str_enname.length && p_str_enname.length <= 30))
            {
                return g_obj_error_handle.error_code.input_length_error;
            }
            //確認是否符合英文姓名的 regular expression
			var b_chk_ret = (p_str_enname.match(str_patten) == p_str_enname);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.english_name_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 驗證 起始時間
    /// </summary>
    /// <param name="p_str_start_date">開始時間</param>
    /// <param name="p_str_end_date">結束時間</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_start_date="2008/01/03";
    /// var p_str_end_date="2008/03/05";
    /// var p_empty_ok=false;
    /// g_obj_verify.isStartAndEndDateValid(p_str_start_date,p_str_end_date,p_empty_ok); //return 1
    /// var p_str_start_date="2008/01/03";
    /// var p_str_end_date="2007/01/05";
    /// var p_empty_ok=false;
    /// g_obj_verify.isStartAndEndDateValid(p_str_start_date,p_str_end_date,p_empty_ok); //return -11
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RayChien] 2009/10/29 Created</remarks>
    isStartAndEndDateValid : function (p_str_start_date, p_str_end_date, p_empty_ok)    //check OK!
    {
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_start_date, "string") || !g_obj_verify.isParamTypeValid(p_str_end_date, "string") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_start_date = p_str_start_date.trim();
            p_str_end_date = p_str_end_date.trim();
            //允許不帶值，且 p_str_start_date 與 p_str_end_date 也真的為空值，則回傳true
            if (p_str_start_date == "" && p_str_end_date == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            // 判斷是否是正確的日期格式
            //toDo 過程有錯
            var int_sdate_chk = g_obj_verify.isDate(p_str_start_date);
            var int_edate_chk = g_obj_verify.isDate(p_str_end_date);

            if(int_sdate_chk !== g_obj_error_handle.success_code)
            {
                return int_sdate_chk;
            }

            if(int_edate_chk !== g_obj_error_handle.success_code)
            {
                return int_edate_chk;
            }
            // 判斷結束日期是否大於等於開始日期
            if ((new Date(p_str_start_date)) < (new Date(p_str_end_date))) 
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.stardate_and_enddate_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },

    /// <summary>
    /// 驗證 年齡
    /// </summary>
    /// <param name="p_str_age">年齡</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <param name="p_int_age_min">年齡最小值</param>
    /// <param name="p_int_age_max">年齡最大值</param>
    /// <example>
    /// var p_str_age=20;
    /// var p_int_age_min=1;
    /// var p_int_age_max=100;
    /// var p_empty_ok=false;
    /// g_obj_verify.isAgeValid(p_str_age,p_int_age_min,p_int_age_max,p_empty_ok); //return 1
    /// var p_str_age=255;
    /// var p_int_age_min=1;
    /// var p_int_age_max=100;
    /// var p_empty_ok=false;
    /// g_obj_verify.isAgeValid(p_str_age,p_int_age_min,p_int_age_max,p_empty_ok); //return -10
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RayChien] 2009/10/29 Created</remarks>
    isAgeValid : function (p_str_age, p_int_age_min, p_int_age_max, p_empty_ok)  //check OK!
    {
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_age, "string") || !g_obj_verify.isParamTypeValid(p_int_age_min, "number") || !g_obj_verify.isParamTypeValid(p_int_age_max, "number") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_age = p_str_age.trim();
            //允許不帶值，且 p_str_age 也真的為空值，則回傳true
            if (p_str_age == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認年齡範圍是否符合
            if(!(p_int_age_min <= parseInt(p_str_age) && parseInt(p_str_age) <= p_int_age_max))
            {
                return g_obj_error_handle.error_code.input_range_error;
            }
            //確認年齡是否符合數字 regular expression
            if(g_obj_verify.isNumber(p_str_age) == g_obj_error_handle.success_code)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.age_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 即時通帳號
    /// </summary>
    /// <param name="p_str_msn">即時通帳號</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_msn = "fish.bala@gmail.com";
    /// var p_empty_ok=false;
    /// g_obj_verify.isMsnValid(p_str_msn,p_empty_ok);  // ture
    /// var p_str_msn = "fish.bala**ail.com";
    /// var p_empty_ok=false;
    /// g_obj_verify.isMsnValid(p_str_msn,p_empty_ok);  // false
    /// </example>
    /// <returns></returns>
    /// <remarks>[RayChien] 2009/10/29 Created</remarks>
    isMsnValid : function (p_str_msn, p_empty_ok)               //check OK!
    {
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_msn, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_msn = p_str_msn.trim();
            //允許不帶值，且 p_str_msn 也真的為空值，則回傳true
            if (p_str_msn == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認Msn帳號是否符合 email regular expression
            if(g_obj_verify.isEmailAddressValid(p_str_msn) == g_obj_error_handle.success_code)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.msn_id_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    
    /// <summary>
    /// 驗證 是否為中文
    /// </summary>
    /// <param name="p_str_ch">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_ch = "中文";
    /// var p_empty_ok=false;
    /// g_obj_verify.isChinese(p_str_ch,p_empty_ok);  // return 1
    /// var p_str_ch = "english";
    /// var p_empty_ok=false;
    /// g_obj_verify.isChinese(p_str_ch,p_empty_ok);  // return -15
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>
    isChinese : function (p_str_ch, p_empty_ok)      //check OK
    {
        try
        {
            // 中文 regular expression:所有的中文
			// var str_patten = /^[\u4e00-\u9fa5]+$/g;
			// var str_patten = /^[\u0391-\uFFE5]+$/g;
            var str_patten = /^[\u4e00-\u9fa5]+$/g;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_ch, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_ch = p_str_ch.trim();
            //允許不帶值，且 p_str_ch 也真的為空值，則回傳true
            if (p_str_ch == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合中文的 regular expression
            var b_chk_ret = (p_str_ch.match(str_patten) == p_str_ch);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.chinese_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    
    /// <summary>
    /// 驗證 是否為英文
    /// </summary>
    /// <param name="p_str_en">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_en = "english";
    /// var p_empty_ok=false;
    /// g_obj_verify.isEnglish(p_str_en,p_empty_ok);  // return 1
    /// var p_str_en = "中文";
    /// var p_empty_ok=false;
    /// g_obj_verify.isEnglish(p_str_en,p_empty_ok);  // return -17
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>
    isEnglish : function (p_str_en, p_empty_ok)     //check OK!
    {
        //English regular expression:英文字母 a-z or A-z 之間的字母
        var str_patten = /^[A-Za-z]+$/;    
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_en, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_en = p_str_en.trim();
            //允許不帶值，且 p_str_enname 也真的為空值，則回傳true
            if (p_str_en == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合英文 regular expression
			var b_chk_ret = (p_str_en.match(str_patten) == p_str_en);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.english_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
        
    },
    /// <summary>
    /// 驗證 是否為電話
    /// </summary>
    /// <param name="p_str_phone">輸入值</param>
    /// <param name="p_int_phone_min">字數限制</param>
    /// <param name="p_int_phone_max">字數限制</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_phone = "2352366";
    /// var p_int_phone_min = 5;
    /// var p_int_phone_max = 10;
    /// var p_empty_ok=false;
    /// g_obj_verify.isPhone(p_str_phone,p_int_phone_min,p_int_phone_max,p_empty_ok);  // return 1
    /// var p_str_phone = "65456648119914849494948411982";
    /// var p_int_phone_min = 5;
    /// var p_int_phone_max = 10;
    /// var p_empty_ok =false;
    /// g_obj_verify.isPhone(p_str_phone,p_int_phone_min,p_int_phone_max,p_empty_ok);  // return -19
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>
    isPhone : function (p_str_phone, p_int_phone_min , p_int_phone_max , p_empty_ok)   //check OK!
    {
        try
        {
            //phone regular expression
            var str_patten = /^[\+]?[0-9]+$/g ;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_phone, "string") || !g_obj_verify.isParamTypeValid(p_int_phone_min, "number") ||!g_obj_verify.isParamTypeValid(p_int_phone_max, "number"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_phone = p_str_phone.trim();
            //允許不帶值，且 p_str_phone 也真的為空值，則回傳true
            if (p_str_phone == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認電話號碼是否符合範圍
            if(!(p_int_phone_min <= p_str_phone.toString().length && p_str_phone.toString().length <= p_int_phone_max))
            {
				return g_obj_error_handle.error_code.input_length_error;
				//return g_obj_error_handle.success_code;
            }
            //確認是否符合電話的 regular expression
            var b_chk_ret = (p_str_phone.match(str_patten) == p_str_phone);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.phone_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    }, 
    /// 20130124 阿捨新增 大陸手機 or 台灣手機 判斷
    /// <summary>
    /// 驗證 是否為手機電話
    /// </summary>
    /// <param name="p_str_phone">輸入值</param>
    /// <param name="p_int_phone_where">手機類型</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>
    isCellPhone : function (p_str_phone, p_int_phone_where, p_empty_ok)   //check OK!
    {
        try
        {
            //phone regular expression
            var str_patten = /^[\+]?[0-9]+$/g ;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_phone, "string") || !g_obj_verify.isParamTypeValid(p_int_phone_where, "number") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_phone = p_str_phone.trim();
            //允許不帶值，且 p_str_phone 也真的為空值，則回傳true
            if (p_str_phone == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            if ( 1 == p_int_phone_where ){
                p_int_phone_length = 11;
            }
            else if ( 2 == p_int_phone_where ){
                p_int_phone_length = 10;
            }
            //確認電話號碼是否符合範圍
            //if(!(p_int_phone_min <= p_str_phone.toString().length && p_str_phone.toString().length <= p_int_phone_max))
            if( !(p_int_phone_length == p_str_phone.toString().length))
            {
				return g_obj_error_handle.error_code.input_length_error;
				//return g_obj_error_handle.success_code;
            }
            //
            if ( 1 == p_int_phone_where ){
                //alert(p_str_phone.substr(0,2));
                // 20130124 阿捨新增 大陸手機 or 台灣手機 判斷
                if ( "13" == p_str_phone.substr(0,2) || "14" == p_str_phone.substr(0,2) || "15" == p_str_phone.substr(0,2) || "18" == p_str_phone.substr(0,2) )
                {
                }
                else
                {
                    return g_obj_error_handle.error_code.phone_format_error;
                }
            }
            else if ( 2 == p_int_phone_where ){ 
                if ( "09" == p_str_phone.substr(0,2) )
                {
                }
                else
                {
                    return g_obj_error_handle.error_code.phone_format_error;
                }
            }
            //確認是否符合電話的 regular expression
            var b_chk_ret = (p_str_phone.match(str_patten) == p_str_phone);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.phone_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    }, 
    /// <summary>
    /// 驗證 是否為printable
    /// </summary>
    /// <param name="p_var_input">輸入值</param>
    /// <example>
    /// var p_var_input = "0923031356";
    /// g_obj_verify.isPrintable(p_var_input);  // true
    /// var p_var_input = "?";
    /// var p_empty_ok =false;
    /// g_obj_verify.isPrintable(p_var_input,false);  // false
    /// </example>
    /// <returns>
    /// 格式正確:回傳 true
    /// 格式錯誤:回傳 false
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>
    isPrintable  : function (p_var_input)     //check!
    {
        //printable regular expression
        var str_patten = /^[\x21-\x7E]+$/g;
        //確認是否符合Printable regular expression
		var b_chk_ret = (p_var_input.match(str_patten) == p_var_input);
        return b_chk_ret;
    },
    /// <summary>
    /// 驗證 是否為QQ碼
    /// </summary>
    /// <param name="p_var_input">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_var_input = "15651131";
    /// var p_empty_ok =false;
    /// g_obj_verify.isQQid(p_var_input,false);  // return 1
    /// var p_var_input = "61";
    /// var p_empty_ok =false;
    /// g_obj_verify.isQQid(p_var_input,false);  // return -21
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>
    
    isQQid  : function ( p_str_qq, p_empty_ok )     
    {
        try
        { 
            //QQ regular expression
            var str_patten = /^\d{5,10}$/;
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_qq, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_qq = p_str_qq.trim();
            //允許不帶值，且 p_str_qq 也真的為空值，則回傳true
            if (p_str_qq == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合qq regular expression
			var b_chk_ret = (p_str_qq.match(str_patten) == p_str_qq);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.qq_id_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 驗證 是否為生日
    /// </summary>
    /// <param name="p_str_date">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_date = "2006/03/02";
    /// var p_empty_ok =false;
    /// g_obj_verify.isBirthday(p_str_date,false);  // return 1
    /// var p_str_date = "2012/02/09";
    /// var p_empty_ok =false;
    /// g_obj_verify.isBirthday(p_str_date,false);  // return --23
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>

    isBirthday  : function ( p_str_date, p_empty_ok )
    {
        try
        {
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_date, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_date = p_str_date.trim();
            //允許不帶值，且 p_str_qq 也真的為空值，則回傳true
            if (p_str_date == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合qq regular expression
            if(g_obj_verify.isDate(p_str_date) !== g_obj_error_handle.success_code)
            {
                return g_obj_verify.isDate(p_str_date);
            }
            var today=new Date();
            //取得今天的年月日
            var p_str_today=today.getFullYear()+"/"+(parseInt(today.getMonth())+1)+"/"+today.getDate();
            //如果所輸入的日期大於今天的日期，則不為生日
            if(g_obj_verify.isStartAndEndDateValid(p_str_date, p_str_today) == g_obj_error_handle.success_code)
                {
                 return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.birthday_format_error;
				//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 驗證 是否為url
    /// </summary>
    /// <param name="p_str_isUrl">輸入值</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_isUrl = "www.yahoo.com.tw";
    /// var p_empty_ok =false;
    /// g_obj_verify.isUrl(p_str_isUrl,false);  // return 1
    /// var p_str_isUrl = "http://  www.yahoo.cmo.tw";
    /// var p_empty_ok =false;
    /// g_obj_verify.isUrl(p_str_isUrl,false);  // return -22
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>

    isUrl  : function ( p_str_isUrl, p_empty_ok )
    {
        try
        {
            // url regular expression
            var str_patten = /^(http\:\/\/)?([a-z0-9][a-z0-9\-]+\.)?[a-z0-9][a-z0-9\-]+[a-z0-9](\.[a-z]{2,4})+(\/[a-z0-9\.\,\-\_\%\?\=\&]?)?$/g; //url regular expression
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            //確認參數型態是否符合
            if(!g_obj_verify.isParamTypeValid(p_str_isUrl, "string"))
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_isUrl = p_str_isUrl.trim();
            //允許不帶值，且 p_str_isUrl 也真的為空值，則回傳true
            if (p_str_isUrl == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //確認是否符合url的regular expression
			var b_chk_ret = (p_str_isUrl.match(str_patten) == p_str_isUrl);
            if (b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.url_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    },

    /// <summary>
    /// 驗證 是否為form裡的所有欄位是否都填寫
    /// </summary>
    /// <param name="p_var_form">輸入值</param>
    /// <example>
    /*<form name="form1" id="form1" action="" methos="post" onSubmit="return autocheck(g_obj_verify);">
        <input type="text" name="user" valu="1235" />
        <input type="submit" name="submit" value="送出" />
      </form>*/
    /// var p_var_form = document.getElementById('form1');
    /// g_obj_verify.isFormelementsEmpty(p_var_form,false);  // true
    /*<form name="form1" id="form1" action="" methos="post" onSubmit="return autocheck(g_obj_verify);">
        <input type="text" name="user" valu="" />
        <input type="submit" name="submit" value="送出" />
      </form>*/
    /// var p_var_form = document.getElementById('form1');
    /// g_obj_verify.isFormelementsEmpty(p_var_form,false);  // false
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>

    isFormelementsEmpty : function ( p_var_form )   //check OK!
    {
        try
        {
            var check=false;
            // 檢查是否將所有欄位都填寫
            for(var i=0; i < p_var_form.length; i++ )
            {
                // 如果有element為空值，改變check為true
                if( p_var_form.elements[i].value == '' )
                {
                    check=true;
                }
            }
            return check;
        }
        catch(e)
        {
            return e.description;
        }
    },
    /// <summary>
    /// 驗證 傳入參數型態是否正確
    /// </summary>
    /// <param name="p_var_input">輸入值</param>
    /// <param name="p_string_type">參數型態</param>
    /// <example>
    /// var p_var_input = "abc";
    /// var p_string_type="string";
    /// g_obj_verify.isFormelementsEmpty(p_var_input,false);  // true
    /// var p_var_input = 123;
    /// var p_string_type="string";
    /// g_obj_verify.isFormelementsEmpty(p_var_input,false);  // false
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2009/11/26 Created</remarks>
    isParamTypeValid : function (p_var_input, p_string_type)   //check OK!
    {
        return typeof(p_var_input) == p_string_type;
    },
	
	/// <summary>
    /// 驗證 輸入的字元不允許有特殊字元，但允許各國語言的格式
    /// </summary>
    /// <param name="p_str_input">input</param>
    /// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
    /// <example>
    /// var p_str_input = "我是誰";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNoneParticularCharacter(p_str_input,p_empty_ok);     // return 1
    /// var p_str_input = "ff jon@ utorabc.com";
    /// var p_empty_ok=false;
    /// g_obj_verify.isNoneParticularCharacter(p_str_input,p_empty_ok);     // return -13
    /// </example>
    /// <returns>
    /// 格式正確:回傳 1
    /// 格式錯誤:回傳錯誤代碼
    /// </returns>
    /// <remarks>[RyanLin] 2010/02/02 Created</remarks>
    isNoneParticularCharacter : function (p_str_input, p_empty_ok)         //check OK!
    {
        try
        {
            // Email regular expression
            var str_chk_reg_asia = /[\u3400-\uD7FFh|\u00A1-\u1FFFh|\s|a-z|A-Z]+$/g; //檢驗日韓中文空白
			//var str_chk_reg_other = /[\u00A1-\u1FFFh|]+$/g; //檢驗歐洲各國語言
            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            // 判斷傳入參數是否符合為型態要求
            if(!g_obj_verify.isParamTypeValid(p_str_input, "string") )
            {
                return g_obj_error_handle.error_code.param_type_is_invalid;
            }
            //去除字串的前後空白
            p_str_input = p_str_input.trim();
            //允許不帶值，且p_str_input也真的為空值，則回傳true
            if (p_str_input == "")
            {
                if (p_empty_ok)
                {
                    return g_obj_error_handle.success_code;
                }
                return g_obj_error_handle.error_code.input_data_is_empty;
            }
            //若p_str_input符合regular expression回傳true
			var b_chk_ret = (p_str_input.match(str_chk_reg_asia) == p_str_input);
			//var c_chk_ret = (p_str_input.match(str_chk_reg_other) == p_str_input);
            if(b_chk_ret)
            {
                return g_obj_error_handle.success_code;
            }
            return g_obj_error_handle.error_code.email_format_error;
			//return g_obj_error_handle.success_code;
        }
        catch(e)
        {
            return e.description;
        }
    }

/// <summary>
/// 驗證 身分證字號
/// </summary>
/// <param name="p_id_num">身分證字號</param>
/// <param name="p_int_country">國家(預設是1:台灣)</param>
/// <param name="p_empty_ok">允不允許不帶值 option= false or true, default=false</param>
/// <example>
/// var p_id_num="F126007210";
/// var p_int_country=1;
/// var p_empty_ok=false;
/// g_obj_verify.isIdValid(p_id_num,p_int_country,p_empty_ok);  // true
/// </example>
/// <returns>
/// 格式正確:回傳 1
/// 格式錯誤:回傳錯誤代碼
/// </returns>
/// <remarks>[RayChien] 2009/10/29 Created</remarks>
/**************************格式不確定 尚為完成****************************
    isIdValid : function (p_id_num, p_int_country ,p_empty_ok)
    {
        try
        {
            p_id_num = p_id_num.toLowerCase();
            // Taiwan ID regular expression
            var str_patten = /^[a-zA-Z][12][0-9]{8}$/g;
            var str_eng_word = "abcdefghjklmnpqrstuvxywzio";
            var str_tmp, str_chksum;
            var i;

            p_empty_ok = (g_obj_verify.isNull(p_empty_ok)) ? false : p_empty_ok;
            p_id_num = p_id_num.trim();
            if (p_id_num == "")
            {
                if (p_empty_ok)
                {
                    return true;
                }
                return false;
            }
            if(p_int_country == "1")
            {
                if (str_patten.test(p_id_num))
                {
                    str_tmp = 10 + str_eng_word.indexOf(p_id_num.substring(0,1));
                    str_chksum  = (str_tmp-(str_tmp % 10)) / 10 + (str_tmp % 10) * 9;
                    for(i=1; i<9; i++)
                    {
                        str_chksum += p_id_num.substring(i,i+1) * (9-i);
                    }
                    str_chksum = (10 - str_chksum % 10) % 10;
                    if(str_chksum == p_id_num.substring(9,10))
                    {
                        return true
                    }
                }
                else
                {
                    return false;
                }
            }


            if(p_int_country == "2")
            {
                var City={
                    11:"北京",
                    12:"天津",
                    13:"河北",
                    14:"山西",
                    15:"蒙古",
                    21:"遼寧",
                    22:"吉林",
                    23:"黑龍江",
                    31:"上海",
                    32:"江蘇",
                    33:"浙江",
                    34:"安徽",
                    35:"福建",
                    36:"江西",
                    37:"山東",
                    41:"河南",
                    42:"湖北",
                    43:"湖南",
                    44:"廣東",
                    45:"江西",
                    46:"海南",
                    50:"重慶",
                    51:"四川",
                    52:"貴州",
                    53:"雲南",
                    54:"西藏",
                    61:"陝西",
                    62:"甘肅",
                    63:"青海",
                    64:"寧夏",
                    65:"新疆",
                    71:"台灣",
                    81:"香港",
                    82:"澳門",
                    91:"其他"
                };
                var iSum=0;
                //China ID regular expression
                if(!/^\d{17}(\d|x)$/g.test(p_id_num) )
                {
                    return false;
                }
                p_id_num=p_id_num.replace(/x$/g,"a");
                if(City[parseInt(p_id_num.substr(0,2))]==null)
                {
                    return false;
                }
                var sBirthday=p_id_num.substr(6,4)+"-"+Number(p_id_num.substr(10,2))+"-"+Number(p_id_num.substr(12,2));
                var d=new Date(sBirthday.replace(/-/g,"/"))
                if(sBirthday!=(d.getFullYear()+"-"+ (d.getMonth()+1) + "-" + d.getDate()))
                {
                    return false;
                }
                for(var j = 17;j>=0;j --)
                    iSum += (Math.pow(2,j) % 11) * parseInt(p_id_num.charAt(17 - j),11)
                if(iSum%11!=1)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return false;
            }


        }
        catch(e)
        {
            return false;
        }
    },
    ***************************************************************************************/

    
}


