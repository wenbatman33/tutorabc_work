$(document).ready(function () {
    //email autocomplete
    $("#mail").autocomplete({
        search: "",
        minLength: 0,
        source: function (req, responseFn) {
            //如果是ie浏览器设置存放email box的宽度
            if ($.browser.msie) {
                $("ul").addClass("ie-ui-width");
            }
            var re = $.ui.autocomplete.escapeRegex(req.term); //取得文本框内容
            //alert(re);
            var n = re.indexOf("@");
            var newre = re;

            if (n != -1) {  //if "@" is found in string, get 0 to n to exclude "@" and everything after it
                newre = re.substring(0, n)
            }
            //alert(newre);
            var wordlist = [newre + "@sina.com", newre + "@163.com", newre + "@qq.com", newre + "@126.com", newre + "@vip.sina.com", newre + "@sina.cn", newre + "@hotmail.com", newre + "@gmail.com", newre + "@sohu.com", newre + "@yahoo.com", newre + "@139.cn", newre + "@189.cn"];
            var matcher = new RegExp("^" + re, "i");
            var a = $.grep(wordlist, function (item, index) {
                return matcher.test(item);
            });

            if ($.browser.msie && ($.browser.version == "6.0")) {
                if (a.length == 0) {
                    $("select").show(); //显示下拉
                }
                else {
                    $("select").hide(); //隐藏下拉
                }
                $("#mail").blur(function () {
                    $("select").show(); //显示下拉
                });
            }
            responseFn(a);
        },
        select: function (event, ui) {
            $("#mail").val(ui.item.value); //选择正确后向文本框赋值
            $("select").show(); //显示下拉
        }
    });

    $("#email").autocomplete({
        search: "",
        minLength: 0,
        source: function (req, responseFn) {
            //如果是ie浏览器设置存放email box的宽度
            if ($.browser.msie) {
                $("ul").addClass("ie-ui-width");
            }
            var re = $.ui.autocomplete.escapeRegex(req.term); //取得文本框内容
            //alert(re);
            var n = re.indexOf("@");
            var newre = re;

            if (n != -1) {  //if "@" is found in string, get 0 to n to exclude "@" and everything after it
                newre = re.substring(0, n)
            }
            //alert(newre);
            var wordlist = [newre + "@sina.com", newre + "@163.com", newre + "@qq.com", newre + "@126.com", newre + "@vip.sina.com", newre + "@sina.cn", newre + "@hotmail.com", newre + "@gmail.com", newre + "@sohu.com", newre + "@yahoo.com", newre + "@139.cn", newre + "@189.cn"];
            var matcher = new RegExp("^" + re, "i");
            var a = $.grep(wordlist, function (item, index) {
                return matcher.test(item);
            });

            if ($.browser.msie && ($.browser.version == "6.0")) {
                if (a.length == 0) {
                    $("select").show(); //显示下拉
                }
                else {
                    $("select").hide(); //隐藏下拉
                }
                $("#email").blur(function () {
                    $("select").show(); //显示下拉
                });
            }
            responseFn(a);
        },
        select: function (event, ui) {
            $("#email").val(ui.item.value); //选择正确后向文本框赋值
            $("select").show(); //显示下拉
        }
    });
    $("#txt_email").autocomplete({
        search: "",
        minLength: 0,
        source: function (req, responseFn) {
            //如果是ie浏览器设置存放email box的宽度
            if ($.browser.msie) {
                $("ul").addClass("ie-ui-width");
            }
            var re = $.ui.autocomplete.escapeRegex(req.term); //取得文本框内容
            //alert(re);
            var n = re.indexOf("@");
            var newre = re;

            if (n != -1) {  //if "@" is found in string, get 0 to n to exclude "@" and everything after it
                newre = re.substring(0, n)
            }
            //alert(newre);
            var wordlist = [newre + "@sina.com", newre + "@163.com", newre + "@qq.com", newre + "@126.com", newre + "@vip.sina.com", newre + "@sina.cn", newre + "@hotmail.com", newre + "@gmail.com", newre + "@sohu.com", newre + "@yahoo.com", newre + "@139.cn", newre + "@189.cn"];
            var matcher = new RegExp("^" + re, "i");
            var a = $.grep(wordlist, function (item, index) {
                return matcher.test(item);
            });

            if ($.browser.msie && ($.browser.version == "6.0")) {
                if (a.length == 0) {
                    $("select").show(); //显示下拉
                }
                else {
                    $("select").hide(); //隐藏下拉
                }
                $("#txt_email").blur(function () {
                    $("select").show(); //显示下拉
                });
            }
            responseFn(a);
        },
        select: function (event, ui) {
            $("#txt_email").val(ui.item.value); //选择正确后向文本框赋值
            $("select").show(); //显示下拉
        }
    });

});

$(document).ready(function () {
    //email autocomplete
    $("#txt_email_addr").autocomplete({
        search: "",
        minLength: 0,
        source: function (req, responseFn) {
            //如果是ie浏览器设置存放email box的宽度
            if ($.browser.msie) {
                $("ul").addClass("ie-ui-width");
            }
            var re = $.ui.autocomplete.escapeRegex(req.term); //取得文本框内容
            //alert(re);
            var n = re.indexOf("@");
            var newre = re;

            if (n != -1) {  //if "@" is found in string, get 0 to n to exclude "@" and everything after it
                newre = re.substring(0, n)
            }
            //alert(newre);
            var wordlist = [newre + "@sina.com", newre + "@163.com", newre + "@qq.com", newre + "@126.com", newre + "@vip.sina.com", newre + "@sina.cn", newre + "@hotmail.com", newre + "@gmail.com", newre + "@sohu.com", newre + "@yahoo.com", newre + "@139.cn", newre + "@189.cn"];
            var matcher = new RegExp("^" + re, "i");
            var a = $.grep(wordlist, function (item, index) {
                return matcher.test(item);
            });

            if ($.browser.msie && ($.browser.version == "6.0")) {
                if (a.length == 0) {
                    $("select").show(); //显示下拉
                }
                else {
                    $("select").hide(); //隐藏下拉
                }
                $("#txt_email_addr").blur(function () {
                    $("select").show(); //显示下拉
                });
            }
            responseFn(a);
        },
        select: function (event, ui) {
            $("#txt_email_addr").val(ui.item.value); //选择正确后向文本框赋值
            $("select").show(); //显示下拉
        }
    });
});