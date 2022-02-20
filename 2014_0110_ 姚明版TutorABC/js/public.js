
/*获取Cookie值*/
function getCookie(c_name)
{
if(document.cookie.length>0){
   c_start=document.cookie.indexOf(c_name + "=")
   if(c_start!=-1){ 
     c_start=c_start + c_name.length+1 
     c_end=document.cookie.indexOf(";",c_start)
     if(c_end==-1) c_end=document.cookie.length
     return unescape(document.cookie.substring(c_start,c_end))
   }
}
return ""
}
// start 弹出指定大小在线客服窗口
function OpenModalDialog() {
    var client_sn = getCookie('client%5Fsn');
        var w = 650;
        var h = 570;
        var left = (screen.width / 2) - (w / 2) - 20;
        var top = (screen.height / 2) - (h / 2) - 60;

        $.ajax({
            type: "POST",
            cache: false,
            url: "/ajax_getOnlineserviceURLstatus.asp",
            data: {
                "client_sn": client_sn
            },
            success: function (memberType) {
                var targetWin = window.open('http://onlineservice.vipabc.com/VIPABC?clientSn=' + client_sn + '&memberType=' + memberType, 'newwin', 'toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=' + w + ', height=' + h + ', top=' + top + ', left=' + left);
            },
            error: function () {
                var targetWin = window.open('http://onlineservice.vipabc.com/VIPABC', 'newwin', 'toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, copyhistory=no, width=' + w + ', height=' + h + ', top=' + top + ', left=' + left);
            }
        });
        targetWin.focus();
}
// end 弹出指定大小在线客服窗口

function ChangeL(index) {
    var id = "footercert" + index;
    if (document.getElementById(id).src.search(".png") < 0) {
        if (document.getElementById(id).src.search("_grey") > 0) {
            document.getElementById(id).src = document.getElementById(id).src.replace("_grey", "")
        }
    }
    else {
        if (document.getElementById(id).src.search("_1") > 0) {
            document.getElementById(id).src = document.getElementById(id).src.replace("_1", "")
        }
    }

}
function ChangeD(index) {
    var id = "footercert" + index;
    if (document.getElementById(id).src.search(".png") < 0) {
        if (document.getElementById(id).src.search("_grey") < 0) {
            document.getElementById(id).src = document.getElementById(id).src.replace(".jpg", "_grey.jpg")
        }
    }
    else {
        if (document.getElementById(id).src.search("_1") < 0) {
            document.getElementById(id).src = document.getElementById(id).src.replace(".png", "_1.png")
        }
    }
}
