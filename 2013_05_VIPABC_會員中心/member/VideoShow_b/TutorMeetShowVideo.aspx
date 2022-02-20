<%@ Page Language="C#" AutoEventWireup="true" CodeFile="TutorMeetShowVideo.aspx.cs" Inherits="TutorMeetShowVideo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>未命名頁面</title>
<script type="text/javascript" src="http://www.tutormeet.com/js/jcbean.js"></script>
<script type="text/javascript">
	// 20120806 Hunter: CS12061806 - TM 錄影檔連結判斷 CN 或 TW Apache.
    setTutorMeetWebHostName("cn.tutormeet.com");
    
	// 顯示錄影檔
    popupTutorMeetPlayback('<%Response.Write(Request["videoname"]); %>' , '<%Response.Write(Convert.ToInt32(Session["mh_usersn"])); %>');
</script>
</head>
	<body onload='window.close();'>
    <form id="form1" runat="server">
    <div>
    </div>
    </form>
</body>
</html>
