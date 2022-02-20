<%@ Page Language="C#" AutoEventWireup="true" CodeFile="VideoInfoForm.aspx.cs" Inherits="VideoInfoForm" %>

<%@ Register src="UI/WebUserContro_RatingAndViewPanel.ascx" tagname="WebUserContro_RatingPanel" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>VideoInfo</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="">
        <uc1:WebUserContro_RatingPanel ID="WebUserContro_RatingPanel1" runat="server" />
    </div>
    </form>
</body>
</html>
