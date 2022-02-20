<TITLE>SaveAspSession.aspx</TITLE>
<%@ Page language="c#" %>
<script runat=server>
// We iterate through the Form collection and assign the names and values
// to ASP.NET session variables! We have another Session Variable, "DestPage"
// that tells us where to go after taking care of our business...
    private void Page_Load(object sender, System.EventArgs e)
    {
        if(Request.Form.Count == 0 )
            Response.Redirect("http://www.vipabc.com/program/member/member_login.asp");
        //for (int i = 0; i < Request.Form.Count; i++)
        //{
            //Session[Request.Form.GetKey(i)] = Request.Form[i].ToString();
        //}
		Session[Request.Form.GetKey(0)] = Request.Form[0].ToString();
        //Server.Transfer("VideoShow.aspx", true);
        Response.Redirect("VideoShow.aspx");
    }
</script>