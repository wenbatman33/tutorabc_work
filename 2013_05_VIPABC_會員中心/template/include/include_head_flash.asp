<!--#include virtual="/backoffice/inc/conn.asp"-->
<!--flash start-->
<div class="flash-all">
<div class="flash">
    <div id="flash_main">	   
    <a name="bannerLink" href=""><span id="bannerImage"></span></a>
	</div>    
	<script type="text/javascript">
		<%
		dim url : url = Request.ServerVariables("PATH_INFO")		
		%>
		var url ="<%=url%>"
		var bannerImage = $("span#bannerImage");
		var bannerLink=$("a[name=bannerLink]");
        var strCookiePathName = "FlashPath";
        var options = { path: '/', expires: 10 };
        //連結的路徑
        var aLinkForMenu = "<%=getWebSiteFullPath()%>";
        //清空设置		
        $.cookie(strCookiePathName, '', options);
        var intHeight = 170;	
		//alert(url);	
		//<!--2013/04/10 lewisli 官网字页Flash动态切换及底层防止flash插件无效-->
		<%
			'set rs=server.CreateObject("adodb.recordset")
			'sql="select b.BannerPicture,b.BannerSwfPic,b.BannerUrl, s.SubMenuAddress from Banner b inner join SubMenus s on b.SubMenuId =s.SubMenuId where b.BannerStatus =1 and b.SubMenuId <> 1"  
			'rs.Open sql,conn,1,1	
			%>
				<% 'do while not rs.EOF %>   			
					//关于我们program/about/index.asp						
                    if(url=='<%'=rs("SubMenuAddress")%>')
					{									
						bannerImage.html("<img src='/VIPABCMaintain/Content/BannerImg/<%'=rs("BannerPicture")%>\'>");
						bannerLink.attr("href","<%'=rs("BannerUrl")%>");  
						$.cookie(strCookiePathName,  '<%'=rs("BannerSwfPic")%>', options);  
					}					 				
					<% 'rs.movenext
					   'loop%>					  
			<%				
			'rs.close
			'set rs=nothing
			%>    
			//alert($.cookie(strCookiePathName));			
         if ($.cookie(strCookiePathName) != null && $.cookie(strCookiePathName)!="") 
        {          		
			strFlashPath = "/VIPABCMaintain/Content/Bannerswf/" + $.cookie(strCookiePathName);			
        }else 
		{		
			//取得cookie來決定檔案的路徑，預設為1
			var strFlashPath = "/VIPABCMaintain/Content/Bannerswf/banner_1.swf";	
			bannerImage.html("<img src=\"/VIPABCMaintain/Content/BannerImg/banner_1.jpg\">");
			bannerLink.attr("href","http://www.vipabc.com/count.asp?code=WCVosPc7WB");  //每天45分钟轻松在家学英语		
		}
		//<!--2013/04/10 lewisli 官网字页Flash动态切换及底层防止flash插件无效-->
        var fo = new FlashObject(strFlashPath, "flash_main", "960",  intHeight, "7", "#000000"); 		
        fo.addParam("wmode", "transparent");
        fo.addParam("allowFullscreen","true");
        fo.write("flash_main");
    </script>
</div>
</div>
<!--flash end-->