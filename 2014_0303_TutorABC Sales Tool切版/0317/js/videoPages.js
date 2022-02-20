$(function(){
		$(".mix").click(function(){
			 embedVideo('6iEFpGVzwzg');
			 var tempStr=$(this).data( "videoLink" )
			 
		});
		$(".popupWarp").click(function(){
			$(".popupWarp").css("display", "none");
			$('#ABCplayer').remove();
		})
		function embedVideo(str){
			$(".popupWarp").css("display", "block");
			var popupWarpBoxHeight=$(".popupWarpBox").height();
			var windowHeight = $(window).height();
			var offsetY=(windowHeight-popupWarpBoxHeight)/2;
			$(".popupWarpBox").css("top", offsetY);
			$(".popupWarpBox").html("<div id=\"ABCplayer\"><iframe width=\"800\"height=\"600\"frameborder=\"0\"title=\"YouTube video player\"type=\"text/html\"src=\"http://www.youtube.com/embed/"+str+"?enablejsapi=1&autoplay=1\"></iframe</div>");
		}
		
		
});