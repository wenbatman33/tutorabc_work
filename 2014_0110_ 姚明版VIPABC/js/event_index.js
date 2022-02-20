$(function() {
	var tag;
	$(".iframe").colorbox({iframe:true, innerWidth: 692, innerHeight: 431 ,fixed:true});
	$(".youtube").colorbox({iframe:true, Width:743, Height: 558, innerWidth:743, innerHeight:550,fixed:true});
	$(".youku").colorbox({iframe:true,  innerWidth:743, innerHeight:550,fixed:true,
							onComplete: function(){ 
							var addressValue = $(this).attr("href");
							Complete(addressValue);
							}
    });
	
	//固定 landing from
	var position = $('.landing').position();
	var sticky_navigation_offset_top =  position.top;
	var upLine=sticky_navigation_offset_top- $('.topNavFixed').height();
	var speed=500;
    var sticky_navigation = function () {
        var scroll_top = $(window).scrollTop();
		var bottomLine=$(document).height()- $('.footer').height()-sticky_navigation_offset_top;
		var scroll_bottom = $(document).height() - $('.landingForm').height()-sticky_navigation_offset_top - $('.footer').height();
		
    if (scroll_top > bottomLine) {
        $('.landingForm').stop().animate({'top':scroll_bottom-20 },speed );
        $('.CustomerServiceIcon').stop().animate({'top':scroll_bottom-20 },speed );

    }else if (scroll_top > upLine) {
        $('.landingForm').stop().animate({'top':scroll_top-580},speed );
        $('.CustomerServiceIcon').stop().animate({'top':scroll_top-580},speed );
    }else{
       $('.landingForm').stop().animate({'top':0},speed );
       $('.CustomerServiceIcon').stop().animate({'top':0},speed );
    }
      
    };
     sticky_navigation();
        

    var timeoutId;
    $(window).scroll(function(){
    if(timeoutId ){
    clearTimeout(timeoutId );  
    }
    timeoutId = setTimeout(function(){  sticky_navigation();  }, 200);

    });


});	
function Complete(str){
	if (/(iPhone|iPad|iPod|iOS)/i.test(navigator.userAgent)) {
		tag="<video width=\"743\" height=\"550\" src=\"http://v.youku.com/player/getRealM3U8/vid/"+str+"/type//video.m3u8\"  autoplay=\"autoplay\" controls=\"controls\"></video>";
		$("#cboxLoadedContent").html(tag)
	} else {  
		tag="<embed src=\"http://player.youku.com/player.php/sid/"+str+"/v.swf\" allowfullscreen=\"false\" quality=\"high\" width=\"743\" height=\"550\" align=\"middle\" type=\"application/x-shockwave-flash\" flashvars=\"winType=index&amp;isAutoPlay=true\" style=\"visibility:visible\">";
		$("#cboxLoadedContent").html(tag)
	}  
} 