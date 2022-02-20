$(function() {
		
	$(".language").click(function() {
		$(".dropdown-menu").toggle();
	});
	
	$(".vipabc_iframe").fancybox({
		'width'				: 800,
		'height'			    : 520,
		'autoScale'			: false,
		'transitionIn'		: 'none',
		'transitionOut'		: 'none',
		'type'				: 'iframe'
	});
		
		
	////固定nav bar
	var sticky_navigation_offset_top = 350;
	var banners_offset_top = $('#banners').offset().top;
	
	
	var sticky_navigation = function(){
	var scroll_top = $(window).scrollTop(); 
	
	if (scroll_top > sticky_navigation_offset_top) { 
		$('.topNavFixed').fadeIn('fast');
		$('.topNavFixed').css({'position': 'fixed', 'top':0, 'left':0 });
	} else {
		$('.topNavFixed').fadeOut(0);
		$('.topNavFixed').css({'position': 'absolute', 'top':sticky_navigation_offset_top, 'left':0 });
	}
	
	
	if (scroll_top > sticky_navigation_offset_top) { 
		$('.regFixedWrap').css({'position': 'fixed', 'top':180 });
		$('.regForm').css({'right': '0px', 'margin-right':'0px' });
		
	} else {
		$('.regFixedWrap').css({'position': 'relative' ,'top':580}); 
		$('.regForm').css({'right': '0px', 'margin-right':'0px' });
	}
	
	
	var scroll_bottom = $(document).height() -700 - $('.footer').height() ;

		if (scroll_top > scroll_bottom) { 
			$('.regFixedWrap').fadeOut(300);
		} else {
			$('.regFixedWrap').fadeIn(300);
		}   
			
	if (scroll_top > sticky_navigation_offset_top) { 
		$('.backTop').css({ 'bottom': 0 });
	} else {
		$('.backTop').css({ 'bottom': 999999 });
	}	
	};
	
	sticky_navigation();
	////SmoothScroll 項目 還需要再調整
	var nav_height  = 123;
	var part_01_top = Number($("#part_01").offset().top-nav_height-5);
	var part_02_top = Number($("#part_02").offset().top-nav_height-5);
	var part_03_top = Number($("#part_03").offset().top-nav_height-5);
	var part_04_top = Number($("#part_04").offset().top-nav_height-5);
	var part_05_top = Number($("#part_05").offset().top-nav_height-5);
	var part_06_top = Number($("#part_06").offset().top-nav_height-5);
	
	
	
	var part_01_bottom = Number(part_01_top+$("#part_01").height());
	var part_02_bottom = Number(part_02_top+$("#part_02").height());
	var part_03_bottom = Number(part_03_top+$("#part_03").height());
	var part_04_bottom = Number(part_04_top+$("#part_04").height());
	var part_05_bottom = Number(part_05_top+$("#part_05").height());
	var part_06_bottom = Number(part_06_top+$("#part_06").height());
	
	
	var part_02_Switch =  false;
	
	$('.index_list_wrap').hover(function(){
		$(this).find('.index_list_on').stop().animate({
					top: 112
				}, 500);
			}, function(){
				$(this).find('.index_list_on').stop().animate({
					top: 300
				}, 500);
	});
	
	

	$(window).scroll(function(){
			////固定nav bar
			sticky_navigation();
			////固定nav bar
			// 捲動的數字
			var scroH = Number($(this).scrollTop());
	
			 
			$("#album").css("background-position", "50% "+(0+(scroH*0.4))+"px");
			// 捲動的數字
			if(scroH>= part_01_top && scroH <=part_01_bottom ){
			set_cur(1);
			}else if(scroH>= part_02_top && scroH <=part_02_bottom){
			set_cur(2);
			if(part_02_Switch ==  false){
				part_02_Switch =  true;
	
			}

			}else if(scroH>= part_03_top && scroH <=part_03_bottom){
			set_cur(3);
			}else if(scroH>= part_04_top && scroH <=part_04_bottom){
			set_cur(4);
			}else if(scroH>= part_05_top && scroH <=part_05_bottom){
			set_cur(5);
			}else if(scroH>= part_06_top && scroH <=part_06_bottom){
			set_cur(6);
		}
	});
	
	
	
	
	
	
	$(".topNav a").click(function() {
		var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');
		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 );
				}
		}
		
 	});
	
	$(".topNav a").bind( "touchstart", function(e){
		var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');
		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 );
				}
		}
	});
	
	$(".main_nav_box a") .click(function() {
		var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');
		
		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
					
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 );
				}
		}
		
 	});
	
	$(".main_nav_box a").bind( "touchstart", function(e){
		var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');

		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 );
				}
		}
		
	});
	
	
	// 回首頁頂天處
	$(".index_top").click(function() {
		var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
			$body.stop().animate({	scrollTop: 0}, 1000 );
		});
		
	$(".index_top").bind( "touchstart", function(e){
		var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
			$body.stop().animate({	scrollTop: 0}, 1000 );
		});

	
	
	
	
});

	
	function set_cur(n){
	 var str = ":nth-child("+n+")";
		if($(".topNav a").hasClass("cur")){
			$(".topNav a").removeClass("cur");

		}
		$(".topNav a"+str).addClass("cur");

	}
	function pageScroll(n){
	
	}
	
	
	
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
	
	
	
	
	

	