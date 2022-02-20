$(function() {
	
	
	//// 增加nav bar 底線
    $(".nav_bar").append("<li id='magic-line'></li>");

	//// 增加nav bar 滑動效果
    var $magicLine = $("#magic-line");
    var $end_x = 0;
	
	var $end_width;
	
    $magicLine.hide();
    $magicLine
        .width($(".current_page_item").width())
        .css("left", $end_x)
        .data("origLeft", $end_x)
        .data("origWidth", $end_width);
        
    $(".nav_bar li").find("a").hover(function() {
        $el = $(this);
        leftPos = $el.position().left;
        newWidth = $el.parent().width();

	$magicLine.show();
		
        $magicLine.stop().animate({
            left: leftPos,
            width: newWidth
        });
    }, function() {
		$magicLine.hide();
        $magicLine.stop().animate({
            left: $magicLine.data("origLeft"),
            width: $magicLine.data("origWidth")
        });    
    });
	//// 增加nav bar 滑動效果
	
	
	////固定nav bar
	var sticky_navigation_offset_top = 350;
	var banners_offset_top = $('#banners').offset().top;
	
	
	var sticky_navigation = function(){
	var scroll_top = $(window).scrollTop(); 
	
	if (scroll_top > sticky_navigation_offset_top) { 
		$('#top_nav_index').fadeIn('fast');
		$('#top_nav_index').css({'position': 'fixed', 'top':0, 'left':0 });
	} else {
		$('#top_nav_index').fadeOut(0);
		$('#top_nav_index').css({'position': 'absolute', 'top':sticky_navigation_offset_top, 'left':0 });
	}
	
	
	if (scroll_top > sticky_navigation_offset_top) { 
		$('.fixed_box').css({'position': 'fixed', 'top':180 });
		$('.registration_form').css({'right': '0px', 'margin-right':'0px' });
		
	} else {
		$('.fixed_box').css({'position': 'relative' ,'top':580}); 
		$('.registration_form').css({'right': '0px', 'margin-right':'0px' });
	}
	
	
	var scroll_bottom = $(document).height() -600- $('.footer').height() ;
		if (scroll_top > scroll_bottom) { 
			$('.fixed_box').fadeOut(300);
		} else {
			$('.fixed_box').fadeIn(300);
		}   
		
		
	if (scroll_top > sticky_navigation_offset_top) { 
		$('.backTop').css({ 'bottom': 0 });
	} else {
		$('.backTop').css({ 'bottom': 999999 });
	}
	
	
	
		
	};
	
	sticky_navigation();
	////SmoothScroll 項目 還需要再調整
	
	var nav_height  = 120;
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
			if(scroH> part_01_top && scroH <=part_01_bottom ){
			set_cur(".part_01");
			}else if(scroH>= part_02_top && scroH <=part_02_bottom){
			set_cur(".part_02");
			if(part_02_Switch ==  false){
				part_02_Switch =  true;
	
			}

			}else if(scroH>= part_03_top && scroH <=part_03_bottom){
			set_cur(".part_03");
			}else if(scroH>= part_04_top && scroH <=part_04_bottom){
			set_cur(".part_04");
			}else if(scroH>= part_05_top && scroH <=part_05_bottom){
			set_cur(".part_05");
			}else if(scroH>= part_06_top && scroH <=part_06_bottom){
			set_cur(".part_06");
		}
	});
	$(window).scroll();
	
	$(".nav_bar li a").click(function() {
		var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');
		
		
		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 , "easeOutExpo");
				}
		}
		
 	});
	
	$(".nav_bar li a").bind( "touchstart", function(e){
		
				var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');
	
		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 , "easeOutExpo");
				}
		}
		
	});
	
	
	
	$(".main_nav_ul li a") .click(function() {
		var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');
		
		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
					
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 , "easeOutExpo");
				}
		}
		
 	});
	
	$(".main_nav_ul li a").bind( "touchstart", function(e){
		var el = "#"+$(this).attr('class');
		var clicked=$(this).hasClass('cur');

		if( clicked!=true){	
				if(el !='#part_07'){
						var num = Number($(el).offset().top-nav_height);
					
						var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
						$body.stop().animate({	scrollTop: num}, 1000 , "easeOutExpo");
				}
		}
		
	});
	
	// 回首頁頂天處
	$(".index_top").click(function() {
		var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
			$body.stop().animate({	scrollTop: 0}, 1000 , "easeOutExpo");
		});
		
	$(".index_top").bind( "touchstart", function(e){
		var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
			$body.stop().animate({	scrollTop: 0}, 1000 , "easeOutExpo");
		});
		
		
	});
	
	
	
	
	// 回首頁頂天處

	function set_cur(n){
		if($(".nav_bar a").hasClass("cur")){
			$(".nav_bar a").removeClass("cur");
			$(".nav_bar a").parent().removeClass("current_page_item");
		}
		$(".nav_bar a"+n).addClass("cur");
		$(".nav_bar li"+n).parent().addClass("current_page_item");
	
		
		$magicLine = $("#magic-line");
		$magicLine.hide();
		
	}
	
	function pageScroll(n){
	
	}



