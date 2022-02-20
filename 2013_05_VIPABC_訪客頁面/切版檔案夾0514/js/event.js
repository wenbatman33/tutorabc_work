$(function() {
	////固定nav bar
	var sticky_navigation_offset_top = $('#top_nav').offset().top;
	var banners_offset_top = $('#banners').offset().top;
	
	var sticky_navigation = function(){
		var scroll_top = $(window).scrollTop(); 
		if (scroll_top > sticky_navigation_offset_top) { 
			$('#top_nav').css({'position': 'fixed', 'top':0, 'left':0 });
		} else {
			$('#top_nav').css({'position': 'relative' }); 
		}   
		if (scroll_top > 300) { 
			$('.fixed_box').css({'position': 'fixed', 'top':80 });
		} else {
			$('.fixed_box').css({'position': 'relative' ,'top':500}); 
		}   
	};
	
	sticky_navigation();
	////SmoothScroll 項目 還需要再調整
	
	var nav_height  = 80;
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
	

	$(window).scroll(function(){
			////固定nav bar
			sticky_navigation();
			////固定nav bar
			
			// 捲動的數字
			var scroH = Number($(this).scrollTop()+nav_height);
			// 捲動的數字
			
			if(scroH> part_01_top && scroH <=part_01_bottom ){
			set_cur(".part_01");
			}else if(scroH>= part_02_top && scroH <=part_02_bottom){
			set_cur(".part_02");
			
			
			
			if(part_02_Switch ==  false){
				
			
				part_02_Switch =  true;
				
			$("#rotate-x .panel").css('-webkit-transform', 'rotateX( 0deg )',
						   '-moz-transform', 'rotateX( 0deg )',
							 '-o-transform', 'rotateX( 0deg )',
								'transform', 'rotateX( 0deg )');
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
	
	
	$(".nav_bar li a").click(function() {
		var el = "#"+$(this).attr('class');
		
		// 如果 捲軸頂天的話 要扣除 topNews的高度
		if($(window).scrollTop()<48){
			var nav_height  = 160;
		}else{
			// 如果 捲沒有軸頂天的話 只要扣除 topNav的高度
			var nav_height  = 80;
	    }
		
		var num =Number($(el).offset().top-nav_height);

		
		var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
			$body.animate({
				scrollTop: num
			}, 1000 , "easeOutExpo");
			
			
 	});
	
	
	$(".index_top").click(function() {
	
	var $body = (window.opera) ? (document.compatMode == "CSS1Compat" ? $('html') : $('body')) : $('html,body');
			$body.animate({
				scrollTop: 0
			}, 1000 , "easeOutExpo");
			
			
 	});

});




function pageScroll(obj){
	
	

}


function set_cur(n){
	if($(".nav_bar a").hasClass("cur")){
		$(".nav_bar a").removeClass("cur");
	}
	$(".nav_bar a"+n).addClass("cur");
}

