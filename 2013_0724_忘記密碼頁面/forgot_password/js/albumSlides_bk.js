$(function(){
		// 先取得必要的元素並用 jQuery 包裝
		// 再來取得 $block 的寬度及設定動畫時間
		$('#album ul li').hide();
		
		var $block = $('#album'),
			 $slides = $block.find('.slides'),
			 $ul = $slides.find('ul'),
		   	 _width = $slides.width(),
		 	 _left = _width * -1,
			_animateSpeed = 400,
			totalPic=$('.slides li').length,
			oldNum=0,
			nowNum=0,
			// 加入計時器, 輪播時間及控制開關
			timer, _showSpeed = 5000, _stop = false;
	
		// 當點擊到 a.prev 時
		 $('.banner_prev').click(function(){
			prv();
		});
		
		// 當點擊到 a.next 時
		$('.banner_next').click(function(){
			next();
		});
		
		
		$('.new_ad_btns li').click(function(){
			
			clearTimeout(timer);
			var index = $('.new_ad_btns li').index($(this))
			//alert(index);
			
			nowNum=index;
		     roll(index) ;
			 
		});
	
		
		function next(){
			clearTimeout(timer);
			nowNum+=1;
			if(nowNum>=totalPic){
				nowNum=0;
			}
			roll(nowNum);
		}
		
		function prv(){
			clearTimeout(timer);
			nowNum-=1;
			if(nowNum<0){
				nowNum=totalPic-1;
			}
			roll(nowNum);
		}
		
		
		
		function roll(nowNum ){
			var $oldItem = $('#album ul li').eq(oldNum);
			var $newItem = $('#album ul li').eq(nowNum);
			
			
			var $hitlinks = $('.new_ad_btns li').eq(nowNum).find('a');
			
			$('.new_ad_btns li a').css('color', '#999');
			
			$hitlinks.css('color', '#fff');
			
			
			var bgimg='url('+ $newItem.attr("src")+')';
			//淡入淡出背景
			var image = $("#album");
			image.fadeOut(100, function () {
				image.css("background", bgimg );
				image.fadeIn(400);	
			});
			
			
			$oldItem.stop().fadeOut(300);
			$newItem.stop().fadeIn(500);
			
			oldNum=nowNum;
			timer = setTimeout(move, _showSpeed);
			timeLineMove()
		
			//myTimeLine= setTimeout(timeLineMove, 100);
		}
		
		
		// 計時器使用
		function move(){
			next();
		};
		
		function timeLineMove(){
			
			var part_width=nowNum*(960/totalPic);
			var goingwidth=(nowNum+1)*(960/totalPic)+"px";
			
			$(".timeline").css('width',part_width);
			
			$(".timeline").stop().animate({"width":goingwidth}, 5000);
			
		};
		
		
		roll(0);
		
	});