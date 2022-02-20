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
			nowNum,
			// 加入計時器, 輪播時間及控制開關
			timer, _showSpeed = 8000, _stop = false, btnOver=false;
	
		// 當點擊到 a.prev 時
		 $('.banner_prev').click(function(){
			prv();
		});
		
		// 當點擊到 a.next 時
		$('.banner_next').click(function(){
			next();
		});
	
	
		
		$('.new_ad_btns li').hover(
			function(){
				btnOver=true;
				$(".timeline").stop();
				var index = $('.new_ad_btns li').index($(this));
				 roll(index);
			},function(){
				btnOver=false;
				timeLineMove()
	
			}
		);
		
		///滑鼠滑到大banner停止時間
		/*  
		$('.slides ul li a').hover(
			function(){
				$(".timeline").stop();
				btnOver=true;
			},function(){
				btnOver=false;
				timeLineMove()
	
			}
		);*/
		
		function next(){
			nowNum+=1;
			if(nowNum>=totalPic){
				nowNum=0;
			}
			roll(nowNum);
		}
		
		function prv(){
			nowNum-=1;
			if(nowNum<0){
				nowNum=totalPic-1;
			}
			roll(nowNum);
		}
		
		
		
		function roll(num){
		
				nowNum=num;
				
				var $oldItem = $('#album ul li').eq(oldNum);
				var $newItem = $('#album ul li').eq(nowNum);
				
				var $hitlinks = $('.new_ad_btns li').eq(nowNum).find('a');
				
				$('.new_ad_btns li a').css('color', '#999');
				$hitlinks.css('color', '#fff');
				
				var bgimg='url('+ $newItem.attr("src")+')';
		
				var image = $("#album");
				////////////////////////////////////////////////
				image.fadeOut(100, function () {
	
					image.css({
						"background": bgimg,
						"background-position": "center"
					  });
					
	
					
					//image.css({"background": bgimg,"filter": "progid:DXImageTransform.Microsoft.AlphaImageLoader(src='" + img.attr("src") + "', sizingMethod='scale')"});
					
					image.fadeIn(400);	
				});
				
				$('#album ul li').stop().fadeOut(300);
				$newItem.stop().fadeIn(500);
				////////////////////////////////////////////////	
				
			
				oldNum=nowNum;
				timeLineMove();
	
			
					
			
		
					
		}
	
		
		
		// 計時器使用
		function move(){
			next();
		};
		
		
		function reRoll(){
			oldNum=nowNum;
			timeLineMove();
		};
	
		function timeLineMove(){
			
			var part_width=nowNum*(960/totalPic);
			var goingwidth=(nowNum+1)*(960/totalPic)+"px";
			
			$(".timeline").css('width',part_width);
			
			
			 if(btnOver==false){
				$(".timeline").stop().animate({"width":goingwidth}, _showSpeed, function() {
						next();
				 });
			 }
			
		};
		
		roll(0);
	
	
		
	});