$(function(){
	$("#album").css({"background":  $('#album ul li').attr("src")});
		var $block = $('#album'),
		$slides = $block.find('.slides'),
		$ul = $slides.find('ul'),
		_width = $slides.width(),
		_left = _width * -1,
		_animateSpeed=1000,
		bgSpeed=200,
		totalPic=$('.slides li').length,
		oldNum=0,
		nowNum,
		myType=true,
		timer, _showSpeed = 8000, _stop = false, btnOver=false;
		
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
				
				if(navigator.appName == "Microsoft Internet Explorer" && navigator.appVersion .split(";")[1].replace(/[ ]/g,"")=="MSIE7.0"){
					myType=false;
				}else if(navigator.appName == "Microsoft Internet Explorer" && navigator.appVersion .split(";")[1].replace(/[ ]/g,"")=="MSIE8.0"){ 
					myType=false;
				} else{
					myType=true;
				}
				
				if(myType==false){
	
							image.css({
							"background": bgimg,
							"background-position": "center"
							});
							
							$('#album ul li').css({
							"left": 1200,
							"display":"none"
							});
							
							$newItem.css({
							"left": 0,
							"display":"list-item"
							});
				}else{
					
							image.fadeOut(bgSpeed, function () {
							image.css({
							"background": bgimg,
							"background-position": "center"
							});
							image.fadeIn(_animateSpeed);	
							});
							$('#album ul li').stop().fadeOut(bgSpeed);
							$newItem.stop().fadeIn(_animateSpeed);
					
					
				}
				////////////////////////////////////////////////
				oldNum=nowNum;
				timeLineMove();
		}
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