$(function(){
		// 先取得必要的元素並用 jQuery 包裝
		// 再來取得 $block 的寬度及設定動畫時間
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
			timer, _showSpeed = 3000, _stop = false;

		// 當點擊到 a.prev 時
		var $prev = $block.find('a.prev').click(function(){
			prv();
		});
		
		// 當點擊到 a.next 時
		var $next = $block.find('a.next').click(function(){
			next();
		});
		
		$('.Thumb').click(function(){
			var indexNum= $(this).index();
			nowNum=indexNum;
			roll(nowNum ,'Right');
		});
		
		function next(){
			nowNum+=1;
			if(nowNum>=totalPic){
				nowNum=0;
			}
			roll(nowNum,'Right');
		}
		function prv(){
			nowNum-=1;
			if(nowNum<=0){
				nowNum=totalPic-1;
			}
			roll(nowNum,'Left');
		}
		
		
		
		function roll(nowNum ,Direction){
			$('#album ul li').stop().animate({left: 9990}, 0, function() {});
			var $oldItem = $('#album ul li').eq(oldNum);
			var $newItem = $('#album ul li').eq(nowNum);
			$oldItem.stop().animate({left: 0}, 0, function() {});
			
			if(oldNum==nowNum){
				//$newItem.stop().animate({left: 0}, 0, function() {});
			}
			
			if(Direction =='Left'){
				$newItem.stop().animate({left: 0-_width}, 0, function() {});
				$oldItem.stop().animate({left: _width}, 1000, function() {});
				$newItem.stop().animate({left: 0}, 1000, function() {});
				
				
			}else if(Direction =='Right'){
				$newItem.stop().animate({left: _width}, 0, function() {});
				
				$oldItem.stop().animate({left: 0-_width}, 1000, function() {});
				$newItem.stop().animate({left: 0}, 1000, function() {});
			}		
			oldNum=nowNum;
		
		}
		
		roll(0);
		// 如果滑鼠移入 $block 時
		
		$('.Thumb').hover(function(){
			// 關閉開關及計時器
			_stop = true;
			clearTimeout(timer);
		}, function(){
			// 如果滑鼠移出 $block 時
			// 開啟開關及計時器
			_stop = false;
			timer = setTimeout(move, _showSpeed);
		}).find('a').focus(function(){
			this.blur();
		});
		
		
		$block.hover(function(){
			// 關閉開關及計時器
			_stop = true;
			clearTimeout(timer);
		}, function(){
			// 如果滑鼠移出 $block 時
			// 開啟開關及計時器
			_stop = false;
			timer = setTimeout(move, _showSpeed);
		}).find('a').focus(function(){
			this.blur();
		});
		
		// 計時器使用
		function move(){
			$next.click();
		};
		timer = setTimeout(move, _showSpeed);
	});