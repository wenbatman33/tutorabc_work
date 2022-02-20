var IMG_WIDTH = 887;
var currentImg=0;
var speed=300;
var maxImages;
var imgs;
var imgsWrap;
var swipeOptions={
	triggerOnTouchEnd : true,	
	swipeStatus : swipeStatus,
	allowPageScroll:"vertical",
	threshold:200		
}

$(function(){	
	
	maxImages= $(".showcase ul li").length;


		
	imgs = $(".showcase");
	imgsWrap = $(".showcaseWrap");
	imgs.swipe( swipeOptions );
	$(".rightThumbnailWrap ul li").click(function() {
	

	currentImg=$(this).index();
	 rollImage();
	});
	
});
function swipeStatus(event, phase, direction, distance){
	if( phase=="end" && direction==null){
		zoomIn();
	
	}	else if( phase=="move" && (direction=="left" || direction=="right") ){
				
		var duration=0;
		if (direction == "left"){
			
			scrollImages((IMG_WIDTH * currentImg) + distance, duration);
		}else if(direction == "right"){
	
			scrollImages((IMG_WIDTH * currentImg) - distance, duration);
		}
	}
	else if ( phase == "cancel"){
		scrollImages(IMG_WIDTH * currentImg, speed);
	}else if ( phase =="end" ){
		if (direction == "right"){
			previousImage()
		}else if (direction == "left")		{	
			nextImage()
		}
	}
}

function scrollImages(distance, duration)
{
	imgsWrap.css("-webkit-transition-duration", (duration/1000).toFixed(1) + "s");
	var value = (distance<0 ? "" : "-") + Math.abs(distance).toString();
	imgsWrap.css("-webkit-transform", "translate3d("+value +"px,0px,0px)");

}
function previousImage(){
	console.log(currentImg);
	currentImg = Math.max(currentImg-1, 0);
	scrollImages( IMG_WIDTH * currentImg, speed);
}

function nextImage(){
	
	console.log(currentImg);
	currentImg = Math.min(currentImg+1, maxImages-1);
	scrollImages( IMG_WIDTH * currentImg, speed);
}
function rollImage(){
	console.log(currentImg);
	currentImg = Math.min(currentImg, maxImages);
	scrollImages( IMG_WIDTH * currentImg, speed);
}

function zoomIn(){
		$(".rightThumbnail").animate({left:-500}, 0 );
		$(".showcase").animate({width:1024 , height:768 ,top: 0, left:0}, 0 );
		
		$(".showcase ul li").css({"width":1024 , "height":768, "float" :"left" });
		$(".showcase ul").css({"display":"block" ,width:99999,"list-style":"none"});
		
	
	    IMG_WIDTH = 1034;
	
		$(".bottomStatus").css("position","absolute");
		$(".bottomStatus").animate({bottom:99999}, 0 );


		$(".topNav").css("position","absolute");
		$(".topNav").animate({top:-9999}, 0 );
		
		$(".sidebar").animate({left:-999}, 0 );
		
		sidebarOpen="false";
		$(".SalesViewPort").animate({left:0}, 0 );
		$('.SalesContainer').transition({ perspective: '1024px', rotateY: '0deg'}, 0);
		$(".zoomOutBtn").css("display","block");
		
		rollImage();

		$(".zoomOutBtn").click(function() {
			zoomOut();
		});
}
function zoomOut(){
	    $(".rightThumbnail").animate({left:0}, 0 );
		$(".showcase").animate({width:877 , height:667 ,top: 52, left:148},0);
		
		$(".showcase ul li").css({"width":877 , "height":667, "float" :"left" });
		$(".showcase ul").css({"display":"block" ,width:99999,"list-style":"none"});
		
	    IMG_WIDTH = 887;
	
		$(".bottomStatus").css("position","absolute");
		$(".bottomStatus").animate({bottom:0}, 0 );

		$(".topNav").css("position","absolute");
		$(".topNav").animate({top:0}, 0 );
		
		$(".sidebar").animate({left:-246}, 0 );
		
		sidebarOpen="false";
		$(".SalesViewPort").animate({left:0}, 0 );
		$('.SalesContainer').transition({ perspective: '1024px', rotateY: '0deg'}, 0);
		$(".zoomOutBtn").css("display","none");

		rollImage();
}