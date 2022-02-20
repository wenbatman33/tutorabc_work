//leoli ����function�]�_�Ө�ajax���s���w�C
$(document).ready(function () {
    $.TipCustom.BindTip();
});

(function ($) {
    var current = $.TipCustom = {};
    $.extend(current, {
        BindTip: function () {
            if(navigator.userAgent.match("MSIE")){
		
		        if(navigator.appName == "Microsoft Internet Explorer" && navigator.appVersion.match(/7./i)=="7."){
			        var topNum =0;
			        var bottomNum =-12;
		        }else{
			        var topNum =26;
			        var bottomNum =28;
		        }
		
	        }else{
		        var topNum =12;
		        var bottomNum =22;
	        }

            var targets = $('[rel~="tooltip"]'),
            target = false,
            tooltip = false,
            alt = false;

            targets.bind('mouseenter', function () {
                var target = $(this);
                var tip = target.attr('alt');
                var tooltip = $('<div id="tooltip"></div>');

                if (!tip || tip == '')
                    return false;

                target.removeAttr('alt');
                tooltip.css('opacity', 0)
               .html(tip)
               .appendTo('body');

                var init_tooltip = function () {
                    if ($(window).width() < tooltip.outerWidth() * 1.5)
                        tooltip.css('max-width', $(window).width() / 2);
                    else
                        tooltip.css('max-width', 340);

                    var pos_left = target.offset().left + (target.outerWidth() / 2) - (tooltip.outerWidth() / 2);
                    var pos_top = target.offset().top - tooltip.outerHeight() - topNum;

                    if (pos_left < 0) {
                        pos_left = target.offset().left + target.outerWidth() / 2 - topNum;
                        tooltip.addClass('left');
                    }
                    else
                        tooltip.removeClass('left');

                    if (pos_left + tooltip.outerWidth() > $(window).width()) {
                        pos_left = target.offset().left - tooltip.outerWidth() + target.outerWidth() / 2 + 20;
                        tooltip.addClass('right');
                    }
                    else
                        tooltip.removeClass('right');

                    if (pos_top < 0) {
                        var pos_top = target.offset().top + target.outerHeight()+bottomNum;
                        tooltip.addClass('top');
                    }
                    else
                        tooltip.removeClass('top');

                    tooltip.css({ left: pos_left, top: pos_top })
                   .animate({  opacity: 1 }, 50);
                };

                init_tooltip();
                $(window).resize(init_tooltip);

                var remove_tooltip = function (e) {
                    tooltip.animate({ opacity: 0 }, 50, function () {
                        $(this).remove();
                        target.attr('alt', tip);
                        
                        target.unbind('mouseleave');
                        tooltip.unbind('click');
                    });
                };

                target.bind('mouseleave', remove_tooltip);
                tooltip.bind('click', remove_tooltip);
            });
        }
    });
} (window.jQuery));