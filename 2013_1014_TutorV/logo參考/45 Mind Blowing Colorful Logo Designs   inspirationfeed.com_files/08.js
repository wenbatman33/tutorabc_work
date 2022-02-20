if(typeof(Nster) == "undefined") 
    var Nster = (function(){
	var get = function (id){ return document.all ? document.all[id] : document.getElementById(id); },
	draw = function (obj){
	    if(obj.length) {
		obj = this.prepare(obj);
		var n = this.name, 
		    e = get(this.name),
		    ih = this.tpl.call(obj);
		if(e) e.innerHTML = ih;
		else {
                var d=document.createElement('div');
                d.innerHTML = ih; d.id = n; var a = get('a'+n);
                if(a) a.parentNode.insertBefore(d, a);                    
                }
		this.filled = 1;
                if (this.cfg.d) timer(this.name, this.cfg.d);
	    }
	},
	cread = function(n){
	    var b=/^\s*(\S+)\s*=\s*(\S*)\s*$/, c=document.cookie.split(';'),i=c.length;
	    while(i--) if(b.exec(c[i]) && RegExp.$1==n) return Number(RegExp.$2);
	    return 0;
	},
	cwrite = function(n,v){
	    document.cookie = n + '=' + v + ';path=/';
	},
	page = function() {
	    var name = this.name, page = cread(name)+1;
	    if(page>5) page = 1;
	    cwrite(name,page);
	    return page;
	},
	prepare = function(obj) {
	    var cfg = this.cfg, l = cfg.nn;
	    while(l--) {
		var i = obj[l],o={img:'data:image/gif;base64,R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==',descr:'',title:'',href:''};
		if(!i) {obj[l] = o; continue;}
		if (i.cl) {t = document.createElement('a');t.href = i.cl;o.cl = t.hostname;} 
                else {o.cl = '';} 
		var image = i.i, 
		    t=i.t,tl=t.length,ctl=cfg.tl,
		    td=i.td,tdl=td.length,ctdl = cfg.tdl;
		o.href = cfg.link+'/'+cfg.id+'/'+i.n + this.ref;
		o.title = tl>ctl? t.slice(0,ctl)+(ctl?'...':''):t;
		o.descr = tdl>ctdl? td.slice(0,ctdl)+(ctdl?'...':''):td;
		if(image && cfg.i) {
		    image = image.toString();
		    image = (function(n){return n>0?new Array(n+1).join('0'):''})(6-image.length)+image;
		    var il = image.length, s1 = image.slice(0, il - 4), s2 = image.slice(il - 4, il - 2);
		    o.img = cfg.cdn+"/"+s1+"/"+s2+"/"+image+"_"+cfg.is+".jpg";
		    o.alt = t.replace(/\"/g,'&quot;');
		}
		obj[l] = o;
	    }
	    return obj;
	},
	hide = function(n){
	    var a = get('a'+n);
	    if(a) a.style.display = 'none';
	},
	call = function () {
	    hide(this.name);
	    var cfg = this.cfg;
var js = document.createElement('SCRIPT');
js.src = cfg.js+'/a/'+cfg.id+'/' + this.page() + '.js';
js.type = "text/javascript";
js.charset = 'UTF-8';
js.async = true;
var hh = document.getElementsByTagName('head')[0];
hh.appendChild(js);             
	},
        timer = function (n, d) {
    
            function elementInViewport(el) {
              var top = el.offsetTop;
              var left = el.offsetLeft;
              var width = el.offsetWidth;
              var height = el.offsetHeight;

              while(el.offsetParent) {
                el = el.offsetParent;
                top += el.offsetTop;
                left += el.offsetLeft;
              }

              return (
                top < (window.pageYOffset + window.innerHeight) &&
                left < (window.pageXOffset + window.innerWidth) &&
                (top + height) > window.pageYOffset &&
                (left + width) > window.pageXOffset
              );
            }  
        function Timer(callback, delay) {
            var timerId, start, remaining = delay;
            this.pause = function() {
                if (timerId) {
                    window.clearTimeout(timerId);
                    remaining -= new Date() - start;
                    timerId = null;
                }
            };

            this.resume = function() {
                if (!timerId) {
                    start = new Date();
                    timerId = window.setTimeout(callback, remaining);
                }
            };
        }      
           var timer = new Timer(function () {window[n].call()}, d);
            function timerControl()
            {
                if (elementInViewport(get(n))) {
                    timer.resume();
                } else {
                    timer.pause();
                }
            }
            window.onresize = timerControl;
            window.onscroll = timerControl;
            timerControl();
        },                
	construct = function (cfg,tpl,ref) {
	    if ( !(this instanceof arguments.callee) ) return new arguments.callee(cfg,tpl,ref);
	    this.name = 'Nster'+cfg.id;
	    this.ref = ref ? '?ref='+encodeURIComponent(ref) : '';
	    window[this.name] = this;
	    this.filled = 0;
	    this.cfg = cfg;
	    this.tpl = tpl;
	    this.draw = draw;
	    this.page = page;
	    this.prepare = prepare;
	    this.call = call;
	    this.call();
	};
	return construct;
    })();
(function(){
	var ref = typeof(refNster) == "undefined" ? '' : refNster
	,cfg = {
		id:208,pid:314,
		nn:4,
		i: 1, is: '150x150',
                d: 0,
		tl: 45, tdl: 0, 
		cdn: 'http://i.nster.net', link: 'http://nster.com/go',js: 'http://d.nster.net'
		}, 
	tpl = function(){return [
		'<style>\
.WgtNews208 {\
	margin-top: 4px;\
	padding-left: 0px;\
}\
.WgtTitle208 {\
     margin-right:1px;\
     margin-left:1px;\
\
}\
.WgtImg208 {\
valign:top;\
margin-top: 4px;\
width:140px;\
height:140px;\
}\
.WgtBgColor2208 {\
  padding: 0px !important;\
}\
.WgtBold208{\
font-weight:bold;\
}\
a.WgtLink208 {\
	color: #0000FF;\
}\
a.WgtLink208:hover {\
	color:#555555;\
}\
a.WgtLink208 {\
	text-decoration:none;\
}\
a.WgtLink208:hover {\
	text-decoration:none;\
}\
.WgtBorder208 {\
	border-color:#FFFFFF;\
}\
.WgtBorder208 {\
	border-style:solid;\
}\
.WgtBorder208 {\
	border-width:0px;\
}\
.WgtBody208 {\
	font-family:Arial, Helvetica, sans-serif;\
}\
.WgtBody208 {\
	font-size:12px;\
}\
.WgtBody208 {\
	font-weight:normal;\
}\
.WgtImg208 {\
border: 1px solid  #CCCCCC !important;\
padding:2px;\
}\
.news33:hover img\
{\
border: 3px solid  #555555 !important;\
padding:0;\
-webkit-box-shadow: 0px 0px 12px rgba(50, 50, 50, 1);\
-moz-box-shadow:    0px 0px 12px rgba(50, 50, 50, 1);\
box-shadow:         0px 0px 12px rgba(50, 50, 50, 1);\
}\
</style>\
<div align="left">\
<div class="WgtBody208 WgtBorder208" style="width:100%;overflow:hidden;height:100%">\
<table cellpadding="0" cellspacing="3" class="news" width="100%" height="100%">\
<tr>\
\
        <td valign="top" align="center" class="WgtBgColor2_208" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="140"><tr>\
			<td  valign="top" align="left"  width="140" height="140">\
				<div class="WgtImgDiv208">\
					<div align="left"><a valign="top" href="',this[0].href,'" target="_blank" class="WgtLink208">\
						<img class="WgtImg208" src="',this[0].img,'" valign="top" border="0" width="150">\
					</a></div>\
				</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="left" style="padding-left: 0px;">\
        	<div>\
        	        <div align="left"><a class="WgtLink208 WgtLink2_208 WgtBody208" href="',this[0].href,'" target="_blank"><span class="WgtTitle208">',this[0].title,'</span></a></div>\
					<span class="WgtDesc208">',this[0].descr,'</span>\
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
        <td valign="top" align="center" class="WgtBgColor2_208" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="140"><tr>\
			<td  valign="top" align="left"  width="140" height="140">\
				<div class="WgtImgDiv208">\
					<div align="left"><a valign="top" href="',this[1].href,'" target="_blank" class="WgtLink208">\
						<img class="WgtImg208" src="',this[1].img,'" valign="top" border="0" width="150">\
					</a></div>\
				</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="left" style="padding-left: 0px;">\
        	<div>\
        	        <div align="left"><a class="WgtLink208 WgtLink2_208 WgtBody208" href="',this[1].href,'" target="_blank"><span class="WgtTitle208">',this[1].title,'</span></a></div>\
					<span class="WgtDesc208">',this[1].descr,'</span>\
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
        <td valign="top" align="center" class="WgtBgColor2_208" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="140"><tr>\
			<td  valign="top" align="left"  width="140" height="140">\
				<div class="WgtImgDiv208">\
					<div align="left"><a valign="top" href="',this[2].href,'" target="_blank" class="WgtLink208">\
						<img class="WgtImg208" src="',this[2].img,'" valign="top" border="0" width="150">\
					</a></div>\
				</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="left" style="padding-left: 0px;">\
        	<div>\
        	        <div align="left"><a class="WgtLink208 WgtLink2_208 WgtBody208" href="',this[2].href,'" target="_blank"><span class="WgtTitle208">',this[2].title,'</span></a></div>\
					<span class="WgtDesc208">',this[2].descr,'</span>\
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
        <td valign="top" align="center" class="WgtBgColor2_208" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="140"><tr>\
			<td  valign="top" align="left"  width="140" height="140">\
				<div class="WgtImgDiv208">\
					<div align="left"><a valign="top" href="',this[3].href,'" target="_blank" class="WgtLink208">\
						<img class="WgtImg208" src="',this[3].img,'" valign="top" border="0" width="150">\
					</a></div>\
				</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="left" style="padding-left: 0px;">\
        	<div>\
        	        <div align="left"><a class="WgtLink208 WgtLink2_208 WgtBody208" href="',this[3].href,'" target="_blank"><span class="WgtTitle208">',this[3].title,'</span></a></div>\
					<span class="WgtDesc208">',this[3].descr,'</span>\
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
</tr>\
</table>\
</div>\
</div>'
	].join('')};
	Nster(cfg,tpl,ref);
})();
		