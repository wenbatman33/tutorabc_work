var __snShow = 'http://show.chicago.sn00.net';
var __snImg = 'http://img.chicago.sn00.net';
pageSN = Number(readCookie('snewsPage'));
if (pageSN >= 5 ) pageSN = 0;
setCookie('snewsPage',pageSN+1);

if (typeof(snInfArr) == 'undefined') snInfArr = [];
snInfNum = snInfArr.length;
snInfArr[snInfNum] = [];

snInfArr[snInfNum]['iid'] = 1373;
snInfArr[snInfNum]['filled'] = 0;
snInfArr[snInfNum]['pid'] = 857;
snInfArr[snInfNum]['mn'] = 4;
snInfArr[snInfNum]['mts'] = 80;
snInfArr[snInfNum]['mds'] = 0;
snInfArr[snInfNum]['show_descr'] = 0;
/*snInfArr[snInfNum]['img_size'] = '0';*/
snInfArr[snInfNum]['img_suffix'] = '140x140';
snInfArr[snInfNum]['img_height'] = 140;
snInfArr[snInfNum]['img_width'] = 140;
snInfArr[snInfNum]['directlink'] = 0;
snInfArr[snInfNum]['shh'] = 'amplify.feedbox.com';
snInfArr[snInfNum]['show_image'] = 1;

snInfArr[snInfNum]['ads_currentNews'] = pageSN;
snInfArr[snInfNum]['ads_maxNews'] = 3;
snInfArr[snInfNum]['ads_news'] = [];
document.write('\
<style>\
.snNews1373 {\
	margin-top: 0px;\
	padding-left: 0px;\
}\
.snTitle1373 {\
     margin-right:1px;\
     margin-left:1px;\
}\
.snImg1373 {\
valign:top;\
}\
.snBgColor21373 {\
  padding: 0px !important;\
}\
.snBold1373{\
font-weight:bold;\
}\
a.snLink1373 {\
	color:#000000;\
}\
a.snLink1373:hover {\
	color:#000000;\
}\
a.snLink1373 {\
	text-decoration:none;\
}\
a.snLink1373:hover {\
	text-decoration:none;\
}\
.snBorder1373 {\
	border-color:#FFFFFF;\
}\
.snBorder1373 {\
	border-style:solid;\
}\
.snBorder1373 {\
	border-width:0px;\
}\
.snBody1373 {\
	font-family:Arial, Helvetica, sans-serif;\
}\
.snBody1373 {\
	font-size:12px;\
}\
.snBody1373 {\
	font-weight:normal;\
}\
\
</style>\
<div align="center">\
<div id="snInformer1373" class="snBody1373 snBorder1373" style="display:none;width:100%;overflow:hidden;height:100%">\
\
<table cellpadding="0" cellspacing="3" class="news" width="100%" height="100%">\
<tr>\
\
        <td valign="top" align="center" class="snBgColor21373" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="100%"><tr>\
			<td  valign="top" align="center"  width="1">\
<div class="snImgDiv1373">\
				<a valign="top" href="" target="_blank" id="sn_link_1373_0" class="snLink1373"><img class="snImg1373" src="http://img.chicago.sn00.net/1x1.gif" id="sn_img_1373_0" valign="top" border="0" width="140"></a>\
</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="center" style="padding-left: 0px;">\
        	<div>\
        	        <a class="snLink1373 snLink2_1373 snBody1373" href="" target="_blank" id="sn_link2_1373_0">\
                 	<span class="snTitle1373" id="sn_title_1373_0"> </span>\
        	        </a><br/><span class="snDesc1373" id="sn_desc_1373_0"> </span>\
\
         \
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
        <td valign="top" align="center" class="snBgColor21373" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="100%"><tr>\
			<td  valign="top" align="center"  width="1">\
<div class="snImgDiv1373">\
				<a valign="top" href="" target="_blank" id="sn_link_1373_1" class="snLink1373"><img class="snImg1373" src="http://img.chicago.sn00.net/1x1.gif" id="sn_img_1373_1" valign="top" border="0" width="140"></a>\
</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="center" style="padding-left: 0px;">\
        	<div>\
        	        <a class="snLink1373 snLink2_1373 snBody1373" href="" target="_blank" id="sn_link2_1373_1">\
                 	<span class="snTitle1373" id="sn_title_1373_1"> </span>\
        	        </a><br/><span class="snDesc1373" id="sn_desc_1373_1"> </span>\
\
         \
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
        <td valign="top" align="center" class="snBgColor21373" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="100%"><tr>\
			<td  valign="top" align="center"  width="1">\
<div class="snImgDiv1373">\
				<a valign="top" href="" target="_blank" id="sn_link_1373_2" class="snLink1373"><img class="snImg1373" src="http://img.chicago.sn00.net/1x1.gif" id="sn_img_1373_2" valign="top" border="0" width="140"></a>\
</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="center" style="padding-left: 0px;">\
        	<div>\
        	        <a class="snLink1373 snLink2_1373 snBody1373" href="" target="_blank" id="sn_link2_1373_2">\
                 	<span class="snTitle1373" id="sn_title_1373_2"> </span>\
        	        </a><br/><span class="snDesc1373" id="sn_desc_1373_2"> </span>\
\
         \
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
        <td valign="top" align="center" class="snBgColor21373" width="25%">\
          <table class="news33" valign="top" cellpadding="0" cellspacing="0" width="100%"><tr>\
			<td  valign="top" align="center"  width="1">\
<div class="snImgDiv1373">\
				<a valign="top" href="" target="_blank" id="sn_link_1373_3" class="snLink1373"><img class="snImg1373" src="http://img.chicago.sn00.net/1x1.gif" id="sn_img_1373_3" valign="top" border="0" width="140"></a>\
</div>\
			</td>\
		 </tr>\
		 <tr>\
			<td valign="top" align="center" style="padding-left: 0px;">\
        	<div>\
        	        <a class="snLink1373 snLink2_1373 snBody1373" href="" target="_blank" id="sn_link2_1373_3">\
                 	<span class="snTitle1373" id="sn_title_1373_3"> </span>\
        	        </a><br/><span class="snDesc1373" id="sn_desc_1373_3"> </span>\
\
         \
			</div>\
            </td>\
		</tr></table>\
        </td>\
\
\
</tr>\
</table>\
</div>\
</div>\
<div id="ubl1373" class="disp-no"></div>\
');
if(typeof snLocalNews == 'undefined') document.write('<sc'+'ript type="text/javascript" src="http://show.chicago.sn00.net/a/1373/' + (pageSN + 1) + '.js"></scr'+'ipt>');

/**
 * Create a new Document object. If no arguments are specified,
 * the document will be empty. If a root tag is specified, the document
 * will contain that single root tag. If the root tag has a namespace
 * prefix, the second argument must specify the URL that identifies the
 *namespace.
 */
function XMLnewDocument(rootTagName, namespaceURL) {
    if (!rootTagName) rootTagName = "";
    if (!namespaceURL) namespaceURL = "";

    if (document.implementation && document.implementation.createDocument) {
        // This is the W3C standard way to do it
        return document.implementation.createDocument(namespaceURL, 
                       rootTagName, null);
    }
    else { // This is the IE way to do it
        // Create an empty document as an ActiveX object
        // If there is no root element, this is all we have to do
        var doc = new ActiveXObject("MSXML2.DOMDocument");

        // If there is a root tag, initialize the document
        if (rootTagName) {
            // Look for a namespace prefix
            var prefix = "";
            var tagname = rootTagName;
            var p = rootTagName.indexOf(':');
            if (p != -1) {
                prefix = rootTagName.substring(0, p);
                tagname = rootTagName.substring(p+1);
            }

            // If we have a namespace, we must have a namespace prefix
            // If we don't have a namespace, we discard any prefix
            if (namespaceURL) {
                if (!prefix) prefix = "a0"; // What Firefox uses
            }
            else prefix = "";

            // Create the root element (with optional namespace) as a
            // string of text
            var text = "<" + (prefix?(prefix+":"):"") + tagname +
                (namespaceURL
                 ?(" xmlns:" + prefix + '="' + namespaceURL +'"')
                 :"") +
                "/>";
            // And parse that text into the empty document
            doc.loadXML(text);
        }
        return doc;
    }
};

/**
 * Synchronously load the XML document at the specified URL and
 * return it as a Document object
 */
function XMLload(url) {
    // Create a new document with the previously defined function
    var xmldoc = XMLnewDocument();
    xmldoc.async = false;  // We want to load synchronously
    xmldoc.load(url);      // Load and parse
    return xmldoc;         // Return the document
}; 


/**
 * Asynchronously load and parse an XML document from the specified URL.
 * When the document is ready, pass it to the specified callback function.
 * This function returns immediately with no return value.
 */
function XMLloadAsync(url, callback) {
    var xmldoc = XMLnewDocument();

    // If we created the XML document using createDocument, use
    // onload to determine when it is loaded
    if (document.implementation && document.implementation.createDocument) {
        xmldoc.onload = function() { callback(xmldoc); };
    }
    // Otherwise, use onreadystatechange as with XMLHttpRequest
    else {
        xmldoc.onreadystatechange = function() {
            if (xmldoc.readyState == 4) callback(xmldoc);
        };
    }

    // Now go start the download and parsing
    xmldoc.load(url);
}; 

/**
 * Parse the XML document contained in the string argument and return
 * a Document object that represents it.
 */
/*XML.parse = function(text) {*/
function XMLparse(text) {
    if (typeof DOMParser != "undefined") {
        // Mozilla, Firefox, and related browsers
        return (new DOMParser()).parseFromString(text, "application/xml");
    }
    else if (typeof ActiveXObject != "undefined") {
        // Internet Explorer.
        var doc = XMLnewDocument( );   // Create an empty document
        doc.loadXML(text);              //  Parse text into it
        return doc;                     // Return it
    }
    else {
        // As a last resort, try loading the document from a data: URL
        // This is supposed to work in Safari. Thanks to Manos Batsis and
        // his Sarissa library (sarissa.sourceforge.net) for this technique.
        var url = "data:text/xml;charset=utf-8," + encodeURIComponent(text);
        var request = new XMLHttpRequest();
        request.open("GET", url, false);
        request.send(null);
        return request.responseXML;
    }
};



function readCookie(name) {
	var b=/^\s*(\S+)\s*=\s*(\S*)\s*$/,
	c=document.cookie.split(';');
	for(var i=0;i<c.length;i++)
	  if(b.exec(c[i]) && RegExp.$1==name) return unescape(RegExp.$2);
	return null;
}

function setCookie(szName, szValue, szExpires, szPath, szDomain, bSecure)
{
        var szCookieText =         escape(szName) + '=' + escape(szValue);
        szCookieText +=            (szExpires ? '; EXPIRES=' + szExpires.toGMTString() : '');
        szCookieText +=            (szPath ? '; PATH=' + escape(szPath) : '');
        szCookieText +=            (szDomain ? '; DOMAIN=' + szDomain : '');
        szCookieText +=            (bSecure ? '; SECURE' : '');

        document.cookie = szCookieText;
}

function getXmlHttp(){
  var xmlhttp;
  try {
    xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
  } catch (e) {
    try {
      xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    } catch (E) {
      xmlhttp = false;
    }
  }
  if (!xmlhttp && typeof XMLHttpRequest!='undefined') {
    xmlhttp = new XMLHttpRequest();
  }
  return xmlhttp;
}

function loadNews(snInfNum){
    if(typeof snLocalNews == 'undefined') snLocalNews = false;
    if(snLocalNews) return;
    var js = document.createElement('SCRIPT');
    js.src = __snShow+'/a/'+iid+'/' + (snInfArr[snInfNum]['ads_currentNews'] + 1) + '.js';
    js.type = "text/javascript";
    var hh = document.getElementsByTagName('head')[0];
    hh.appendChild(js);
    //__id('ubl'+iid).appendChild(js)
}
function __id(sI){
    return document.all ? document.all[sI] : document.getElementById(sI);
}
__Images =
{
  count: [] /* keep track of the number of images */
  ,loaded: [] /* keeps track of how many images have loaded */
  ,onComplete: function(iid){
    var oInformer = __id('snInformer'+iid);
    if(oInformer) oInformer.style.display = 'block';
  } /* fires when all images have finished loadng */
  ,onLoaded: function(){} /*fires when an image finishes loading*/
  ,loaded_image: "" /*access what has just been loaded*/
  ,images: [] /*keeps an array of images that are loaded*/
  ,incoming:[] /*this is for the process queue.*/
  /* this will pass the list of images to the loader*/
  ,reset: function(iid)
  {
    //make sure to reset the counters
    this.loaded[iid] = 0;
    this.count[iid] = 0;

    //reset the images array 
    this.images[iid] = [];
    //reset queue array also
    this.incoming[iid] = [];
  }
  ,add: function(image,iid)
  {
    if(!image) return;
    //increment the number of images
    this.count[iid]++;

    //store the image name
    this.incoming[iid].push(image);

    //start processing the images one by one
    this.process(iid);

  }
  ,process: function(iid)
  {
    //pull the next image off the top and load it
    this.load(this.incoming[iid].shift(),iid);
  }
  /* this will load the images through the browser */
  ,load: function(image,iid)
  {
    if(!image) return
    var this_ref = this;
    var i = new Image;

    i.onload = function()
    {
      //store the loaded image so we can access the new info
      this_ref.loaded_image = i;

      //push images onto the stack
      this_ref.images[iid].push(i);

      //note that the image loaded
      this_ref.loaded[iid] +=1;

      //fire the onloaded
      (this_ref.onLoaded)();

      //if all images have been loaded launch the call back
      if(this_ref.count[iid] == this_ref.loaded[iid])
      {
        (this_ref.onComplete)(iid);
      }
      //load the next image
      else
      {
        this_ref.process(iid);
      }
    }
    i.src = image;
  }
  ,done:function(iid){
    return (this.count[iid] == this.loaded[iid])
  }
}
function fillNews(iid, xmlDoc){
	var oldframe = __id('snf'+iid);
	if(oldframe) oldframe.style.display = 'none';
	var xmlDoc = XMLparse(xmlDoc);
	var snInfNum = 0;
	for (snInfNum in snInfArr) {
		if (snInfArr[snInfNum]['iid'] == iid) break;
	}
	snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']] = [];
        __Images.reset(iid);
	for(var i = 0 ; i < snInfArr[snInfNum]['mn'] ; i++){
		var descr = '';
		if (xmlDoc.getElementsByTagName('guid')[i] == undefined) continue;
		if (xmlDoc.getElementsByTagName('description').length != 0){
			if (xmlDoc.getElementsByTagName('description')[i].firstChild != null) {
				descr = xmlDoc.getElementsByTagName('description')[i].firstChild.nodeValue;
			}
		}
                var image = __snImg+"/1x1.gif";
                if(snInfArr[snInfNum]['show_image']) {
                    image = xmlDoc.getElementsByTagName('image')[i].firstChild.nodeValue;
		    var prefix = image%2 ? 1:2;
                    if(image.length < 6){
                            while (image.length < 6) {
                                    image = "0"+image;
                            }
                    }
                    var d1 = image.slice(0, image.length - 4),
                    d2 = image.slice(image.length - 4, image.length - 2),
                    img_suffix = "sized",
                    img_size = snInfArr[snInfNum]['img_size'];
                    if (img_size !== undefined){
                            if (img_size != '80')
                            {
                                    img_suffix = img_size;
                            }
                    }else if(snInfArr[snInfNum]['img_suffix']){
                            img_suffix = snInfArr[snInfNum]['img_suffix'];
                    }
                    //image = 'http://img'+prefix+'.ns.sn00.net'+"/"+d1+"/"+d2+"/"+image+"_"+img_suffix+".jpg";
                    image = __snImg+"/"+d1+"/"+d2+"/"+image+"_"+img_suffix+".jpg";
                }
                __Images.add(image,iid);
		snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i]= [
			xmlDoc.getElementsByTagName('guid')[i].firstChild.nodeValue,
			xmlDoc.getElementsByTagName('title')[i].firstChild.nodeValue,
			xmlDoc.getElementsByTagName('link')[i].firstChild.nodeValue,
			descr,
			image,
			xmlDoc.getElementsByTagName('partnerId')[i].firstChild.nodeValue,
			xmlDoc.getElementsByTagName('directLink')[i].firstChild.nodeValue,
			xmlDoc.getElementsByTagName('partnerName')[i].firstChild.nodeValue
                        ];
	}
	ads_loadNews(iid);
}

function ads_loadNews(iid){
	var snInfNum = 0;
	for (snInfNum in snInfArr) {
		if (snInfArr[snInfNum]['iid'] == iid) break;
	}
        var oInformer = __id('snInformer'+iid);
        if(oInformer) oInformer.style.display = 'none';
	if (snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']]==undefined){
		loadNews(snInfNum);
		return;
	} else if(!snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']].length) {
                return;
        }
	for (var i = 0; i < snInfArr[snInfNum]['mn']; i++) {
		var sn_link = __id("sn_link_"+iid+"_"+i),
		    sn_link2 = __id("sn_link2_"+iid+"_"+i),
		    sn_link3 = __id("sn_link3_"+iid+"_"+i),
		    sn_partner = __id("sn_partner_"+iid+"_"+i),
		    sn_descr = __id("sn_desc_"+iid+"_"+i),
		    sn_img = __id("sn_img_"+iid+"_"+i),
		    sn_title = __id("sn_title_"+iid+"_"+i);

		if(snInfArr[snInfNum]['filled']) {
		  // Очистка
		  if(sn_title) {
		    sn_title.innerHTML = "";
		    sn_title.style.display = "none";
		  }
		  if (sn_partner){
			  sn_partner.innerHTML = "";
			  sn_partner.style.display = "none";
		  }
		  sn_link.href = "";
		  if(sn_link2) sn_link2.href = "";
		  if(sn_link3) sn_link3.href = "";
		  if (snInfArr[snInfNum]['show_descr'] && sn_descr) {
			  sn_descr.innerHTML = "";
		  }
		  if (sn_img) {
			  sn_img.style.display = 'none';
			  sn_img.src = __snImg+"/1x1.gif";
			  sn_img.alt = "";
			  sn_img.title = "";
		  }
		}
		// Заполнение
		if (snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i]==undefined) continue;
		var snTitle = snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][1];
		if(snTitle.length > snInfArr[snInfNum]['mts']){
			snTitle = snTitle.slice(0,snInfArr[snInfNum]['mts']-3)+"...";
		}
		var descr = snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][3];
		if(descr.length > snInfArr[snInfNum]['mds']){
			descr = descr.slice(0,snInfArr[snInfNum]['mds']-3)+"...";
		}
		var snLink = 'http://'+snInfArr[snInfNum]['shh']+'/link/?'+
			'pid='+snInfArr[snInfNum]['pid']+'&iid='+iid+'&d='+snInfArr[snInfNum]['directlink']+
			'&nid='+snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][0]+
			'&npid='+snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][5]+
			'&pd='+snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][6]+
			'&url='+escape(snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][2])+
			'&r='+escape(document.location.href);
		snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i]['reallink'] = snLink;
		//snLink = snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][2];
		snLink = snLink + '&f=1';
		if(sn_title) {
		  sn_title.innerHTML = snTitle;
		  sn_title.style.display = "inline";
		}
		if (sn_partner){
			var name = snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][7];
			sn_partner.innerHTML = name;
			sn_partner.style.display = "inline";
		}
		
		(function(sn_link,iid,i){
			sn_link.onmouseover = function(){
				snGetLink(sn_link.id,iid,i)
			}
		})(sn_link,snInfNum,i);	
		sn_link.href = snLink;
		if(sn_link2) {
		    sn_link2.href = snLink;
		    (function(sn_link,iid,i){
			sn_link.onmouseover = function(){
				snGetLink(sn_link.id,iid,i)
			}
		})(sn_link2,snInfNum,i);	
		}
		if(sn_link3) {
		    sn_link3.href = snLink;
		    (function(sn_link,iid,i){
			sn_link.onmouseover = function(){
				snGetLink(sn_link.id,iid,i)
			}
		})(sn_link3,snInfNum,i);	
		}
		if (snInfArr[snInfNum]['show_descr']) {
			if (sn_descr)
				sn_descr.innerHTML = descr;
		}
		if (snInfArr[snInfNum]['show_image']) {
			if (sn_img) {
				sn_img.src = snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i][4];
				sn_img.style.display = 'inline';
				sn_img.alt = snTitle;
				sn_img.title = snTitle;
			}
		}
		snInfArr[snInfNum]['filled']=1;
	}
        if(__Images.done(iid)) __Images.onComplete(iid);
}

function snGetLink(id, snInfNum, i){
	var e = __id(id);
	if (e) e.href = snInfArr[snInfNum]['ads_news'][snInfArr[snInfNum]['ads_currentNews']][i]['reallink'];
	return true;
}

function ads_arrOut(id){
	__id(id).style.backgroundColor="#eeeeff";
}

function ads_hasPrev(){
	var st = __id("smap").style;
	if (ads_currentNews > 0){
		st.backgroundColor="#9999ff";
		st.cursor="pointer";
	}else{
		st.backgroundColor="#eeeeff";
		st.cursor="default";
	}
}

function ads_hasNext(){
	var st = __id("sman").style;
	if (ads_currentNews+1 < ads_maxNews){
		st.backgroundColor="#9999ff";
		st.cursor="pointer";
	}else{
		st.backgroundColor="#eeeeff";
		st.cursor="default";
	}
}

function ads_prev(){
	if (ads_currentNews > 0){
		ads_currentNews = ads_currentNews-1;
		ads_loadNews();
		ads_hasPrev();
	}
}

function ads_next(){
	if (ads_currentNews+1 < ads_maxNews){
		ads_currentNews = ads_currentNews+1;
		ads_loadNews();
		ads_hasNext();
	}
}

/**
 * Get argument value from URL by name
 */ 
function gup( name ) 
{
  name = name.replace(/[\[]/,"\\\[").replace(/[\]]/,"\\\]");
  var regexS = "[\\?&]"+name+"=([^&#]*)";
  var regex = new RegExp( regexS );
  var results = regex.exec( window.location.href );
  if( results == null )
    return "";
  else
    return results[1];
}



    