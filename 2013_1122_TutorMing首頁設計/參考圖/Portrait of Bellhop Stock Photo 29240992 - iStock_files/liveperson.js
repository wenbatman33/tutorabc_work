
var lpMTagConfig={"lpServer":"sales.liveperson.net","lpNumber":"72069871","lpProtocol":(document.location.toString().indexOf("https:")==0)?"https":"http","lpTagLoaded":false,"pageStartTime":(new Date()).getTime()}
function lpAddMonitorTag(src){if(!lpMTagConfig.lpTagLoaded){if(typeof(src)=="undefined"||typeof(src)=="object"){src=lpMTagConfig.lpMTagSrc?lpMTagConfig.lpMTagSrc:"/hcp/html/mTag.js";}
if(src.indexOf("http")!=0){src=lpMTagConfig.lpProtocol+"://"+lpMTagConfig.lpServer+src+"?site="+lpMTagConfig.lpNumber;}
else{if(src.indexOf("site=")<0){if(src.indexOf("?")<0)src=src+"?";else src=src+"&";src=src+"site="+lpMTagConfig.lpNumber;}};var s=document.createElement("script");s.setAttribute("type","text/javascript");s.setAttribute("charset","iso-8859-1");s.setAttribute("src",src);document.getElementsByTagName("head").item(0).appendChild(s);}}
lpMTagConfig.calculateSentPageTime=function(){var t=(new Date()).getTime()-lpMTagConfig.pageStartTime;lpAddVars('page','pageLoadTime',Math.round(t/1000)+" sec");};if(window.attachEvent)window.attachEvent("onload",lpMTagConfig.calculateSentPageTime);else window.addEventListener("load",lpMTagConfig.calculateSentPageTime,false);if(window.attachEvent)window.attachEvent("onload",lpAddMonitorTag);else window.addEventListener("load",lpAddMonitorTag,false);if(typeof(lpMTagConfig.pageVar)=="undefined")lpMTagConfig.pageVar=new Array();if(typeof(lpMTagConfig.sessionVar)=="undefined")lpMTagConfig.sessionVar=new Array();if(typeof(lpMTagConfig.visitorVar)=="undefined")lpMTagConfig.visitorVar=new Array();if(typeof(lpMTagConfig.onLoadCode)=="undefined")lpMTagConfig.onLoadCode=new Array();if(typeof(lpMTagConfig.dynButton)=="undefined")lpMTagConfig.dynButton=new Array();function lpAddVars(scope,name,value){if(name.indexOf('OrderTotal')!=-1||name.indexOf('OrderNumber')!=-1){if(value==''||value==0)return;else lpMTagConfig.sendCookies=false}
value=lpTrimSpaces(value.toString());switch(scope){case"page":lpMTagConfig.pageVar[lpMTagConfig.pageVar.length]=escape(name)+"="+escape(value);break;case"session":lpMTagConfig.sessionVar[lpMTagConfig.sessionVar.length]=escape(name)+"="+escape(value);break;case"visitor":lpMTagConfig.visitorVar[lpMTagConfig.visitorVar.length]=escape(name)+"="+escape(value);break;}}
function onloadEMT(){var LPcookieLengthTest=document.cookie;if(lpMTag.lpBrowser=="IE"&&LPcookieLengthTest.length>1000){lpMTagConfig.sendCookies=false;}}
function lpTrimSpaces(stringToTrim){return stringToTrim.replace(/^\s+|\s+$/g,"");}
function lpSendData(varscope,varname,varvalue){if(typeof(lpMTag)!="undefined"&&typeof(lpMTag.lpSendData)!="undefined")
lpMTag.lpSendData(varscope.toUpperCase()+"VAR!"+varname+"="+varvalue,true);}
lpMTagConfig.ifVisitorCode=[];try{if(typeof(lpUnit)=="undefined")var lpUnit='istock';if(typeof(lpLanguage)=="undefined")var lpLanguage='english';if(typeof(lpAddVars)!="undefined")lpAddVars("page","unit",lpUnit);if(typeof(lpAddVars)!="undefined")lpAddVars("session","language",lpLanguage);lpMTagConfig.defaultInvite="chat-"+lpUnit+"-"+lpLanguage;}catch(e){}
lpMTagConfig.onLoadCode[lpMTagConfig.onLoadCode.length]=onloadEMT;if(typeof(lpMTagConfig.dynButton)!="undefined"){lpMTagConfig.dynButton[lpMTagConfig.dynButton.length]={"name":"chat-"+lpUnit+"-"+lpLanguage+"-header","pid":"lpHeader","afterStartPage":true};lpMTagConfig.dynButton[lpMTagConfig.dynButton.length]={"name":"chat-"+lpUnit+"-"+lpLanguage+"-footer","pid":"lpFooter","afterStartPage":true};}
if(typeof(lpSignupFlag)!="undefined"){lpAddVars('session','SignupFlag',lpSignupFlag);}
if(typeof(lpLoginFlag)!="undefined"){lpAddVars('session','LoginFlag',lpLoginFlag);}
if(typeof(lpMemberName)!="undefined"){lpAddVars('session','MemberName',lpMemberName);}
if(typeof(lpMemberFirstName)!="undefined"){lpAddVars('session','MemberFirstName',lpMemberFirstName);}
if(typeof(lpMemberLastName)!="undefined"){lpAddVars('session','MemberLastName',lpMemberLastName);}
if(typeof(lpMemberType)!="undefined"){lpAddVars('session','MemberType',lpMemberType);}
if(typeof(lpAccountType)!="undefined"){lpAddVars('session','AccountType',lpAccountType);}
if(typeof(lpVisitorID)!="undefined"){lpAddVars('session','VisitorID',lpVisitorID);}
if(typeof(lpFrameFlag)!="undefined"){lpAddVars('page','FrameFlag',lpFrameFlag);}
if(typeof(lpDivFlag)!="undefined"){lpAddVars('page','DivFlag',lpDivFlag);}
if(typeof(lpCurrencyType)!="undefined"){lpAddVars('page','CurrencyType',lpCurrencyType);}
if(typeof(lpCartTotal)!="undefined"){lpAddVars('page','CartTotal',lpCartTotal);}
if(typeof(lpCartCount)!="undefined"){lpAddVars('page','CartCount',lpCartCount);}
if(typeof(lpOfferCode)!="undefined"){lpAddVars('page','OfferCode',lpOfferCode);}
if(typeof(lpCheckoutTypeSelected)!="undefined"){lpAddVars('page','CheckoutTypeSelected',lpCheckoutTypeSelected);}
if(typeof(lpUpload_OrderNumber)!="undefined"){lpAddVars('page','Upload_OrderNumber',lpUpload_OrderNumber);}
if(typeof(lpDownload_OrderNumber)!="undefined"){lpAddVars('page','Download_OrderNumber',lpDownload_OrderNumber);}
if(typeof(lpDownload_OrderTotal)!="undefined"){lpAddVars('page','Download_OrderTotal',lpDownload_OrderTotal);}
if(typeof(lpCreditWalletValue)!="undefined"){lpAddVars('page','CreditWalletValue',lpCreditWalletValue);}
if(typeof(lpStoreCheckout_CurrencyType)!="undefined"){lpAddVars('page','StoreCheckout_CurrencyType',lpStoreCheckout_CurrencyType);}
if(typeof(lpStoreCheckout_OrderCount)!="undefined"){lpAddVars('page','StoreCheckout_OrderCount',lpStoreCheckout_OrderCount);}
if(typeof(lpStoreCheckout_OrderTotal)!="undefined"){lpAddVars('page','StoreCheckout_OrderTotal',lpStoreCheckout_OrderTotal);}
if(typeof(lpStoreCheckout_OrderNumber)!="undefined"){lpAddVars('page','StoreCheckout_OrderNumber',lpStoreCheckout_OrderNumber);}
if(typeof(lpStoreCheckoutStage)!="undefined"){lpAddVars('page','StoreCheckoutStage',lpStoreCheckoutStage);}