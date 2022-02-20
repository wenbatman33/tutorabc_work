for(var i = 0; i < 16; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
gv_vAlignTable['u3'] = 'center';gv_vAlignTable['u12'] = 'center';u15.tabIndex = 0;

u15.style.cursor = 'pointer';
$axure.eventManager.click('u15', function(e) {

if (true) {

	self.location.href=$axure.globalVariableProvider.getLinkUrl('Record_Video.html');

}
});
u4.tabIndex = 0;

u4.style.cursor = 'pointer';
$axure.eventManager.click('u4', u4Click);
InsertAfterBegin(document.body, "<div class='intcases' id='u4LinksClick'></div>")
var u4LinksClick = document.getElementById('u4LinksClick');
function u4Click(e) 
{
windowEvent = e;


	ToggleLinks(e, 'u4LinksClick');
}

InsertBeforeEnd(u4LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u4Clicku28082ea70fc64a348f4755f30de999a1(event)'>未登入</div>");
function u4Clicku28082ea70fc64a348f4755f30de999a1(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_未登_.html');

	ToggleLinks(e, 'u4LinksClick');
}

InsertBeforeEnd(u4LinksClick, "<div class='intcaselink' onmouseout='SuppressBubble(event)' onclick='u4Clickud134814b745e41978ca79e56cf805d47(event)'>已登入</div>");
function u4Clickud134814b745e41978ca79e56cf805d47(e)
{

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_已登_.html');

	ToggleLinks(e, 'u4LinksClick');
}
gv_vAlignTable['u8'] = 'center';gv_vAlignTable['u10'] = 'center';gv_vAlignTable['u1'] = 'center';gv_vAlignTable['u14'] = 'center';gv_vAlignTable['u6'] = 'center';