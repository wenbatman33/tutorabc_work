for(var i = 0; i < 14; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
gv_vAlignTable['u4'] = 'center';gv_vAlignTable['u1'] = 'center';gv_vAlignTable['u9'] = 'center';gv_vAlignTable['u6'] = 'center';u2.tabIndex = 0;

u2.style.cursor = 'pointer';
$axure.eventManager.click('u2', function(e) {

if (true) {

	self.location.href=$axure.globalVariableProvider.getLinkUrl('首頁_已登_.html');

}
});
gv_vAlignTable['u11'] = 'center';gv_vAlignTable['u13'] = 'center';u7.tabIndex = 0;

u7.style.cursor = 'pointer';
$axure.eventManager.click('u7', function(e) {

if (true) {

	self.location.href=$axure.globalVariableProvider.getLinkUrl('Start_Session.html');

}
});
