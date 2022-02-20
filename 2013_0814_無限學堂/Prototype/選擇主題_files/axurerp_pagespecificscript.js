for(var i = 0; i < 18; i++) { var scriptId = 'u' + i; window[scriptId] = document.getElementById(scriptId); }

$axure.eventManager.pageLoad(
function (e) {

});
gv_vAlignTable['u3'] = 'center';document.getElementById('u16_img').tabIndex = 0;

u16.style.cursor = 'pointer';
$axure.eventManager.click('u16', function(e) {

if (true) {

	SetPanelVisibility('u13','hidden','none',500);

	self.location.href=$axure.globalVariableProvider.getLinkUrl('進入諮詢室.html');

}
});
gv_vAlignTable['u12'] = 'center';gv_vAlignTable['u15'] = 'center';gv_vAlignTable['u8'] = 'center';gv_vAlignTable['u10'] = 'center';gv_vAlignTable['u17'] = 'center';gv_vAlignTable['u1'] = 'center';document.getElementById('u9_img').tabIndex = 0;

u9.style.cursor = 'pointer';
$axure.eventManager.click('u9', function(e) {

if (true) {

	SetPanelVisibility('u13','','none',500);

	SetPanelVisibility('u4','hidden','none',500);

}
});
gv_vAlignTable['u6'] = 'center';document.getElementById('u2_img').tabIndex = 0;

u2.style.cursor = 'pointer';
$axure.eventManager.click('u2', function(e) {

if (true) {

	SetPanelVisibility('u4','','none',500);

}
});
document.getElementById('u7_img').tabIndex = 0;

u7.style.cursor = 'pointer';
$axure.eventManager.click('u7', function(e) {

if (true) {

	SetPanelVisibility('u4','hidden','none',500);

}
});
