<!--#include virtual="/lib/include/global.inc" -->

<%
Dim intContractType : intContractType = getRequest("ContractType", CONST_DECODE_NO)	'1:一般短期合約 2: Learning Card
Dim intClientSn : intClientSn = getRequest("ClientSn", CONST_DECODE_NO)
Dim strRandomKey : strRandomKey = getRequest("RandomKey", CONST_DECODE_NO)
Dim strContractSn : strContractSn = getRequest("ContractSn", CONST_DECODE_NO)
Dim arrParameter
Dim strCname : strCname = ""
Dim intClassNum : intClassNum = 0
Dim intPeriod : intPeriod = now
Dim bolAgree : bolAgree = false
Dim datContractStartDate : datContractStartDate = now
Dim datEndDate

'response.write "intContractType: " & intContractType & "<br/>"
'response.write "intClientSn: " & intClientSn & "<br/>"
'response.write "strRandomKey: " & strRandomKey & "<br/>"
'response.write "strContractSn: " & strContractSn & "<br/>"
'response.end

if(isEmptyOrNull(strContractSn)) then
	strContractSn = getRequest("contract_sn", CONST_DECODE_NO)
end if

if(not isEmptyOrNull(intClientSn)) then
	strSql="SELECT cname FROM dbo.client_basic WITH (NOLOCK) WHERE sn = @ClientSn"
	arrParameter = Array(intClientSn)
	arr_result = excuteSqlStatementRead(strSql,arrParameter,CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then
			strCname = arr_result(0, 0)
		end if
	end if
end if
if(not isEmptyOrNull(strRandomKey) or not isEmptyOrNull(intClientSn)) then
	strSql = "SELECT TOP 1 RandomKey1, class , period , used FROM dbo.RandomKey WITH(NOLOCK) WHERE RandomKey1 = @RandomKey1 OR ClientSn = @ClientSn "
	arrParameter = Array(strRandomKey,intClientSn)
	arr_result = excuteSqlStatementRead(strSql,arrParameter,CONST_VIPABC_RW_CONN)
	'response.write g_str_sql_statement_for_debug
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then
			strRandomKey = arr_result(0, 0)
			intClassNum = arr_result(1, 0)
			intPeriod = arr_result(2, 0)
		end if
	end if
end if
if(not isEmptyOrNull(intPeriod)) then
	datEndDate = DateAdd ("d" , intPeriod , now) 
end if

if(not isEmptyOrNull(strContractSn)) then
	strSql="SELECT cagree, sdate FROM dbo.client_temporal_contract WITH (NOLOCK) WHERE sn = @ContractSn"
	arrParameter = Array(strContractSn)
	arr_result = excuteSqlStatementRead(strSql,arrParameter,CONST_VIPABC_RW_CONN)
	if (isSelectQuerySuccess(arr_result)) then
		if (Ubound(arr_result) >= 0) then
			bolAgree = arr_result(0, 0)
			datContractStartDate = arr_result(1, 0)
		end if
	end if
	if(not isEmptyOrNull(intPeriod)) then
		datEndDate = DateAdd ("d" , intPeriod , datContractStartDate) 
	end if
end if
%>
<script type="text/javascript">
var strContractType = "<%=intContractType%>";
var strClientSn = "<%=intClientSn%>";
var strRandomKey = "<%=strRandomKey%>";
	//列印的function
	function printElem(options)
	{
		$('#main_content').printElement(options);
	}

	$(document).ready(function(){
		//如果選了列印，列印區塊檔案
		$("#print_contract").click(function() {
				 printElem({overrideElementCSS: ['/css/css_cn_test.css']});
			 });

	    //同意
	    $("#agree").click(function () {
	        document.f1.submit();
	    });
	});
</script>
<!-- Menber Content Start -->
<div class="main_con" id="main_content">
<!--內容start-->
<div class="terms_main2">
	<form name="f1" id="f1" method="post" action='contract_agree.asp'>
	
    <div class="line01"></div>
    <div class="title2 t-center big">VIPABC短期会员使用规则</div>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td colspan="2">TutorGroup麦奇教育集团（即VIPABC.com，为香港注册且运营网站，并于中国地区指定上海麦奇教育信息咨询有限公司为授权经销商，简称本公司），会员：<%=strCname%>（以下简称会员），兹就会员加入本公司VIPABC网络数位学习系统，并于会籍有效期限内，依据本合同书所订的条款参与网络咨询课程及服务，双方同意就相关合同内容约定如下：</td>
			</tr>
			<tr>
				<td width="4%" align="left" valign="top">一、</td>
				<td width="96%" align="left">
					会员使用「VIPABC」网络咨询课程及服务前，同意自行配备上网所需之各项计算机软硬设备，并负担接续因特网及电话等相关费用。于使用网站服务时，应遵守中华人民共和国相关法令及因特网之国际使用惯例与礼节。
				</td>
			</tr>
			<tr>
				<td align="left" valign="top">二、</td>
				<td align="left">					
					本公司于契约有效期限内，应提供会员在线系统之使用安装服务，如系会员个人计算机之软硬件或电信业者因特网系统等障碍因素，则不在提供服务范围之内。会员亦不得以网络带宽或其他环境因素，致使其视讯质量不良时，向本公司要求或主张任何之权益。
				</td>
			</tr>
			<tr>
				<td align="left" valign="top">三、</td>
				<td align="left">会员如遇网络数字学习系统无法正常使用或有异动之情形，应迅速与客服人员联系。</td>
			</tr>
			<tr>
				<td align="left" valign="top">四、</td>
				<td align="left">会员基本规则</td>
			</tr>
            <tr>
                <td align="left" valign="top"></td>
				<td align="left">
                    <table>
                        <tr>
                            <td width="2%" align="left" valign="top">1.</td>
				            <td width="98%" align="left">
					            会员于本契约存续期间得使用<%=intClassNum%>堂小班制咨询服务课程或大会堂课程（不得使用一对一咨询服务课程）。本契约存续期间为自<%=Year(datContractStartDate)%>年<%=Month(datContractStartDate)%>月<%=Day(datContractStartDate)%>日起，至<%=Year(datEndDate)%>年<%=Month(datEndDate)%>月<%=Day(datEndDate)%>日止，为期<%=intPeriod%>天，不得申请展延。会员理解于契约存续期间内，无论堂数使用之情形，契约时间与咨询堂数等同重要。契约期满或虽未届止但堂数于期限内使用完毕时，本公司对于会员之服务即视为终止，会员不得再有所主张或请求。
				            </td>
                        </tr>
                        <tr>
                            <td width="2%" align="left" valign="top">2.</td>
				            <td width="98%" align="left">
					            本公司短期入会方案系以免费或优惠价格提供会员，因而会员开始使用本顾问服务后，不能以个人因素解除、终止合约并要求相关费用之退还，亦不得将会籍资格或账号转让他人使用。
				            </td>
                        </tr>
                        <tr>
                            <td width="2%" align="left" valign="top">3.</td>
				            <td width="98%" align="left">
					            会员必须于咨询前24小时在线订位，以确认咨询时间。
				            </td>
                        </tr>
                        <tr>
                            <td width="2%" align="left" valign="top">4.</td>
				            <td width="98%" align="left">
					            若因事无法出席，请于在线咨询开始前4小时取消，否则预约咨询之堂数将被计入成为使用堂数。
				            </td>
                        </tr>
                        <tr>
                            <td width="2%" align="left" valign="top">5.</td>
				            <td width="98%" align="left">
					            为尊重自身及其他会员权益，预约后请务必准时进入在线咨询室。如有逾时之情形，会员将无法进入该堂之咨询，并就该预约堂数仍为计入使用堂数。
				            </td>
                        </tr>
                        <tr>
                            <td width="2%" align="left" valign="top">6.</td>
				            <td width="98%" align="left">
					            本公司提供多国籍语言顾问进行咨询服务，但无法指定特定顾问进行咨询服务。
				            </td>
                        </tr>
                        <tr>
                            <td width="2%" align="left" valign="top">7.</td>
				            <td width="98%" align="left">
					            网络预约：www.vipabc.com<br/>
                                客服信箱：vipservices@vipabc.com<br/>
                                本公司中国区授权经销商客户服务专线： <%=CONST_SERVICE_PHONE2%> <br/>
				            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
				<td align="left" valign="top">五、</td>
				<td align="left">下列行为系本公司VIPABC网站及在线咨询服务禁止之行为，但不以这些行为为限；会员如有不当行为或违反相关服务条款时，本公司得视违反情节轻重或影响层面，而予以暂停或撤销其会籍，除不得要求退还未使用堂数之费用外，本公司并保留该损害赔偿之请求权利。</td>
			</tr>
            <tr>
                <td align="left" valign="top"></td>
				<td align="left">
                    <table>
                        <tr>
                            <td width="3%" align="left" valign="top">1.</td>
				            <td width="97%" align="left">
					            会员入侵或破坏本公司网站内部各项设施，或任何有损及本公司名誉之行为。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">2.</td>
				            <td width="97%" align="left">
					            利用在线视讯系统施行：暴露、猥亵、性暗示或其他不雅或足以影响本公司咨询服务进行之行为。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">3.</td>
				            <td width="97%" align="left">
					            利用在线系统施行言语或文字骚扰、谩骂、攻击，或其他不受欢迎之行为。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">4.</td>
				            <td width="97%" align="left">
					            透过可供文件传输软件，传送非法软件或具破坏或影响他人权益之情形。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">5.</td>
				            <td width="97%" align="left">
					            在咨询进行中或服务系统上，有销售营利、政治请托、宗教宣扬或其他与课程内容无关之行为。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">6.</td>
				            <td width="97%" align="left">
					            为维护双方之权益，本公司禁止职员及语言顾问与会员在非咨询服务时间或利用其它以外之方式互动（例如；实时通讯等软硬件设施）。如在此规范外之一切行为或意外发生，概与本公司无涉。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">7.</td>
				            <td width="97%" align="left">
					            利用公开竞标、拍卖或以其他方式租、借、转卖本公司会籍资格及服务项目。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">8.</td>
				            <td width="97%" align="left">
					            私自提供账号、密码予他人使用本公司服务项目。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">9.</td>
				            <td width="97%" align="left">
					            会员咨询之课程内容与教材属于本公司所有，会员不得擅将咨询课程之录音、录像或教材 内容，公开播送或贩卖。
				            </td>
                        </tr>
                        <tr>
                            <td width="3%" align="left" valign="top">10.</td>
				            <td width="97%" align="left">
					            其他不当或违反法令及公序良俗之行为。
				            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
				<td align="left" valign="top">六、</td>
				<td align="left">会员同意咨询期间所参与之相关录像、录音及文字图文件等内容，无偿供本公司系列产品之编辑与研发、会员教学观摩、业务推广、产品发表或销售之使用。本契约有效期间届至或解除契约或提前终止后，本条款仍继续有效。</td>
			</tr>
            <tr>
				<td align="left" valign="top">七、</td>
				<td align="left">双方同意本规则内容之一部或部分判定无效时，并未影响契约其它条款之效力。为免生争议及认定上歧异，任何口头非文字载明之承诺、条件，均视为无效约定。</td>
			</tr>
            <tr>
				<td align="left" valign="top">八、</td>
				<td align="left">就本约未尽事宜，双方同意依本公司内部或网站上有关之公告、说明、服务条款补充，并依中华人民共和国法律与其它相关法令解释并适用之。双方同意因本规则发生争议时，应本诚信原则善尽最大努力协议处理。如因而争讼，双方合意以本公司所在地之人民法院为第一审管辖法院。</td>
			</tr>
            <tr>
				<td align="left" valign="top">九、</td>
				<td align="left">本契约书经会员确认并同意后，将以电子文件方式留存于本公司系统中，以为双方契约签署及遵循之依据。</td>
			</tr>
		</table>
		<div class="dode"></div>
		<p class="t-center">
			<input type="button" value="+ 打印合约" id="print_contract" class="btn_1 m-left10" />
			<%if("1" <> bolAgree) then%>
			<input type="button" name="agree_or_yet" id="agree" value="+ <%=getWord("CONTRACT_116")%>"
            class="btn_1 m-left10" />
			<!--<input type="button" name="agree_or_yet" id="disagree" value="+ <%=getWord("CONTRACT_117")%>"
				class="btn_1 m-left10" />-->
			<%end if%>
        <input type="button" name="print_or_prepage" style="display: none" id="print_contract"
            value="+ <%=getWord("CONTRACT_118")%>" class="btn_1 m-left10" onclick="print();" />
		</p>
    </form>
</div>
<!--內容end-->
</div>
<!-- Main Content End -->
