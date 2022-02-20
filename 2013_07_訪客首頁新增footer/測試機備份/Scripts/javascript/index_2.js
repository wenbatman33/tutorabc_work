//根據select 改變國碼
function changeNationalCode(obj)
{
	    var str_nation_code = obj.options[obj.selectedIndex].value;
		if(str_nation_code == "0" )
		{
		    //請選擇：do nothing
            $("#txt_phone_nationcode").val("");
		} 
		else if(str_nation_code == "-1")
		{	
		    //其他
			document.getElementById("txt_phone_nationcode").value = "";
			document.getElementById("txt_phone_nationcode").disabled="";
			document.getElementById("hdn_phone_nationcode").value = "";
		} else
		{
			document.getElementById("txt_phone_nationcode").value = str_nation_code;
			document.getElementById("txt_phone_nationcode").disabled="disabled";
			document.getElementById("hdn_phone_nationcode").value = str_nation_code;
		}
}

function changeNationalCode1(obj) {
    var str_nation_code = obj.options[obj.selectedIndex].value;
    if (str_nation_code == "0") {
        //請選擇：do nothing
        $("#txt_phone_nationcode1").val("");
    }
    else if (str_nation_code == "-1") {
        //其他
        document.getElementById("txt_phone_nationcode1").value = "";
        document.getElementById("txt_phone_nationcode1").disabled = "";
        document.getElementById("hdn_phone_nationcode1").value = "";
    } else {
        document.getElementById("txt_phone_nationcode1").value = str_nation_code;
        document.getElementById("txt_phone_nationcode1").disabled = "disabled";
        document.getElementById("hdn_phone_nationcode1").value = str_nation_code;
    }
}