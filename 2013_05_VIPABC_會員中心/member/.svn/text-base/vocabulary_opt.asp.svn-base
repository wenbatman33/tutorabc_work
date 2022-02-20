
<%
Dim str_opt_type,str_vocabulary
Dim obj_mtl_voc_opt
Dim int_client_sn
Dim bol_process_state
str_opt_type = Request("voc_opt_type")
str_vocabulary = Request("voc")
int_client_sn = Request("data")
if(str_opt_type = "" or str_vocabulary = "" or int_client_sn = "") then
	Response.write("failed")
	Response.end
end if
Set obj_mtl_voc_opt = Server.CreateObject("TutorLib.MaterialOpt.MaterialVocabularyOpt")
if(str_opt_type = "del") then
	bol_process_state = obj_mtl_voc_opt.delfromVocabularyBank(int_client_sn,str_vocabulary)
else
	bol_process_state = obj_mtl_voc_opt.addToVocabularyBank(int_client_sn,str_vocabulary)
end if
Set obj_mtl_voc_opt = Nothing
if(bol_process_state = True) then
	Response.write("success")
else
	Response.write("failed")
end if

%>