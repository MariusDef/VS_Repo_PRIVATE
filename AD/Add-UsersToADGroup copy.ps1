Import-Module ActiveDirectory


$Users = "WIDAP245E002","wibrieskornj","wioeder","WIDAP820S008","wilorscheidj","wigrueberc","wilochm","widigischau","wihoellers","wischenkelbk","wiewerk","WIDAP280F005","exbormannn","wistude","wistude","WIDAP345O003","wiemgross","wihertwigs","dewinkler","wischoeneber","wibuchholzb","wilitz","WIDAP240D007","wikurtenbacc","wischwippert","winerkewitz","wiisnuc09","WILOGIMAT1","wiaxera","wikonradr","wiprassel","wischumachef","wiwenzelmann","wipreukschaj","WIDAP210J002","wiwenzelmann","wibernacchia","widiekmann","winix","wimeiera","WIDAP240C001","WIDAP307G002","wibuchmuellj","WITAP520L009","wilimbachd","witerstegent","WIDAP245C004","wiblankc"
#$Users = "wiforsbachi","wipinnenj","WITAP520O007","wihagemannn","wikoschel","wijonen","wiover","wiisnuc08","wikremer","wihurtenbacp","WIDAP820S003","wigerntken","wimanroths","wischwippert","wihenschelk","wikoppa","wiludwigp","wigta01","WIDAP240C008","wilorsem","wiriedela","wigaertners","wiulpinsk","wiannahme03","widavidm","wiherbstv","wihesselers","WIDAP550E001","wiisnuc28","delimbergr","wikardex518pw","wimayr","wiexzaiserm","WIDAP170B025","wischulz","wischaefert","wibreithausg","wischarfensa","WIDAP230C001","wibfallg","widuesseldoc","wiexkluthc","wischmidtp","wischmitzhud","wikrusek","wischenkelbk","WIDAP170C044","wischiffera","wisalzv","wipottm"
$Computers = "WINUC295$","WIPC3D288$","WIPC3D287$","WINUC421$","WIPC3D211$","WINB0527$","WIPC1109$","WINUC31$","WIPC1166$","WIPC1210$","WITAB47$","WINUC256$","WINB0960$","WINB1071$","WIPC3D281$","WINUC332$","WINB0798$","WINB0597$","WIPC0838$","WINB0580$","WIPC3D152$","WIPC1113$","WINUC310$","WIPC0738$","WINB1210$","WINB1241$","WINUC109$","WIPC1005$","WINB0831$","WINB1103$","WIPC1127$","WIPC0674$","WIPC3D240$","WIPC3D337$","WINUC143$","WINB1295$","WINB1144$","WINB1143$","WINB1093$","WIPC1001$","WINUC271$","WINUC178$","WIPC0765$","WITAB108$","WIVMDL01$","WIPC1015$","WINUC280$","WINB1299$"
#$Computers = "WINB0958$","WINB1317$","WITAB60$","WINB0890$","WINB1232$","WINB1223$","WIPC0844$","WINUC108$","WIPC0820$","WIVMPH01$","WINUC405$","WIPC0964$","WIPC1117$","WIPC3D329$","WIPC1118$","WIPC3D279$","WIPC0910$","WIPC0716$","WINUC281$","WIPC1159$","WINB0772$","WINB0742$","WINB1068$","WIPC0998$","WIPC0900$","WITAB08$","WINB1234$","WINUC419$","WINUC128$","WITAB27$","WINUC174$","WINB1083$","WIPC3D247$","WINUC221$","WIPC3D238$","WINB0922$","WINB1032$","WINB0661$","WINUC283$","WIPC0690$","WINB1157$","WIPC1141$","WITAB114$","WINB1212$","WIPC0875$","WIPC3D239$","WINUC204$","WINB0858$","WIPC0701$","WINB0632$"

$Computers | Out-Null


Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Browser-User-Apply" -Members $Users -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Office-User-Apply" -Members $Users -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Windows-User-Apply" -Members $Users -Server "wirtgen33"

Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Browser-Computer-Apply" -Members $Computers -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Office-Computer-Apply" -Members $Computers -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Windows-Computer-Apply" -Members $Computers -Server "wirtgen33"
