Import-Module ActiveDirectory


$Users = "wiexwarkenl","exwildemannr","wipinnenj","wineue","wifergenj","wiwagenerc","wigauera","wihaehnk","wibernacchia","wischmohlc","wiholls","wihackerd","wikolbek","wikristat","wiaudings","wibeckenstrm","wiblocks","wiengels","wivonkopposj","wikreuterm","wipuetzt","wischmitz","widuesseldoc","wigerstmannl","wifischerj","wilahrt","wiwertenbrud","wijaxt","wikyeku","wisalzg","wischmohlc","wiweberk","wischieferp","wigauera","wiodenbacha","wischneiderj","wischmitzt","withorunm","witerstegent","wiheigeln","wisuelzeran","wistockm","wihoellers","wiloehlef","wireuter","demaubachc","wireufelsc","wiwesterp","wilangen","wilissc"
$Computers = "winb0927$","winb0549$","winb0609$","winb0633$","winb0668$","winb0691$","winb0734$","winb0771$","winb0773$","winb0774$","winb0808$","winb0809$","winb0923$","winb0926$","winb1021$","winb1022$","winb1023$","winb1024$","winb1150$","winb1153$","winb1154$","winb1155$","winb1157$","winb1160$","winb1162$","winb1163$","winb1171$","wipc0567$","wipc0569$","wipc0579$","wipc0661$","wipc0717$","wipc0720$","wipc0879$","wipc0882$","wipc0883$","wipc0890$","wipc0897$","wipc1015$","wipc1012$","wipc1017$","wipc1094$","wipc1166$","wipc1027$","wipc1099$","wipc1228$","wipc1240$","wipc3d229$","wipc3d235$","wipc3d270$"
$Computers | Out-Null


Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Browser-User-Apply" -Members $Users -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Office-User-Apply" -Members $Users -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Windows-User-Apply" -Members $Users -Server "wirtgen33"

Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Browser-Computer-Apply" -Members $Computers -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Office-Computer-Apply" -Members $Computers -Server "wirtgen33"
Add-ADGroupMember -Identity "WIGOrg_MSBaseline-Windows-Computer-Apply" -Members $Computers -Server "wirtgen33"
