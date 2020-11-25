Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs =oWMP.cdromCollection
if colCDROMs.Count >= 1 then
for i = o to colCDROMs.Count - 1
colCDROMs.Item(i).Eject
colCDROMs.Item(i).Eject
Next' cdrom
End if