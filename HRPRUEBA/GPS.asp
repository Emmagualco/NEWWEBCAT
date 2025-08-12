<%
Function GetXML(url)
Dim xmlobj
Set xmlobj = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
xmlobj.setTimeouts 30000, 30000, 30000, 30000
xmlobj.Open "GET", url, False
xmlobj.send()
If xmlobj.status = 200 Then
    GetXML = xmlobj.responseXML.xml
Else
    Response.Write "The geocoding server could not be reached."
    Response.End
End If
End Function

address = "Mariano Benitez 1436, X5000"


url="http://maps.google.com/maps/geo?output=xml&key=XXXXXXX &q="&Server.URLEncode(address)






Response.Write "You entered: "&address&"<br />"

xml = GetXML(url) 'Function to return raw XML as text

if InStr(xml,"<coordinates>")>0 then
coords = split(xml,"<coordinates>") 'Get everything after the opening coordinates tag
coords2 = split(coords(1),"</coordinates>") 'Get everything before the ending coordinates tag
coordinatesSplit = split(coords2(0),",") 'split it at the commas
lng = coordinatesSplit(0) 'The first value is the longitude
lat = coordinatesSplit(1) 'The second value is the latitude

Response.Write "The geo-coded coordinates are: "&lat&" "&lng
else
'No coordinates were returned
Response.Write "The address could not be geocoded."
Response.End
end if
%>