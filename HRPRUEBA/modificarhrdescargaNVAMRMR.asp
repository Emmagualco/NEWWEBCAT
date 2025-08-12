<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
tipo=request.querystring("tipo")
dni=request.querystring("dni")
clave=request.querystring("clave")
cantdescarga=request.form("descargad")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.open "select * from hojarutamr where numviaje=" & request.querystring("hr") & " and remitente='" & request.querystring("vdesde") & "' and destinatario='" & request.querystring("vhasta") & "' order by idhojaruta desc",cn
rs("Descarga")=request.form("descargad")
rs("remitodescarga")=request.form("remitodescargad")
rs("quellevadescarga")=request.querystring("quellevacarga")

carga=rs("carga")
cargaorigen=rs("cargaorigen")
rs("cargadestino")=(carga+cargaorigen)-int(cantdescarga)
rs("remitocargadestino")=request.form("remitodescargad")
if (carga+cargaorigen)-cantdescarga=0 then
	rs("quellevadestino")="VACIO"
else
	rs("quellevadestino")=request.querystring("quellevacarga")
end if
rs.update
rs.close
response.redirect("listarhr1nvamr.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)%>