<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
tipo=request.querystring("tipo")
dni=request.querystring("dni")
clave=request.querystring("clave")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
if tipo="MILK RUN" then
	rs.open "select * from hojaruta where numviaje=" & request.querystring("hr") & " and remitente='" & request.querystring("vdesde") & "' and destinatario='" & request.querystring("vhasta") & "' order by idhojaruta",cn
	rs("cargaDestino")=rs("cargaDestino")+int(request.form("cargad"))
	if rs("remitocargadestino")="" or isnull(rs("remitocargadestino")) then
		rs("remitocargadestino")=request.form("remitocargad")
	else
		rs("remitocargadestino")=trim(rs("remitocargadestino"))&"/ "& trim(request.form("remitocargad"))
	end if
	rs("quellevadestino")=request.form("contenidod")

	rs("carga")=rs("carga")+int(request.form("cargad"))
	if rs("remitocarga")="" or isnull(rs("remitocarga")) then
		rs("remitocarga")=request.form("remitocargad")
	else
		rs("remitocarga")=trim(rs("remitocarga"))& "/ " & trim(request.form("remitocargad"))
	end if
	rs("quellevacarga")=request.form("contenidod")
	rs.update
	rs.close
end if

if tipo="MOVIMIENTOS INTERNOS" then
	rs.open "select * from hojaruta where numviaje=" & request.querystring("hr") & " and remitente='" & request.querystring("vdesde") & "' and destinatario='" & request.querystring("vhasta") & "' order by idhojaruta desc",cn
	rs("cargaDestino")=rs("cargaDestino")+int(request.form("cargad"))
	if rs("remitocargadestino")="" or isnull(rs("remitocargadestino")) then
		rs("remitocargadestino")=request.form("remitocargad")
	else
		rs("remitocargadestino")=trim(rs("remitocargadestino"))&"/ "& trim(request.form("remitocargad"))
	end if
	rs("quellevadestino")=request.form("contenidod")

	rs("carga")=rs("carga")+int(request.form("cargad"))
	if rs("remitocarga")="" or isnull(rs("remitocarga")) then
		rs("remitocarga")=request.form("remitocargad")
	else
		rs("remitocarga")=trim(rs("remitocarga"))& "/ " & trim(request.form("remitocargad"))
	end if
	rs("quellevacarga")=request.form("contenidod")
	rs.update
	rs.close
end if

response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)%>