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
rs.open "select * from hojaruta where numviaje=" & request.querystring("hr") & " and remitente='" & request.querystring("vdesde") & "' and destinatario='" & request.querystring("vhasta") & "' order by idhojaruta desc",cn
if request.form("hasta")="" then
	rs("destinatario")=request.querystring("vhasta")
else
	rs("destinatario")=request.form("hasta")
end if
rs.update
rs.close
response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)%>