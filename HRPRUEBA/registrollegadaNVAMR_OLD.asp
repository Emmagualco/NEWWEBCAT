<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
dni=request.querystring("dni")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.open "Select * from hojaruta where numviaje=" & request.querystring("hr") & " and fechallegada is null order by idhojaruta",cn
if rs.eof=false then
	idhr=rs("idhojaruta")
	rs("fechallegada")=date()
	rs("horallegada")=time()
	rs.update
end if
rs.close

rs.open"select * from viajes where numviaje=" & request.querystring("hr"),cn
if rs("etapa")="G" then
	rs("etapa")="A"
	rs.update
end if
rs.close

dni=request.querystring("dni")
clave=request.querystring("clave")
tipo=request.querystring("tipo")
response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)%>