<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
cadena="select fechallegada,horallegada from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta"
set rs1=cn.execute(cadena)
fechallegada=rs1("fechallegada")
horallegada=trim(rs1("horallegada"))
cadena="select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
set rs2=cn.execute(cadena)
if not isnull(rs2("fechasalida")) then
	fechasalida=rs2("fechasalida")
else
	fechasalida=date()
end if
if not isnull(rs2("horasalida")) then
	horasalida=trim(rs2("horasalida"))
else
	horasalida=time()
end if
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.open "Select * from viajes where numviaje=" & request.querystring("hr"),cn
rs("fechallegada")=fechasalida
rs("horallegada")=horasalida
rs("fechasalida")=fechallegada
rs("horasalida")=horallegada
rs("etapa")="C"
rs.update
rs.close

rs1.CursorType = adOpenKeyset
rs1.LockType = adLockOptimistic
rs1.open "select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc",cn
rs1("fechasalida")=fechasalida
rs1("horasalida")=horasalida
rs1.update
rs1.close

response.redirect("loginhrnva.html")%>