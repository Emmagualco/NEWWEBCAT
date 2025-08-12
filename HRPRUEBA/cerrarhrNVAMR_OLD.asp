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
cadena="select fechasalida,horasalida from hojarutamr where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
set rs2=cn.execute(cadena)
if rs2.eof=false then
	fechasalida=rs2("fechasalida")
	horasalida=trim(rs2("horasalida"))
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

	rs.open"select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
	rs("horasalida")=horasalida
	rs("fechasalida")=fechasalida
	rs.update
	rs.close
	response.redirect("loginhrnva.html")
else
	cadena="select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
	set rs6=cn.execute(cadena)
	if rs6.eof=false then
		fechasalida=rs6("fechasalida")
		horasalida=trim(rs6("horasalida"))
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

		rs.open"select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
		rs("horasalida")=horasalida
		rs("fechasalida")=fechasalida
		rs.update
		rs.close
		response.redirect("loginhrnva.html")
	end if
end if%>