<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
usuario=request.querystring("usuario")
clave=request.querystring("clave")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
cade="select * from hojaruta where numviaje=" & request.querystring("numviaje") & " order by idhojaruta desc"
set rs1=cn.execute(cade)
fechallegada=rs1("fechasalidarasa")
horallegada=rs1("horasalidarasa")
rs1.close
rs1.CursorType = adOpenKeyset
rs1.LockType = adLockOptimistic
rs1.open"select * from hojaruta where numviaje=" & request.querystring("numviaje") & " order by idhojaruta desc",cn
rs1("quellevadestino")="VACÍO"
rs1("cargadestino")=0
if rs1("carga")>0 then
	rs1("descarga")=rs1("carga")
	rs1("quellevadescarga")=rs1("quellevacarga")
	rs1("remitodescarga")=trim(rs1("remitocarga"))
else
	if rs1("cargaorigen")>0 then
		rs1("descarga")=rs1("cargaorigen")
		rs1("quellevadescarga")=rs1("quellevaorigen")
		rs1("remitodescarga")=trim(rs1("remitocargaorigen"))
	end if		
end if
rs1.update

rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.open "Select * from viajes where numviaje=" & request.querystring("numviaje"),cn
rs("fechallegada")=fechallegada
rs("horallegada")=trim(horallegada)
rs("etapa")="C"
rs.update
rs.close
response.redirect("hrabiertas.asp" & "?usuario="& usuario & "&clave=" & clave)%> 