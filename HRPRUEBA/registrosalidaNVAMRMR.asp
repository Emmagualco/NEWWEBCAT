<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
dni=request.querystring("dni")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.open "Select * from hojarutamr where numviaje=" & request.querystring("hr") & " and fechallegada is not null and fechasalida is null order by idhojaruta desc",cn
if rs.eof=false then
	idhr=rs("idhojaruta")
	rs("fechasalida")=date()
	rs("horasalida")=time()
	rs("observacionesconductor")=ucase(trim(request.form("obser")))
	rs.update
end if
rs.close


dni=request.querystring("dni")
clave=request.querystring("clave")
tipo=request.querystring("tipo")
response.redirect("listarhr1nvamr.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)%>