<%set cn=server.createobject("ADODB.Connection")
set cn1=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
set rs7=server.createobject("ADODB.Recordset")
dni=request.querystring("dni")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
cad1="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn1.open(cad1)
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.open "Select * from hojaruta where numviaje=" & request.querystring("hr") & " and fechallegada is null order by idhojaruta",cn
if rs.eof=false then
	rs7.CursorType = adOpenKeyset
	rs7.LockType = adLockOptimistic
	rs7.open"select * from hojaruta where destinatario='" & trim(rs("remitente")) & "' and numviaje=" & request.querystring("hr"),cn
	rs7("fechasalidareal")=date()
	rs7("horasalidareal")=request.form("apps")
	rs7.update
	rs7.close
	idhr=rs("idhojaruta")
	rs("fechallegada")=date()
	rs("horallegada")=time()
	if isnull(rs("observacionesconductor")) or rs("observacionesconductor")="" then
		rs("observacionesconductor")=ucase(trim(request.form("obser")))
	end if
	rs("fechallegadareal")=date()
	rs("horallegadareal")=request.form("appll")	
	rs.update
	dni=request.querystring("dni")
	clave=request.querystring("clave")
	tipo=request.querystring("tipo")
	response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
end if%>