<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
set rs7=server.createobject("ADODB.Recordset")
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cade="select * from hojaruta where idhojaruta=" & request.querystring("idhr") & " order by idhojaruta desc"
set rs1=cn.execute(cade)
If Not IsNull(RS1.Fields("horallegada")) And Not IsNull(RS1.Fields("horasalida")) Then
	numviaje=rs1("numviaje")
	remitente=trim(rs1("destinatario"))
	hojaderuta=trim(rs1("hojaderuta"))
	tipo=trim(rs1("tipo"))
	cliente=trim(rs1("cliente"))
	cargad=rs1("cargadestino")
	remitocargad=trim(rs1("remitocargadestino"))
	comoviajas=trim(rs1("quellevadestino"))
	rs1.close
	rs.CursorType = adOpenKeyset
	rs.LockType = adLockOptimistic
	rs.open "hojaruta",cn
	rs.addnew
	rs("numviaje")=numviaje
	rs("remitente")=remitente
	cad="select idtercero from terceros where rstercero='" & remitente & "'"	
	set rs4=cn.execute(cad)
	idremitente=rs4("idtercero")
	rs("idremitente")=idremitente
	rs("destinatario")=request.form("destino")
	cad="select idtercero from terceros where rstercero='" & request.form("destino") & "'"	
	set rs5=cn.execute(cad)
	iddestinatario=rs5("idtercero")
	rs("iddestinatario")=iddestinatario
	rs("origen")="PFCBA"
	rs("destino")="PFCBA"
	rs("plataformaorigen")="CORDOBA"
	rs("plataformadestino")="CORDOBA"
	rs("hojaderuta")=hojaderuta
	rs("tipo")=tipo
	rs("cargaorigen")=cargad
	rs("remitocargaorigen")=remitocargad
	rs("quellevaorigen")=comoviajas
	rs("cargadestino")=cargad
	rs("remitocargadestino")=remitocargad
	rs("quellevadestino")=comoviajas
	rs("cliente")=cliente
	rs("fechanuevotramo")=date
	rs("horanuevotramo")=time
	rs.update
	rs.close
	dni=request.querystring("dni")
	clave=request.querystring("clave")
	tipo=request.querystring("tipo")
	response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
else
	dni=request.querystring("dni")
	clave=request.querystring("clave")
	tipo=request.querystring("tipo")
	response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
end if%>