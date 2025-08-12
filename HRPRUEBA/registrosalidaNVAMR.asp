<%set cn=server.createobject("ADODB.Connection")
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
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.open "Select * from hojaruta where numviaje=" & request.querystring("hr") & " and fechallegada is not null and fechasalida is null order by idhojaruta",cn
if rs.eof=false then
	llego=rs("horallegada")
	ahora = Left(Time(), 5)
	estuvo=datediff("n",llego,ahora)
	if trim(rs("cliente"))="RENAULT ARGENTINA S.A.  PRODUCCIÓN" then
		if trim(rs("remitente"))<>trim(rs("destinatario")) then
			if trim(rs("remitente"))="PLAYA CAT" then
				if request.form("cargo")=checked then
					rs("cargado")=0
				else
					rs("cargado")=1
				end if
			else
				if request.form("cargo")=checked then
					rs("cargado")=0
				else
					rs("cargado")=1
				end if
				if request.form("descargo")=checked then
					rs("descargado")=0
				else
					rs("descargado")=1
				end if
			end if
			idhr=rs("idhojaruta")
			rs("fechasalida")=date()
			rs("horasalida")=time()
			if  isnull(rs("observacionesconductor")) or rs("observacionesconductor")="" then
				rs("observacionesconductor")=request.querystring("obser")
			else
				rs("observacionesconductor")=rs("observacionesconductor") & " " & request.querystring("obser")
			end if		
			rs.update
			rs1.CursorType = adOpenKeyset
			rs1.LockType = adLockOptimistic
			rs1.open"select * from hojaruta where remitente='" & rs("destinatario") & "' and numviaje=" & request.querystring("hr"), cn
			if rs1.eof=false then
				if request.form("cargo")=checked then
					rs1("cargado")=0
				else
					rs1("cargado")=1
				end if
				if request.form("descargo")=checked then
					rs1("descargado")=0
				else
					rs1("descargado")=1
				end if
				rs1.update
			end if
			rs1.close 
		else
			idhr=rs("idhojaruta")
			rs("fechasalida")=date()
			rs("horasalida")=time()
			if  isnull(rs("observacionesconductor")) or rs("observacionesconductor")="" then
				rs("observacionesconductor")=request.querystring("obser")
			else
				rs("observacionesconductor")=rs("observacionesconductor") & " " & request.querystring("obser")
			end if		
			rs.update
		end if
	else
		cargaorigen=rs("cargadestino")
		quellevaorigen=rs("quellevadestino")
		remitocarga=ucase(trim(rs("remitocarga")))
		rs7.CursorType = adOpenKeyset
		rs7.LockType = adLockOptimistic
		rs7.open "Select * from hojaruta where numviaje=" & request.querystring("hr") & " and remitente='" & rs("destinatario") & "' and fechallegada is null and fechasalida is null order by idhojaruta",cn
		if rs7.eof=false then
			rs7("cargaorigen")=cargaorigen
			rs7("quellevaorigen")=quellevaorigen
			rs7("cargadestino")=cargaorigen
			rs7("quellevadestino")=quellevaorigen
			rs7("remitocarga")=remitocarga
			rs7.update
		end if
		rs7.close

		idhr=rs("idhojaruta")
		rs("fechasalida")=date()
		rs("horasalida")=time()
		if  isnull(rs("observacionesconductor")) or rs("observacionesconductor")="" then
			rs("observacionesconductor")=request.querystring("obser")
		else
			rs("observacionesconductor")=rs("observacionesconductor") & " " & request.querystring("obser")
		end if
		rs("remitocarga")=remitocarga
		rs("remitocargadestino")=remitocarga	
		rs.update
	end if
end if
rs.close

dni=request.querystring("dni")
clave=request.querystring("clave")
tipo=request.querystring("tipo")
response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)%>