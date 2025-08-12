<%set cn=server.createobject("ADODB.Connection")
set cn1=server.createobject("ADODB.Connection")
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
cad1="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=BFGNXT"
cn1.open(cad1)
cadena="select fechallegada,horallegada from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta"
set rs1=cn.execute(cadena)
fechallegada=rs1("fechallegada")
horallegada=trim(rs1("horallegada"))
rs1.close
cadena="select fechasalida,horasalida from hojarutamr where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
set rs2=cn.execute(cadena)
if rs2.eof=false then
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
	idmedio=rs("idmedio")
	cad2="select patente from medios where idmedio=" & idmedio
	set rs2=cn.execute(cad2)
	cadena = rs2.Fields("patente")
	cadena1 = Replace(cadena, " ", "")
	cadena2 = Replace(cadena1, "*", "")
	cadena3 = Replace(cadena2, "-", "")
	cadena4 = Replace(cadena3, "_", "")
	cadena5 = Replace(cadena4, "/", "")
	cadena6 = Replace(cadena5, "%", "")
	cadena7 = Replace(cadena6, "#", "")
	cadena8 = Replace(cadena7, "$", "")
	cadena9 = Replace(cadena8, "(", "")
	cadena10 = Replace(cadena9, ")", "")
	cadena11 = Replace(cadena10, "+", "")
	rs2.close

	rs("fechallegada")=fechasalida
	rs("horallegada")=horasalida
	rs("fechasalida")=fechallegada
	rs("horasalida")=horallegada
	rs("etapa")="C"
	rs.update



	rs4.CursorType = adOpenKeyset
	rs4.LockType = adLockOptimistic
	rs4.open "select fechasalida,horasalida from hojarutamr where numviaje=" & request.querystring("hr") & " order by idhojaruta desc",cn
	rs4("fechasalida")=fechasalida
	rs4("horasalida")=horasalida
	rs4.update
	rs4.close




	if not isnull(rs("shipment"))  then
		rs1.CursorType = adOpenKeyset
		rs1.LockType = adLockOptimistic
		rs1.open"select * from logitex where patenteid='" & cadena11 & "' and logitexnroordentrabajo is null",cn1
		if rs1.eof=false then
			rs1.Fields("logitexfechafin") = Date() & " " & Time()
			rs1.Fields("logitexenviar") = 0
			rs1("clienteid")=0
			rs1.Update
		end if
		rs1.Close	
	end if

	if not isnull(rs("remitoiveco")) then
		rs1.CursorType = adOpenKeyset
		rs1.LockType = adLockOptimistic
		rs1.open"select * from logitex where patenteid='" & cadena11 & "' and logitexnroordentrabajo='" & trim(rs("remitoiveco")) & "'",cn1
		if rs1.eof=false then
			rs1.Fields("logitexfechafin") = Date() & " " & Time()
			rs1.Fields("logitexenviar") = 0
			rs1.Update
		end if
		rs1.Close	
	end if
	
	rs.close

	rs.open"select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
	rs("horasalida")=horasalida
	rs("fechasalida")=fechasalida
	rs.update
	rs.close
	response.redirect("loginhrnva.html")
else
	caden="select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
	set rs6=cn.execute(caden)
	if rs6.eof=false then
		if not isnull(rs6("fechasalida")) then
			fechasalida=rs6("fechasalida")
		else
			fechasalida=date()
		end if
		if not isnull(rs6("horasalida")) then
			horasalida=trim(rs6("horasalida"))
		else
			horasalida=time()
		end if
		rs.CursorType = adOpenKeyset
		rs.LockType = adLockOptimistic
		rs.open "Select * from viajes where numviaje=" & request.querystring("hr"),cn
		idmedio=rs("idmedio")
		cad2="select patente from medios where idmedio=" & idmedio
		set rs2=cn.execute(cad2)
		cadena = rs2.Fields("patente")
		cadena1 = Replace(cadena, " ", "")
		cadena2 = Replace(cadena1, "*", "")
		cadena3 = Replace(cadena2, "-", "")
		cadena4 = Replace(cadena3, "_", "")
		cadena5 = Replace(cadena4, "/", "")
		cadena6 = Replace(cadena5, "%", "")
		cadena7 = Replace(cadena6, "#", "")
		cadena8 = Replace(cadena7, "$", "")
		cadena9 = Replace(cadena8, "(", "")
		cadena10 = Replace(cadena9, ")", "")
		cadena11 = Replace(cadena10, "+", "")
		rs2.close
	
		rs("fechallegada")=fechasalida
		rs("horallegada")=horasalida
		rs("fechasalida")=fechallegada
		rs("horasalida")=horallegada
		rs("etapa")="C"
		rs.update

		if not isnull(rs("shipment"))  then
			rs1.CursorType = adOpenKeyset
			rs1.LockType = adLockOptimistic
			rs1.open"select * from logitex where patenteid='" & cadena11 & "' and logitexnroordentrabajo is null",cn1
			if rs1.eof=false then
				rs1.Fields("logitexfechafin") = Date() & " " & Time()
				rs1.Fields("logitexenviar") = 0
				rs1.Update
			end if
			rs1.Close	
		end if

		if not isnull(rs("remitoiveco")) then
			rs1.CursorType = adOpenKeyset
			rs1.LockType = adLockOptimistic
			rs1.open"select * from logitex where patenteid='" & cadena11 & "' and logitexnroordentrabajo='" & trim(rs("remitoiveco")) & "'",cn1
			if rs1.eof=false then
				rs1.Fields("logitexfechafin") = Date() & " " & Time()
				rs1.Fields("logitexenviar") = 0
				rs1.Update
			end if
			rs1.Close	
		end if

		rs.close

		rs.open"select fechasalida,horasalida from hojaruta where numviaje=" & request.querystring("hr") & " order by idhojaruta desc"
		rs("horasalida")=horasalida
		rs("fechasalida")=fechasalida
		rs.update
		rs.close
		response.redirect("loginhrnva.html")
	end if
end if%>