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
cad1="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=BFGNXT"
cn1.open(cad1)
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs1.CursorType = adOpenKeyset
rs1.LockType = adLockOptimistic
rs.open "Select * from hojaruta where numviaje=" & request.querystring("hr") & " and fechallegada is null order by idhojaruta",cn
if rs.eof=false then
	if trim(rs("cliente"))="RENAULT ARGENTINA S.A.  PRODUCCIÓN" then
		if trim(rs("remitente"))<>"PLAYA CAT" then
			cade="select horasalida from hojaruta where destinatario='" & trim(rs("remitente")) & "' and numviaje=" & request.querystring("hr")
			set rs7=cn.execute(cade)
			salio=left(rs7("horasalida"),5)
			rs7.close
			ahora = Left(Time(), 5)
			viajo=datediff("n",salio,ahora)
			if viajo > 31 or viajo < 10 then%>
				<html>
					<head>
					<style type="text/css">

						body{
							background: url(imagenruta.png) "CCC no-repeat;
							background-size: cover;
							background-position:center center;
							background-attachment: fixed;

							-webkit-background-size: cover;
							-moz-background-size: cover;
							-o-background-zixe: cover;
						}

						.boton_1{
    							text-decoration: none;
							padding: 20px;
    							font-family: arial black;
    							text-transform: uppercase;
    							padding-left: 10px;
    							padding-right: 10px;
    							font-weight: 200;
    							font-size: 80px;
   	 						color: #000000;
    							background-color: #29FB04;
    							border-radius: 15px;
    							border: 2px solid #0016b0;
 	 					}
  						.boton_1:hover{
    							color: #9b0e0e;
    							background-color: #bc6355;
    							text-decoration: none;
  						}


  						.boton_2{
    							text-decoration: none;
    							padding: 20px;
    							font-family: arial black;
    							text-transform: uppercase;
    							padding-left: 10px;
    							padding-right: 10px;
    							font-weight: 200;
    							font-size: 80px;
    							color: #FFFFFF;
    							background-color: #0A20F9;
    							border-radius: 15px;
    							border: 2px solid #0016b0;
  						}
  						.boton_2:hover{
    							color: #9b0e0e;
    							background-color: #bc6355;
    							text-decoration: none;
  						}

  						.boton_3{
    							text-decoration: none;
    							padding: 20px;
    							font-family: arial black;
    							text-transform: uppercase;
    							padding-left: 10px;
    							padding-right: 10px;
    							font-weight: 200;
    							font-size: 35px;
    							color: #1C2833;
    							background-color:#D4EFDF;
    							border-radius: 15px;
    							border: 2px solid #27AE60;
  						}
  						.boton_3:hover{
    							color: #9b0e0e;
    							background-color:  #AED581;
    							text-decoration: none;
  						}

  						.boton_4{
    							text-decoration: none;
    							padding: 10px;
    							font-family: verdana black;
    							text-transform: uppercase;
    							padding-left: 10px;
    							padding-right: 10px;
    							font-weight: 100;
   	 						font-size: 14px;
    							color: #ffffff;
    							background-color:#FD0316;
    							border-radius: 15px;
    							border: 2px solid #0016b0;
  						}
  						.boton_4:hover{
    							color: #9b0e0e;
    							background-color:  #AED581;
    							text-decoration: none;
  						}

  						input[type=date] {
							font-size: 35px;
							width: 240px;
						}

						.caja {
   							margin:20px auto 40px auto;	
   							border:1px solid #d9d9d9;
   							height:30px;
   							overflow: hidden;
   							width: 230px;
   							position:relative;
						}
	
						select {
   							background-color:#FFCE33;
   							border: none;
   							font-size: 40px;
   							font-face: "Century Gothic"
   							height: 30px;
   							padding: 0px;
   							width: 500px;
   							size: 10px;
						}
						select:focus{ outline: none;}

						.caja::after{
							content:"\025be";
							display:table-cell;
							padding-top:7px;
							text-align:center;
							width:30px;
							height:30px;
							background-color:#d9d9d9;
							position:absolute;
							top:0;
							right:0px;	
							pointer-events: none;
						}
						input[type=radio] { border: 0px; width: 5%; height: 3em; }

  						.boton_personalizado{
    							text-decoration: none;
    							padding: 10px;
    							font-weight: 1000;
    							font-size: 40px;
    							color: #ffffff;
    							background-color: #1883ba;
    							border-radius: 6px;
	    						border: 2px solid #0016b0;
  						}
						img{
							border-radius: 50%;
						}

						input[type=text]{
							width:50%;
							border:2px solid #aaa;
							border-radius:4px;
							margin:8px 0;
							outline:none;
							padding:8px;
							box-sizing:border-box;
							transition:3s;
							font-size: 33px;
						}

						input[type=text]:focus{
							border-color:dodgerBlue;
							box-shadow:0 0 8px 0 dodgerblue;
						}


						input[type=time]{
							width:200px;
							border:2px solid #aaa;
							border-radius:4px;
							margin:8px 0;
							outline:none;
							padding:8px;
							box-sizing:border-box;
							transition:3s;
							font-size: 33px;
						}

						input[type=time]:focus{
							border-color:dodgerBlue;
							box-shadow:0 0 8px 0 dodgerblue;
						}

						input[type=number]{
							width:30%;
							border:2px solid #aaa;
							border-radius:4px;
							margin:8px 0;
							outline:none;
							padding:8px;
							box-sizing:border-box;
							transition:3s;
							font-size: 33px;
						}

						input[type=number]:focus{
							border-color:dodgerBlue;
							box-shadow:0 0 8px 0 dodgerblue;
						}

						#visor_imagenes {
  							text-align: center;
						}

						.blanco-violeta {
  							color: #FFFFFF;
  							background-color: #9B59B6;
						}

						.blanco-naranja {
  							color: #FFFFFF;
  							background-color: #E67E22;
						}


						input[type=checkbox] {
  							transform: scale(4);
						}
					</style>

					</head>
					<body bgcolor="#9091A0">
						<form name="form1" action="registrollegadaNVAMR1.asp?hr=<%=request.querystring("hr")%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">								
							<center><font face="verdana" color="white" size="8">Ha tardado&nbsp;<%=viajo%> minutos para llegar a&nbsp;<%=trim(rs("destinatario"))%>, por favor, complete los siguientes datos:</font></center>
							<br>
	 						<center><font face="verdana" color="black" size="2">Hora Salida Anterior:&nbsp<input type="time" id="apps" name="apps" max="<%=left(time(),5)%>" value="<%=salio%>" required></font></center>		
							<center><font face="verdana" color="black" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hora Llegada:&nbsp;<input type="time" id="appll" name="appll" max="<%=left(time(),5)%>"  min="<%=salio%>" required></font></center>		
							<center><font face="verdana" color="black" size="12">Observaciones:</font><input type="text" name="obser" maxlength="500" size="80" required></center>
							<br>
							<table style="text-align: left; width: 100%;" border="0">
								<tr>
									<td align="right" VALIGN="middle">
										<input type="submit" class="boton_1" value="REGISTRAR">&nbsp;&nbsp;&nbsp;&nbsp;
									</td>									
						</form>
						<td VALIGN="bottom"><br>
							<form name="form2" action="listarhr1NVA.asp?hr=<%=request.querystring("hr")%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>" method="POST">
								&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" class="boton_2" value="CANCELAR">
							</form>
						</td>
								</tr>
						</table>		
					</body>
				</html>
			<%else
				idhr=rs("idhojaruta")
				rs("fechallegada")=date()
				rs("horallegada")=time()
				if request.form("obs")<>"" then
					rs("observacionesconductor")=ucase(trim(request.form("obser")))
				end if
				rs.update
				dni=request.querystring("dni")
				clave=request.querystring("clave")
				tipo=request.querystring("tipo")
				response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
			end if
			rs6.CursorType = adOpenKeyset
			rs6.LockType = adLockOptimistic
			rs6.open"select * from hojaruta where numviaje=" & request.querystring("hr") & " and remitente='" & trim(rs("destinatario")) & "'",cn
			if rs6.eof=false then
				if request.form("cargo")=checked then
					rs6("cargado")=0
				else
					rs6("cargado")=1
				end if
				if request.form("descargo")=checked then
					rs6("descargado")=0
				else
					rs6("descargado")=1
				end if
				rs6.update
			end if
			rs6.close
			rs.update
		else
			idhr=rs("idhojaruta")
			rs("fechallegada")=date()
			rs("horallegada")=time()
			if request.form("obs")<>"" then
				rs("observacionesconductor")=ucase(trim(request.form("obser")))
			end if
			rs.update

			rs4.CursorType = adOpenKeyset
			rs4.LockType = adLockOptimistic
			rs4.open"select * from viajes where numviaje=" & request.querystring("hr"),cn
			if rs4("etapa")="G" then
				rs4("etapa")="A"
			end if
			rs4.update
			rs4.close

			dni=request.querystring("dni")
			clave=request.querystring("clave")
			tipo=request.querystring("tipo")
			response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
		end if
	else
		rs1.open"select * from viajes where numviaje=" & request.querystring("hr"),cn
		if not isnull(rs1("shipment")) Or Not IsNull(RS1("remitoiveco")) then
			idmedio=rs1("idmedio")
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

			idhr=rs("idhojaruta")
			rs("fechallegada")=date()
			rs("horallegada")=time()
			if request.form("obs")<>"" then
				rs("observacionesconductor")=ucase(trim(request.form("obser")))
			end if
			rs.update
			if rs1("etapa")="G" then
				rs1("etapa")="A"
				rs1.update
				cadena="select cliente from hojaruta where numviaje=" & request.querystring("hr")
				set rs4=cn.execute(cadena)
				ncliente=trim(rs4("cliente"))
				rs4.close 
				cadena="select clientebfg from cliente where rscliente='" & ncliente & "'"
				set rs4=cn.execute(cadena)
				idcli=rs4("clientebfg")
				rs4.close
				remito=trim(rs1("remitoiveco"))
				if not isnull(rs1("shipment"))then
					rs5.CursorType = adOpenKeyset
					rs5.LockType = adLockOptimistic
					rs5.open"select * from logitex where patenteid='" & cadena11 & "' and logitexnroordentrabajo is null",cn1
					if rs5.eof=false then
						rs5.Fields("logitexfechainicio") = Date() & " " & Time()
						rs5.Fields("logitexfechafin") = NULL
						rs5.Fields("logitexenviar") = 1
						rs5.Fields("clienteid")=idcli
						rs5.Update
					else
						rs5.addnew
						rs5("patenteid")=cadena11
						rs5.Fields("logitexfechainicio") = Date() & " " & Time()
						rs5.Fields("logitexenviar") = 1
						rs5.Fields("clienteid")=idcli
						rs5.Update
					end if
					rs5.Close	
				end if
				if not isnull(rs1("remitoiveco"))then
					rs5.CursorType = adOpenKeyset
					rs5.LockType = adLockOptimistic
					rs5.open"logitex",cn1
					rs5.addnew
					rs5("patenteid")=cadena11
					rs5.Fields("logitexfechainicio") = Date() & " " & Time()
					rs5.Fields("logitexenviar") = 1
					rs5.Fields("clienteid")=idcli
					rs5.Fields("logitexnroordentrabajo")=remito
					rs5.Update
					rs5.Close	
				end if
				dni=request.querystring("dni")
				clave=request.querystring("clave")
				tipo=request.querystring("tipo")
				response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
			else
				dni=request.querystring("dni")
				clave=request.querystring("clave")
				tipo=request.querystring("tipo")
				response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
			end if
		else
			idhr=rs("idhojaruta")
			rs("fechallegada")=date()
			rs("horallegada")=time()
			if request.form("obs")<>"" then
				rs("observacionesconductor")=ucase(trim(request.form("obser")))
			end if
			rs.update

			rs7.CursorType = adOpenKeyset
			rs7.LockType = adLockOptimistic
			rs7.open"select * from viajes where numviaje=" & request.querystring("hr"),cn
			if rs7("etapa")="G" then
				rs7("etapa")="A"
				rs7.update
			end if
			rs7.close
			dni=request.querystring("dni")
			clave=request.querystring("clave")
			tipo=request.querystring("tipo")
			response.redirect("listarhr1nva.asp?dni="&dni & " &clave=" & clave & " &tipo=" & tipo)
		end if
		rs1.close
	end if
end if
rs.close%>