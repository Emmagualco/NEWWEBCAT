<html>

<head>

<script language="JavaScript">
function Abrir_iqnet (pagina) {
var opciones="toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, width=800, height=255, top=85, left=200";
window.open(pagina,"",opciones);
}
</script>




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

  .boton_2{
    text-decoration: none;
    padding: 20px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 200;
    font-size: 35px;
    color: #ffffff;
    background-color: #152b8f;
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
    font-family: arial;
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

input[type=text]:focus{
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
</style>

</head>

<br>
<br>
<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
mes=month(now)
if mes<10 then
	mes="0" & mes
else
	mes=mes
end if
dia=day(now)
if dia<10 then
	dia="0" & dia
else
	dia=dia
end if%>
<%hoy=mes&"/"&dia&"/"&year(now())%>
<%encontro=0
cade="select idconductor from conductores where documento='" & request.form("txtdni") & "' and estado='" & "A" & "' and claveweb='" & request.form("clave") & "'"
set rs6=cn.execute(cade)
if rs6.eof=false then
	idconductor=rs6("idconductor")
	cadena="select numviaje from viajes where idconductor=" & idconductor & " and hojaderuta is not null and fechasalida ='" & hoy & "' and fechallegada is null and web='" & "S" & "' order by numviaje desc"
	set rs1=cn.execute (cadena)
	if rs1.eof=false then
		dni=request.form("txtdni")
		clave=request.form("clave")
		do while rs1.eof=false		
			cade="select * from hojaruta where numviaje=" & rs1("numviaje") & " and tipo='" & request.form("tipo") & "'"
			set rs2=cn.execute(cade)
			if rs2.eof=false then
				encontro=1
				hojaruta=rs2("numviaje")
				numviaje=rs2("numviaje")
				tipo=trim(rs2("tipo"))
				clicli=trim(rs2("cliente"))
				exit do
			end if
			rs1.movenext			
		loop
		rs1.close
		if encontro=0 then%>
			<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
			<head>  
				<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
	  			<title>Sin Hoja de Ruta Por Tipo</title>
  				<link rel="stylesheet" href="estilohr.css">
			</head>
			<body>
				<br>
				<br>
				<br>
				<br>
				<div class="box-green">
					<center><font face="verdana" size="12">No existe Hoja de Ruta generada para este Tipo de Movimiento</font></center>
				</div>
				<br>
				<br>
				<div>	
					<form name="form4" action="loginhr.html">
						<center><input type="submit" class="boton_2" value="Regresar a la Página Principal"></center>
					</form>
				</div> 
			</body>
			</html>	
		<%else
			productos_Sql = "Select * from conductores  where documento='" & request.form("txtdni") & "' and estado='" & "A" & "'"
			set rs = CreateObject("ADODB.Recordset")
			rs.Open productos_Sql, Cn
			if rs.eof=false then%>
				<%idconductor=rs("idconductor")%>
				<center>	
				<table border="0">
					<tbody>
						<tr>
							<td><img src="<%=rs("foto")%>" width="150" height="150"></td>				
							<td><font face="verdana" size="7"><b>Conductor:&nbsp;<%=rs("nombre")%></font></b></td>
							<td><a href="javascript:Abrir_iqnet('consultalc.asp?idconductor=<%=idconductor%>&tipo=<%=tipo%>')"><img src="treslineas.png"width="70" height="50"></a></td>
						</tr>
					</tbody>
				</table>
				</center>
				<br>
				<%cade="select * from viajes where numviaje=" & numviaje
				set rs1=cn.execute (cade)
				if rs1.eof=false then
					hojaruta=trim(rs1("numviaje"))
					cad="select patente,idtipomedio from medios where idmedio="& rs1("idmedio")
					set rs3=cn.execute(cad)
					patente=trim(rs3("patente"))
					idtipo=rs3("idtipomedio")
					cad="select desctipomedio from tipomedio where idtipomedio=" & idtipo
					set rs4=cn.execute(cad)%>
					<center><font face="verdana" size="12" color="green"><b>Dominio:&nbsp;</b></font><font face="verdana" size="7" color="green"><%=patente%>-<%=trim(rs4("desctipomedio"))%></font></center><br>
					<center><font face="verdana" size="12"><b>Hoja de Ruta: &nbsp;<%=hojaruta%></b></center>
					<%cade="select * from hojaruta where numviaje=" & numviaje & " and (horallegada is null or horasalida is null) and tipo='" & request.form("tipo") & "'"
					set rs2=cn.execute(cade)
					if rs2.eof=false then%>
						<center>
						<%tipo=rs2("tipo")%>
						<font face="verdana" size="10"><b>Tipo:&nbsp;<%=tipo%><b></font><br>
						CLIENTE:<%=clicli%>
						
						<table border="0">
				<tr>
					<td colspan="2"><hr></td>
				</tr>
				<tr>
					<td align="right"><font face="verdana" size="12"><b>Origen:</b></font></td>
					<td><font face="verdana" size="12" color="blue"><%=rs2("remitente")%></font></td>
				</tr>
				<tr>
					<td align="right"><font face="verdana" size="12"><b>Destino:</b></font></td>
					<td><font face="verdana" size="12" color="blue"><%=rs2("destinatario")%></font></td>
				</tr>

				<tr>
					<td colspan="2"><hr></td>
				</tr>
			</table>

			<%if not isnull(rs2("ruta")) then%>
				<center><font face="verdana" size="3"> <a href="<%=rs2("ruta")%>">Mostrar Mapa del Recorrido</a></font></center>
			<%end if%>
						<%if isnull(rs2("horallegada")) and isnull(rs2("horasalida")) then%>
							<form name="form5" action="modihr.asp" method="POST">
								<table style="text-align: left; width: 800px;" border="0" cellpadding="2" cellspacing="2">
									<tbody>
										<tr>
											<td><font size="5"><span style="font-family: Verdana;">Para Modificar el Origen o el Destino, haga Click en la imágen-></span></font></td>
											<td><a href="modihr.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>"><img style="width: 100px; height: 100px;" src="editar.jpg"></a></td>
										</tr>
									</tbody>
								</table>					
							</form>
							<%llega="Llego a " & trim(rs2("destinatario"))%>
							<form name="form1" action="registrollegada.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>" method="POST">
								<center><input type="submit" class="boton_2" name="btnllegada" value="<%=llega%>"></center>
							</form>
							<form name="form2" action="loginhr.html">
								<center><input type="submit" class="boton_2" value="Ir a la Página Principal"></center>
							</form>
						<%end if%>
	

			<%if isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then%>
				<font face="verdana" size="6" color="red"><b>Fecha Llegada: &nbsp;<%=rs2("fechallegada")%><br>
				&nbsp;&nbsp;Hora Llegada:&nbsp;<%=rs2("horallegada")%></b></font>
				<hr>
				<table border="0">
					<tr>
						<td>
							<form name ="form4" method="POST" action="cargo.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>">
								<center><input type="submit" class="boton_3" name="cargo" value="CARGO"></center>
							</form>	
						</td>
						<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						<td>
							<form name ="form4" method="POST" action="descargo.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&cargad=<%=rs2("cargaorigen")%>&remitocargad=<%=trim(rs2("remitocargaorigen"))%>&comoviajad=<%=trim(rs2("quellevaorigen"))%>&clave=<%=request.form("clave")%>">
								<center><input type="submit" class="boton_3" name="descargo" value="DESCARGO"></center>
							</form>	
						</td>

					</tr>
				</table>

				<table>
					<tr>
						<td colspan="2"><hr></td>
					</tr>
				</table>

				<%sale="Salgo de " & trim(rs2("destinatario"))%>
				<form name="form1" action="registrosalida.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.form("txtdni")%>&tipo=<%=tipo%>&clave=<%=request.form("clave")%>" method="POST">
					<center><input type="submit" class="boton_2" name="btnsalida" value="<%=sale%>"></center>
				</form>
				<form name="form2" action="loginhr.html">
					<center><input type="submit" class="boton_2" value="Ir a la Página Principal"></center>
				</form>
			<%end if%>











						</center>
					<%else%>
						<html>
							<head></head>
							<body>
								<center>
								<font face="verdana" size="10"><b>Tipo:&nbsp;<%=request.form("tipo")%><b></font>
								<br>
								<%tipo=request.form("tipo")%>
								<font face="verdana" size="10" color="blue"><b><u>GENERACIÓN DE TRAMO</u></b></font><br>
								<%cade="select top(1) destinatario,cliente from hojaruta where numviaje=" & rs1("numviaje") & " and tipo='" & request.form("tipo")& "' order by idhojaruta desc"
								set rs4=cn.execute(cade)
								oriori=rs4("destinatario")
								clicli=rs4("cliente")

								rs4.close%>
								<form name="form1" action="generarhr.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.form("txtdni")%>&hojaruta=<%=hojaruta%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>" method="POST">
									<%cade1="select rscliente from cliente where hrweb='" & "S" & "' order by rscliente"
									set rs4=cn.execute(cade1)%>
									CLIENTE:<select name="cliente" required>
									<%do while rs4.eof=false%>
										<option><%=rs4("rscliente")%></option>
										<%rs4.movenext
									loop
									rs4.close%>
									<option selected><%=clicli%></option>
									</select>
										<table border="0">
										<tr>
											<td style="width:50%" align="right"><font face="verdana" size="12">Origen:</font></td>
											<td style="width:50%">
												<%cad="select rstercero from terceros where estado='" & "A" & "' and tipomov='" & "I" & "' order by rstercero"
												set rs4=cn.execute(cad)%>
												<select name="desde">
												<%do while rs4.eof=false%>
													<option><%=rs4("rstercero")%></option>
													<%rs4.movenext
												loop
												rs4.close%>
												<option selected><%=oriori%></option>
											    </select>
											</td>
										</tr>

										<tr>	
											<td style="width:50%" align="right"><font face="verdana" size="12">Destino:</font></td>
											<td style="width:50%">
												<%cad="select rstercero from terceros where estado='" & "A" & "' and tipomov='" & "I" & "' order by rstercero"
												set rs4=cn.execute(cad)%>
												<select name="hasta">
												<%do while rs4.eof=false%>
													<option><%=rs4("rstercero")%></option>
													<%rs4.movenext
												loop
												rs4.close%>
											    </select>
											</td>
										</tr>

									</table>
									<br>
									<center><input type="submit" class="boton_2" value="    Generar Tramo    "></center>
								</form>	
								<form name="form2" action="loginhr.html">
									<center><input type="submit" class="boton_2" value="Ir a la Página Principal"></center>
								</form>
								</center>
							</body>		
						</html>
					<%end if%>
				<%end if%>
			<%end if%>
		<%end if%>
	<%else%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
		<head>  
			<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
	  		<title>Sin Hoja de Ruta Por Tipo</title>
  			<link rel="stylesheet" href="estilohr.css">
		</head>
		<body background="fondodepantalla.jpg">
			<br>
			<br>
			<br>
			<br>
			<div class="box-green">
				<center><font face="verdana" size="12">No existen Hojas de Ruta generadas para este Conductor</font></center>
			</div>
			<br>
			<br>
			<div>	
				<form name="form4" action="loginhr.html">
					<center><input type="submit" class="boton_2" value="Regresar a la Página Principal"></center>
				</form>
			</div> 
		</body>
		</html>	
	<%end if%>
<%else%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>  
		<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
		<title>Error de Ingreso</title>
			<link rel="stylesheet" href="estilohr.css">
		</head>
		<body>
			<br>
			<br>
			<br>
			<br>
			<div class="box-green">
				<center><font face="verdana" size="12">No existe combinación N° DNI/Password. Por favor, verifique los datos ingresados.</font></center>
			</div>
			<br>
			<br>
			<div>	
				<form name="form4" action="loginhr.html">
					<center><input type="submit" class="boton_2" value="Regresar a la Página Principal"></center>
				</form>
			</div> 
		</body>
	</html>	
<%end if%>
</body>
</html>