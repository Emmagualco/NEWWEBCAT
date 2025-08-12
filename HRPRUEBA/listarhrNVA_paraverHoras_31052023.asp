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
    font-family: arial black;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 200;
    font-size: 60px;
    color: #040504;
    background-color: #24F010;
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
cade="select idconductor,claveweb from conductores where documento='" & request.form("txtdni") & "' and estado='" & "A" & "' and claveweb='" & request.form("clave") & "'"
set rs6=cn.execute(cade)
if rs6.eof=false then
	idconductor=rs6("idconductor")
	clave=rs6("claveweb")
	cadena="select numviaje from viajes where idconductor=" & idconductor & " and hojaderuta is not null and fechasalida ='" & hoy & "' and fechallegada is null and web='" & "S" & "' order by numviaje desc"
	set rs1=cn.execute (cadena)
	if rs1.eof=false then
		horaserver=rs("horasalida")
		hora=hour(horaserver)
		if hora<10 then
			hora="0" & hour(horaserver)
		end if
		minu=minute(horaserver)
		if minu<10 then
			minu="0" & minute(horaserver)
		end if
		cadhora=hora & ":" & minu
		ampm=(right(horaserver, 2))

		horaserver=time()
		hora=hour(horaserver)
		if hora<10 then
			hora="0" & hour(horaserver)
		end if
		minu=minute(horaserver)
		if minu<10 then
			minu="0" & minute(horaserver)
		end if
		cadhorahoy=hora & ":" & minu

		dni=request.querystring("dni")
		clave=request.querystring("clave")
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
	  			<title>Hoja de Ruta Por Tipo</title>
			</head>
			<body bgcolor="#9091A0">
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
					<form name="form4" action="loginhrnva.html">
						<center><input type="submit" class="boton_2" value="Regresar a la Página Principal"></center>
					</form>
				</div> 
			</body>
			</html>	
		<%else%>
			<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
			<head>  
				<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
	  			<title>Hoja de Ruta Por Tipo</title>
			</head>
			<body bgcolor="#9091A0">
			<%productos_Sql = "Select * from conductores  where documento='" & request.form("txtdni") & "' and claveweb='" & request.form("clave") & "' and estado='" & "A" & "'"
			set rs = CreateObject("ADODB.Recordset")
			rs.Open productos_Sql, Cn
			if rs.eof=false then%>
				<%idconductor=rs("idconductor")%>
				<%clave=rs("claveweb")%>
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
					<%cade="select * from hojaruta where numviaje=" & numviaje & " and (horallegada is null or horasalida is null) and tipo='" & request.form("tipo") & "'"
					set rs2=cn.execute(cade)
					if rs2.eof=false then%>

						<%tipo=trim(rs2("tipo"))%>						
						<%if not isnull(rs2("ruta")) then%>
							<center><font face="verdana" size="3"> <a href="<%=rs2("ruta")%>">Mostrar Mapa del Recorrido</a></font></center>
						<%end if%>
						<%if isnull(rs2("horallegada")) and isnull(rs2("horasalida")) then%>

							<table width="100%" bgcolor="white" border="0">
								<tr>
									<td width="20%"><img src="LogoCATNUEVO.gif"></td>
									<td width="60%" align="right">
										<form name="form5" action="modihrnva.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>" method="POST">
											<font size="5" color="black" face="verdana">Modificar Destino-></font>
									<td>
									<td>
											<a href="modihrnva.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>"><img style="width: 80px; height: 80px;" src="editar.jpg"></a>
										</form>
									</td>
									<td width="20%" align="center"><a href="javascript:Abrir_iqnet('consultalc.asp?idconductor=<%=idconductor%>&clave=<%=clave%>&tipo=<%=tipo%>')"><img src="treslineas.png"width="70" height="50"></a></td>
								<tr>
							</table>
							
							<center>	
							<br>
							<br>
							<font face="verdana" size="20" color="white"><b>EN CAMINO A <%=trim(rs2("destinatario"))%></b></font>
							<br>
							<br>
							<br>
							<br>
							<br>
							<center>
							<img src="camioncitocargado.jpg">
							</center>
							<br>
							<br>
							<font face="verdana" size="20" color="white"><b>Contenido actual</b><br>
							<%if trim(rs2("quellevadestino"))="VACÍO"then%>
								<%=(rs2("quellevadestino"))%>
							<%else%>
								<%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%>
							<%end if%>
							<form name="form1" action="registrollegadaNVA.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>" method="POST">
								<center><input type="submit" class="boton_2" name="btnllegada" value="                  LLEGADA                  "></center>
							</form>
							<br>
							<br>
							<br>
							<form name="form2" action="loginhrnva.html">
								<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
							</form>
						<%end if%>
	

			<%if isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then%>

				<table width="100%" bgcolor="white" border="0">
					<tr>
						<td width="20%"><img src="LogoCATNUEVO.gif"></td>
						<td width="60%">&nbsp;</td>
						<td width="20%" align="center"><a href="javascript:Abrir_iqnet('consultalc.asp?idconductor=<%=idconductor%>&tipo=<%=tipo%>')"><img src="treslineas.png"width="70" height="50"></a></td>
					<tr>
				</table>

							<br>
							<br>
							<center>	
							<font face="verdana" size="20" color="white"><b>ESTOY EN&nbsp;<%=trim(rs2("destinatario"))%></b></font>
							<br>
							<br>
							<br>
							<br>
							<br>
							<center>
							<img src="camioncitocargado.jpg">
							</center>
							<br>
							<br>
							<center><font face="verdana" size="20" color="white"><b>Contenido actual</b><br>
							<%if trim(rs2("quellevadestino"))="VACÍO" then%>
								<%=(rs2("quellevadestino"))%>
							<%else%>
								<%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%>
							<%end if%>
							<%if trim(rs2("quellevadestino"))="VACÍO" then%>
								<center>
								<table border="0">
									<tr>
										<td>
											<form name ="form4" method="POST" action="cargoNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.form("txtdni")%>&clave=<%=request.form("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>">
												<center><input type="submit" class="boton_3" name="cargo" value="CARGO"></center>
											</form>	
										</td>
									</tr>
								</table>
								</center>
							<%else%>
								<table border="0">
									<tr>
										<td>
											<form name ="form4" method="POST" action="cargoNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.form("txtdni")%>&clave=<%=request.form("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>">
												<center><input type="submit" class="boton_3" name="cargo" value="CARGO"></center>
											</form>	
										</td>
										<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
										<td>
											<form name ="form4" method="POST" action="descargoNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.form("txtdni")%>&clave=<%=request.form("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&cargad=<%=rs2("cargaorigen")%>&remitocargad=<%=trim(rs2("remitocargaorigen"))%>&comoviajad=<%=trim(rs2("quellevaorigen"))%>">
												<center><input type="submit" class="boton_3" name="descargo" value="DESCARGO"></center>
											</form>	
										</td>
									</tr>
								</table>
							<%end if%>
							<form name="form1" action="registrosalidaNVA.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>" method="POST">
								<center><font face="verdana" color="white">Observaciones:</font><input type="text" name="obser" maxlength="500"></center><br>
								<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
							</form>
							<br>
							<br>
							<form name="form2" action="loginhrnva.html" method="POST">
								<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
							</form>



			<%end if%>











						</center>
					<%else%>
						<html>
							<head></head>
							<body bgcolor="#9091A0">
							<center>
							<table width="100%" bgcolor="white" border="0">
								<tr>
									<td width="20%"><img src="LogoCATNUEVO.gif"></td>
									<td width="60%">&nbsp;</td>
									<td width="20%" align="center"><a href="javascript:Abrir_iqnet('consultalc.asp?idconductor=<%=idconductor%>&tipo=<%=tipo%>')"><img src="treslineas.png"width="70" height="50"></a></td>
								<tr>
							</table>
							<br>
							<br>
							<%cade="select * from hojaruta where numviaje="&numviaje & " order by idhojaruta desc"
							set rs4=cn.execute(cade)%>
							<font face="verdana" size="20" color="white"><b>ÚLTIMO DESTINO:&nbsp;<%=rs4("destinatario")%></b></font><br><br>
							<font face="verdana" size="20" color="white"><b>ESPERANDO UN NUEVO DESTINO.....</b></font>
							<br>
							<br>
							<center>
							<img src="camioncitocargado.jpg">
							<br>
							<br>
							<font face="verdana" size="20" color="white"><b>Contenido actual</b><br>
							<%if trim(rs4("quellevadestino"))="VACÍO" then%>
								<%=(rs4("quellevadestino"))%>
							<%else%>
								<%=rs4("cargadestino") & " " & trim(rs4("quellevadestino"))%>
							<%end if%>
							<br>
							<form name="form1" action="listarhr1NVA.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>" method="POST">
								<center><input type="submit" class="boton_2" name="btnllegada" value="                  REFRESCAR                 "></center>
							</form>
							<br>
							<%if not isnull(rs4("fechavalidacion")) then%>
								<form name="form2" action="validacerrarhrnva.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>" method="POST">
									<center><a href="validacerrarhrnva.asp?hr=<%=numviaje%>&dni=<%=request.form("txtdni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.form("clave")%>"><img style="width: 150px; height: 150px;" src="cerrarhr.jpg"></a></center>
								</form>
							<%end if%>
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
		<body bgcolor="#9091A0">
			<br>
			<br>
			<br>
			<br>
			<div class="box-green">
				<center><font face="verdana" size="12">No existen Hojas de Ruta generadas para este Conductor/Tipo de Movimiento</font></center>
			</div>
			<br>
			<br>
			<div>	
				<form name="form4" action="loginhrnva.html">
					<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
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
				<form name="form4" action="loginhrnva.html">
					<center><input type="submit" class="boton_2" value="Regresar a la Página Principal"></center>
				</form>
			</div> 
		</body>
	</html>	
<%end if%>
</body>
</html>