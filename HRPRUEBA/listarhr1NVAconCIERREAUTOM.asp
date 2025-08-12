<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>  
	<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
  	<title>M&oacute;dulo de Hoja de Ruta</title>

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
width:600px;
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
}px

.blanco-naranja {
  color: #FFFFFF;
  background-color: #E67E22;
}


input[type=checkbox] {
  transform: scale(4);
}
</style>

</head>

<%
dim idhr
dim destino


set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
set rs7=server.createobject("ADODB.Recordset")
set rs8=server.createobject("ADODB.Recordset")
set rs9=server.createobject("ADODB.Recordset")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
mes=month(now())%>
<%if mes<10 then
	mes="0" & month(now())
end if%>	
<%dia=day(now())%>
<%if dia <10 then
	dia="0" & day (now())
end if%>					
<%anio=year(now())%>
<%cadenadesde=anio & "-" & mes & "-" & dia%>
<%desdecadena=mes & "/" & dia & "/" & anio%>		
<%if request.querystring("tipo")="MILK RUN" then
	movtipo="M"	
end if%>
<%if request.querystring("tipo")="MOVIMIENTOS INTERNOS" then
	movtipo="I"
end if%>
<%if request.querystring("tipo")="SPOT" then
	movtipo="S"
end if%>

<%encontro=0
encontropat=0
cade="select * from conductores where documento='" & request.querystring("dni") & "' and estado='" & "A" & "' and claveweb='" & request.querystring("clave") & "'"
set rs6=cn.execute(cade)
if rs6.eof=true then%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>  
		<meta content="text/html; charset=utf-8" http-equiv="content-type">
		<title>Error de Ingreso</title>
			<link rel="stylesheet" href="estilohr.css">
		</head>
		<body bgcolor="#9091A0">
			<br>
			<br>
			<br>
			<br>
			<div class="box-green">
				<center><font face="verdana" size="12">No existe combinaci&oacute;n N&deg; DNI/Password. Por favor, verifique los datos ingresados.</font></center>
			</div>
			<br>
			<br>
			<div>
				<form name="form2" action="loginhrnva.html">
					<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
				</form>
			</div> 
		</body>
	</html>	
<%else%>
	<%if movtipo="I" then
	do while rs6.eof=false
        	caden="select * from viajes where idconductor=" & RS6("idconductor") & " and etapa='" & "A" & "' and fechasalida<='" & desdecadena & "' and fechallegada is null and web='" & "S" & "' and tipomov='" & "I" & "'"
		set rs=cn.execute(caden)
        	If RS.EOF = False Then
            		numviaje = RS("numviaje")
			idconductor=rs6("idconductor")
			clave=rs6("claveweb")
            		encontro = 1
        	End If
        	RS.Close
        	RS6.MoveNext
    	Loop
	RS6.Close
	if encontro=0 then
		cade="select * from conductores where documento='" & request.querystring("dni") & "' and estado='" & "A" & "' and claveweb='" & request.querystring("clave") & "'"
		set rs6=cn.execute(cade)	
		If RS6.EOF = False Then%>
			<%Do While RS6.EOF = false
        			caden= "select * from viajes where idconductor=" & RS6("idconductor") & " and etapa='" & "G" & "' and fechasalida<='" & desdecadena & "' and fechallegada is null and web='" & "S" & "'"
				set rs=cn.execute(caden)
        			If RS.EOF = False Then
            				numviaje = RS("numviaje")
	    				idconductor=rs6("idconductor")
		    			clave=rs6("claveweb")
        	    			encontro = 1
            				cad="select * from viajes where idmedio=" & RS("idmedio") & " and etapa='" & "A" & "'"
					set rs7=cn.execute(cad)
            				If RS7.EOF = False Then
                				encontro = 0
						encontropat=1
                				RS7.Close
                				Exit Do
            				Else
                				RS7.Close
            				End If
        			End If
        			RS.Close
        			RS6.MoveNext
    			Loop
		end if
		RS6.Close
	end if%>
		<%If encontropat = 1 Then%>
			<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">			
			<html>
				<head>  
					<meta content="text/html; charset=utf-8" http-equiv="content-type">
		  			<title>Hoja de Ruta Por Tipo</title>
					<link rel="stylesheet" href="estilohr.css">
				</head>

				<body bgcolor="#9091A0">
					<br>
					<center>
						<form name="form10" action="consultainformeshrweb.asp?tipo=<%=trim(request.querystring("tipo"))%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
							<td width="50%" align="right"><font face="verdana" color="white" size="5"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
						</form>
					</center>
					<br>
					<br>
					<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
					<br>
					<br>
					<div class="box-green">
						<center><font face="verdana" size="12">Su Patente tiene una Hoja de Ruta ABIERTA para este Tipo de Movimiento</font></center>
					</div>
					<br>
					<br>
					<div>
						<form name="form2" action="loginhrnva.html">
							<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
						</form>
					</div> 
				</body>
			</html>	
		<%else%>
		<%If encontro = 0 Then%>
			<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
			<html>
			<head>  
				<meta content="text/html; charset=utf-8" http-equiv="content-type">
	  			<title>Hoja de Ruta Por Tipo</title>
				<link rel="stylesheet" href="estilohr.css">
			</head>
			<body bgcolor="#9091A0">
				<br>
				<br>
				<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
				<br>
				<div class="box-green">
					<center><font face="verdana" size="12">No existe Hoja de Ruta generada para este Tipo de Movimiento</font></center>
				</div>
				<br>
				<br>
				<center>
				<form name="form10" action="consultainformeshrweb.asp?tipo=<%=trim(tipo)%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
					<td width="50%" align="right"><font face="verdana" color="white" size="5"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
				</form>
				</center>
				<br>
				<br>
				<div>
					<form name="form2" action="loginhrnva.html">
						<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
					</form>
				</div> 
			</body>
		</html>	
	<%else%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
		<html>
		<head>  
			<meta content="text/html; charset=utf-8" http-equiv="content-type">
	  		<title>Hoja de Ruta Por Tipo</title>
		</head>
		<%cade="select * from viajes where numviaje=" & numviaje
		set rs1=cn.execute (cade)
		if rs1.eof=false then
			hojaruta=trim(rs1("numviaje"))
			cad="select patente,idtipomedio from medios where idmedio="& rs1("idmedio")
			set rs3=cn.execute(cad)
			patente=trim(rs3("patente"))
			idtipo=rs3("idtipomedio")
			cad="select desctipomedio from tipomedio where idtipomedio=" & idtipo
			set rs4=cn.execute(cad)
			cade="select * from hojaruta where numviaje=" & numviaje & " and tipo='" & request.querystring("tipo") & "' order by idhojaruta desc"
			set rs2=cn.execute(cade)
			if rs2.eof=false then
				if isnull(rs2("horallegada")) and isnull(rs2("horasalida")) then%>
					<body bgcolor="#0A20F9">
				<%end if
				if isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then%>
					<body bgcolor="#29FB04">
				<%end if
				if not isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then%>
					<body background="fondo_ondulado.jpg">
				<%end if
				tipo=trim(rs2("tipo"))						
				if not isnull(rs2("ruta")) then%>
					<center><font face="verdana" size="3"> <a href="<%=rs2("ruta")%>">Mostrar Mapa del Recorrido</a></font></center>
				<%end if
				if isnull(rs2("horallegada")) and isnull(rs2("horasalida")) then%>
					<table width="100%" bgcolor="white" border="0">
						<tr>
							<td width="15%"><img src="LogoCATNUEVO.gif"></td>
							<td width="10%"><a href="cambiodeclave.asp?dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img src="cambioclave.jpg"width="70" height="70"></a></td>
							<td width="15%" align="right"><font size="5" color="black" face="verdana">Modificar Destino-></font></td>
							<td>
								<form name="form5" action="modihrnva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>" method="POST">
									<a href="modihrnva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img style="width: 80px; height: 80px;" src="editar.jpg"></a>
								</form>
							</td>
							<form name="form10" action="consultainformeshrweb.asp?idhr=<%=rs2("idhojaruta")%>&tipo=<%=trim(tipo)%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
								<td width="50%" align="right"><font face="verdana"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
							</form>
						</tr>
					</table>							
					<center>	
					<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
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
					<%if trim(rs2("quellevadestino"))="VACIO"then%>
						<%=(rs2("quellevadestino"))%>
					<%else%>
						<%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%>
					<%end if%>
					<form name="form1" action="registrollegadaNVA.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>" method="POST">
						<center><input type="submit" class="boton_1" name="btnllegada" value="                  LLEGADA                  "></center>
					</form>
					<br>
					<br>
					<br>
					<form name="form2" action="loginhrnva.html">
						<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
					</form>
				<%end if
				if isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then%>
					<table width="100%" bgcolor="white" border="0">
						<tr>
							<td width="20%"><img src="LogoCATNUEVO.gif"></td>
							<td width="20%"><a href="cambiodeclave.asp?dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img src="cambioclave.jpg"width="70" height="70"></a></td>
							<form name="form10" action="consultainformeshrweb.asp?idhr=<%=rs2("idhojaruta")%>&tipo=<%=trim(tipo)%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
								<td width="60%" align="right"><font face="verdana"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
							</form>
						</tr>
					</table>
					<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
					<br>
					<center>	
					<font face="verdana" size="20" color="black"><b>ESTOY EN&nbsp;<%=trim(rs2("destinatario"))%></b></font>
					<br>
					<br>
					<br>
					<br>
					<br>
					<img src="mulita.jpg" width="250" height="200">
					</center>
					<br>
					<br>
					<center><font face="verdana" size="20" color="black"><b>Contenido actual</b><br>
					<%if trim(rs2("quellevadestino"))="VACIO" then%>
						<%=(rs2("quellevadestino"))%>
					<%else%>
						<%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%>
					<%end if%>
					<%if trim(rs2("quellevadestino"))="VACIO" then%>
						<center>
						<table border="0">
							<tr>
								<td>
									<form name ="form4" method="POST" action="cargoNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>">
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
									<form name ="form4" method="POST" action="cargoNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>">
										<center><input type="submit" class="boton_3" name="cargo" value="CARGO"></center>
									</form>	
								</td>
								<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
								<td>
									<form name ="form4" method="POST" action="descargoNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=trim(tipo)%>&cargad=<%=rs2("cargaorigen")%>&remitocargad=<%=trim(rs2("remitocargaorigen"))%>&comoviajad=<%=trim(rs2("quellevaorigen"))%>">
										<center><input type="submit" class="boton_3" name="descargo" value="DESCARGO"></center>
									</form>	
								</td>
							</tr>
						</table>
					<%end if%>
					<%if rs2("carga")=0 and rs2("descarga")=0 then%>
						<form name="form1" action="validasalidahrNVA.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>" method="POST">
							<center><font face="verdana" color="black">Observaciones:</font><input type="text" name="obser" maxlength="500"></center><br>
							<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
						</form>
					<%else%>
						<form name="form1" action="registrosalidaNVA.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>" method="POST">
							<center><font face="verdana" color="black">Observaciones:</font><input type="text" name="obser" maxlength="500"></center><br>
							<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
						</form>
					<%end if%>
					<br>
					<br>
					<form name="form2" action="loginhrnva.html" method="POST">
						<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
					</form>
				<%end if%>
				</center>
				<%if not isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then
					idhr=rs2("idhojaruta")
					destino=trim(rs2("destinatario"))%>
					<table width="100%" bgcolor="white" border="0">
						<tr>
							<td width="20%"><img src="LogoCATNUEVO.gif"></td>
							<td width="10%">
								<a href="cambiodeclave.asp?dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img src="cambioclave.jpg"width="70" height="70"></a>
							</td>
							<form name="form10" action="consultainformeshrweb.asp?idhr=<%=rs2("idhojaruta")%>&tipo=<%=trim(tipo)%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
								<td width="60%" align="right"><font face="verdana"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
							</form>
							<td width="20%" align="center">
								<form name="form2" action="validacerrarhrnva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>" method="POST">
									<center><a href="validacerrarhrnva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img style="width: 70px; height: 50px;" src="cerrarhr.jpg"></a></center>
								</form>
							</td>
						<tr>
					</table>
					<br>
					<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>	
					<br>
					<center><font face="verdana" size="20" color="white"><b><u>GENERACIÓN DE NUEVO TRAMO</u></b></font></center>
					<br>
					<br>
					<br>
					<br>
					<br>
					<font face="verdana" size="12" color="white"><b>Origen:&nbsp;<%=destino%></b></font>
					<br>
					<br>
					<form name="form6" action="generarnuevotramo.asp?idhr=<%=idhr%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&tipo=<%=trim(tipo)%>" method="POST">
						<table border="0">
							<tr>
								<td align="left"><font face="verdana" size="12" color="white"><b>Pr�ximo Destino:</b></font></td>
								<td>
									<%cad="select rstercero from terceros where estado='" & "A" & "' and tipomov='" & "I" & "'  order by rstercero"
									set rs=cn.execute(cad)%>
									<select name="destino" required>
										<%do while rs.eof=false
											if trim(rs("rstercero"))<>destino then%>
												<option><%=rs("rstercero")%></option>
											<%end if
											rs.movenext
										loop
										rs.close%>
									</select>
								</td>
							</tr>
							<tr><td colspan="2">&nbsp;</td></tr>
							<tr><td colspan="2">&nbsp;</td></tr>
							<tr><td align="center" colspan="2"><img src="camionnuevotramo.png" width="250" height="200"></td></tr>
							<tr><td colspan="2">&nbsp;</td></tr>
							<tr><td colspan="2">&nbsp;</td></tr>
							<tr><td align="center" colspan="2"><font face="verdana" size="20" color="white"><b>Contenido actual</b></td></tr>
							<tr><td align="center" colspan="2">
								<%if trim(rs2("quellevadestino"))="VACIO"then%>
									<font face="verdana" size="20" color="white"><b><%=(rs2("quellevadestino"))%></b>
								<%else%>
									<font face="verdana" size="20" color="white"><b><%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%></b>
								<%end if%></td>
							</tr>
							<tr><td colspan="2">&nbsp;</td></tr>
							<tr><td colspan="2">&nbsp;</td></tr>
							<tr>
								<td colspan="2" align="center">
									<input type="submit" class="boton_2" value="Generar Nuevo Tramo">
								</td>
							</tr>
							<tr><td colspan="2">&nbsp;</td></tr>
						</table>
					</form>
					<br>
					<br>
					<br>
					<br>
					<br>	
					<br>
					<center>
					<form name="form2" action="loginhrnva.html" method="POST">
						<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
					</form>
				</center>
				<%end if
			end if
		end if
		end if
	end if
end if%>

<%if movtipo="M" or movtipo="S" then
	do while rs6.eof=false
		if movtipo="M" then
        		caden="select * from viajes where idconductor=" & RS6("idconductor") & " and etapa='" & "A" & "' and fechasalida<='" & desdecadena & "' and fechallegada is null and web='" & "S" & "' and tipomov='" & "M" & "'"
		end if
		if movtipo="S" then
        		caden="select * from viajes where idconductor=" & RS6("idconductor") & " and etapa='" & "A" & "' and fechasalida<='" & desdecadena & "' and fechallegada is null and web='" & "S" & "' and tipomov='" & "S" & "'"
		end if
		set rs=cn.execute(caden)
        	If RS.EOF = False Then
            		numviaje = RS("numviaje")
			idconductor=rs6("idconductor")
			clave=rs6("claveweb")
            		encontro = 1
        	End If
        	RS.Close
        	RS6.MoveNext
    	Loop
	RS6.Close
	if encontro=0 then
		cade="select * from conductores where documento='" & request.querystring("dni") & "' and estado='" & "A" & "' and claveweb='" & request.querystring("clave") & "'"
		set rs6=cn.execute(cade)	
		If RS6.EOF = False Then
			Do While RS6.EOF = false
				if movtipo="M" then
        				caden= "select * from viajes where idconductor=" & RS6("idconductor") & " and etapa='" & "G" & "' and fechasalida<='" & desdecadena & "' and fechallegada is null and web='" & "S" & "' and tipomov='" & "M" & "' order by numviaje"
				end if
				if movtipo="S" then
        				caden= "select * from viajes where idconductor=" & RS6("idconductor") & " and etapa='" & "G" & "' and fechasalida<='" & desdecadena & "' and fechallegada is null and web='" & "S" & "' and tipomov='" & "S" & "' order by numviaje"
				end if
				set rs=cn.execute(caden)
	        		If RS.EOF = False Then
        	    			numviaje = RS("numviaje")
	    				idconductor=rs6("idconductor")
		    			clave=rs6("claveweb")
        	    			encontro = 1
            				cad="select * from viajes where idmedio=" & RS("idmedio") & " and etapa='" & "A" & "'"
					set rs7=cn.execute(cad)
        	    			If RS7.EOF = False Then
                				encontro = 0
						encontropat=1
                				RS7.Close
                				Exit Do
	            			Else
        	        			RS7.Close
            				End If
        			End If
        			RS.Close
	        		RS6.MoveNext
    			Loop
		end if
		RS6.Close
	end if
	If encontropat = 1 Then%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">	
		<html>
			<head>  
				<meta content="text/html; charset=UTF-8" http-equiv="content-type">
		  		<title>Hoja de Ruta Por Tipo</title>
				<link rel="stylesheet" href="estilohr.css">
			</head>
			<body bgcolor="#9091A0">
				<br>
				<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
				<br>
				<br>
				<div class="box-green">
					<center><font face="verdana" size="12">Su Patente tiene una Hoja de Ruta ABIERTA para este Tipo de Movimiento</font></center>
				</div>
				<br>
				<br>
				<div>
					<form name="form2" action="loginhrnva.html">
						<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
					</form>
				</div> 
			</body>
		</html>	
	<%else
		if encontro = 0 Then%>
			<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
			<html>
				<head>  
					<meta content="text/html; charset=UTF-8" http-equiv="content-type">
		  			<title>Hoja de Ruta Por Tipo</title>
					<link rel="stylesheet" href="estilohr.css">
				</head>
				<body bgcolor="#9091A0">
					<br>
					<br>
					<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
					<br>
					<div class="box-green">
						<center><font face="verdana" size="12">No existe Hoja de Ruta generada para este Tipo de Movimiento</font></center>
					</div>
					<br>
					<br>
					<center>
						<form name="form10" action="consultainformeshrweb.asp?tipo=<%=request.querystring("tipo")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
							<td width="50%" align="right"><font face="verdana" color="white" size="5"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
						</form>
					</center>
					<br>
					<br>
					<div>
						<form name="form2" action="loginhrnva.html">
							<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
						</form>
					</div> 
				</body>
			</html>	
		<%else%>
			<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
			<html>
				<head>  
					<meta content="text/html; charset=UTF-8" http-equiv="content-type">
			  		<title>Hoja de Ruta Por Tipo</title>
				</head>
				<%cade="select * from viajes where numviaje=" & numviaje
				set rs1=cn.execute (cade)
				if rs1.eof=false then
					hojaruta=trim(rs1("numviaje"))
					cad="select patente,idtipomedio from medios where idmedio="& rs1("idmedio")
					set rs3=cn.execute(cad)
					patente=trim(rs3("patente"))
					idtipo=rs3("idtipomedio")
					cad="select desctipomedio from tipomedio where idtipomedio=" & idtipo
					set rs4=cn.execute(cad)
					cade="select  * from hojaruta where numviaje=" & numviaje & " and tipo='" & request.querystring("tipo") & "' order by idhojaruta"
					set rs2=cn.execute(cade)
					if rs2.eof=false then
						tramo=0
						do while rs2.eof=false
							tramo=tramo+1
					        	entrar = 0
        						vDesde = Trim(RS2("remitente"))
        						vhasta = Trim(RS2("destinatario"))
        						If tramo = 1 Then
            							If IsNull(RS2("horallegada")) And IsNull(RS2("horasalida")) Then
                							entrar = 1
                							Exit Do
            							End If
							        If IsNull(RS2("horasalida")) And Not IsNull(RS2("horallegada")) Then
                							entrar = 1
                							Exit Do
            							End If
        						End If
        						If tramo = 2 Then
					            		If IsNull(RS2("horallegada")) And IsNull(RS2("horasalida")) Then
                							entrar = 1
                							Exit Do
            							End If
            							If IsNull(RS2("horasalida")) And Not IsNull(RS2("horallegada")) Then
                							entrar = 1
                							Exit Do
            							End If
        						End If
        						If tramo = 3 Then
            							If IsNull(RS2("horallegada")) And IsNull(RS2("horasalida")) Then
                							entrar = 1
                							Exit Do
            							End If
            							If IsNull(RS2("horasalida")) And Not IsNull(RS2("horallegada")) Then
                							entrar = 1
                							Exit Do
            							End If
        						End If
        						If tramo = 4 Then
            							If IsNull(RS2("horallegada")) And IsNull(RS2("horasalida")) Then
                							entrar = 1
                							Exit Do
            							End If
            							If IsNull(RS2("horasalida")) And Not IsNull(RS2("horallegada")) Then
                							entrar = 1
                							Exit Do
            							End If
        						End If
        						If tramo = 5 Then
            							If IsNull(RS2("horallegada")) And IsNull(RS2("horasalida")) Then
                							entrar = 1
                							Exit Do
            							End If
            							If IsNull(RS2("horasalida")) And Not IsNull(RS2("horallegada")) Then
                							entrar = 1
                							Exit Do
            							End If
        						End If
        						RS2.MoveNext
    						Loop
					end if
					rs2.close 
					if entrar=1 then
						cadena="select  top(1) * from hojaruta where numviaje=" & numviaje & " and tipo='" & request.querystring("tipo") & "' and remitente='" & vDesde & "' and destinatario='" & vhasta & "' order by idhojaruta"	
						set rs2=cn.execute(cadena)
						caden="select * from cliente where rscliente='" & trim(rs2("cliente")) & "'"
						set rs6=cn.execute(caden)
						controlacarga=rs6("controlacarga")
						rs6.close
						if isnull(rs2("horallegada")) and isnull(rs2("horasalida")) then%>
							<body bgcolor="#09C2F4">
							<table width="100%" bgcolor="white" border="0">
								<tr>
									<td width="15%"><img src="LogoCATNUEVO.gif"></td>
									<td width="10%"><a href="cambiodeclave.asp?dni=<%=dni%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>"><img src="cambioclave.jpg"width="70" height="70"></a></td>
									<td width="15%" align="right"><font size="5" color="black" face="verdana">Modificar Destino-></font></td>
									<td>
										<form name="form5" action="modihrnva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
											<a href="modihrnva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>"><img style="width: 80px; height: 80px;" src="editar.jpg"></a>
										</form>
									</td>
									<form name="form10" action="consultainformeshrweb.asp?tipo=<%=request.querystring("tipo")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
										<td width="50%" align="right"><font face="verdana"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
									</form>
								</tr>
							</table>				
							<center>	
							<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
							<br>
							<font face="verdana" size="20" color="white"><b>EN CAMINO A <%=trim(rs2("destinatario"))%></b></font>
							<br>
							<br>
							<center>	
								<img src="camioncitocargado.jpg">
							</center>
							<br>
							<font face="verdana" size="20" color="white"><b>Contenido actual</b><br>
							<%if controlacarga=0 then
								if rs2("cargado")=1 then%>
									CARGADO
								<%else%>
									VACIO
								<%end if
							else
								if trim(rs2("quellevadestino"))="VACIO" then%>
									<%=(rs2("quellevadestino"))%>
								<%else%>
									<%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%>
								<%end if%>
							<%end if%>
							<form name="form1" action="registrollegadaNVAMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">								
								<center><font face="verdana" color="black">Observaciones:</font><input type="text" name="obser" maxlength="500"></center>
								<center><input type="submit" class="boton_1" name="btnllegada" value="                  LLEGADA                  "></center>
							</form>
							<br>
							<form name="form2" action="loginhrnva.html">
								<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
							</form>
							</body>
						<%end if
						if isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then%>
							<body bgcolor="#F2B5F7">
							<table width="100%" bgcolor="white" border="0">
								<tr>
									<td width="20%"><img src="LogoCATNUEVO.gif"></td>
									<td width="20%"><a href="cambiodeclave.asp?dni=<%=dni%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>"><img src="cambioclave.jpg"width="70" height="70"></a></td>
									<form name="form10" action="consultainformeshrweb.asp?tipo=<%=request.querystring("tipo")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
										<td width="60%" align="right"><font face="verdana"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
									</form>
								<tr>
							</table>
							<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
							<br>
							<center>	
							<font face="verdana" size="20" color="black"><b>ESTOY EN&nbsp;<%=trim(rs2("destinatario"))%></b></font>
							<br>
							<br>
							<img src="mulita.jpg" width="250" height="200">
							</center>
							<br>
							<br>
							<center><font face="verdana" size="20" color="black"><b>Contenido actual</b><br>
							
								<%if trim(rs2("remitente"))="PLAYA CAT" then

									if controlacarga=0 then%>
										<center>
										<%if rs2("cargado")=1 then%>
											<font color="white" face="verdana" size="12">CARGADO</font>
										<%else%>
											<font color="white" face="verdana" size="12">VACIO</font>
										<%end if%>
										<br>
										<font color="white" face="verdana" size="12">Carg&oacute;?&nbsp;&nbsp;	
            									<input type="checkbox" name="cargo" id="cargo" value="Valor"/></font>
										</center>
										<form name="form4" action="registrosalidaNVAMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
											<center><font face="verdana" color="black" size="12">Observaciones:</font><input type="text" name="obser" maxlength="500"></center>
											<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
										</form>
									<%else
										if trim(rs2("quellevadestino"))="VACIO" then%>
											<%=(rs2("quellevadestino"))%>
										<%else%>
											<%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%>
										<%end if%>
										<%if trim(rs2("quellevadestino"))="VACIO" then%>
											<center>
											<table border="0">
												<tr>
													<td>
														<form name ="form4" method="POST" action="cargoNVAMR.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>">
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
														<form name ="form4" method="POST" action="cargoNVAMR.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>">
															<center><input type="submit" class="boton_3" name="cargo" value="CARGO"></center>
														</form>	
													</td>
													<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
													<td>
														<form name ="form4" method="POST" action="descargoNVAMR.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>&cargad=<%=rs2("cargaorigen")%>&remitocargad=<%=trim(rs2("remitocargaorigen"))%>&comoviajad=<%=trim(rs2("quellevaorigen"))%>">
															<center><input type="submit" class="boton_3" name="descargo" value="DESCARGO"></center>
														</form>	
													</td>
												</tr>
											</table>
										<%end if
										if rs2("carga")=0 and rs2("descarga")=0 then%>
											<form name="form1" action="validasalidahrNVAMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
												<center><font face="verdana" color="black">Observaciones:</font><input type="text" name="obser" maxlength="500"></center>
												<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
											</form>
										<%else%>
											<form name="form1" action="registrosalidaNVAMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
												<center><font face="verdana" color="black">Observaciones:</font><input type="text" name="obser" maxlength="500"></center>
												<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
											</form>
										<%end if%>
									<%end if%>
								<%else
									if controlacarga=0 then%>
										<center>
										<%if rs2("cargado")=1 then%>
											<font color="white" face="verdana" size="12">CARGADO</font>
										<%else%>
											<font color="white" face="verdana" size="12">VACIO</font>
										<%end if%>
										<br>
										<font color="white" face="verdana" size="12">Carg&oacute;?&nbsp;</font>	
        									<input type="checkbox" name="cargo" id="cargo" value="Valor"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<font color="white" face="verdana" size="12">Descarg&oacute;?&nbsp;</font>	
            									<input type="checkbox" name="descargo" id="descargo" value="Valor"/>
											</center>
										<form name="form1" action="registrosalidaNVAMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
											<center><font face="verdana" color="black" size="12">Observaciones:</font><input type="text" name="obser" maxlength="500"></center>
											<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
										</form>
									<%else%>
										<%if trim(rs2("quellevadestino"))="VACIO" then%>
											<%=(rs2("quellevadestino"))%>
										<%else%>
											<%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%>
										<%end if%>
										<%if trim(rs2("quellevadestino"))="VACIO" then%>
											<center>
											<table border="0">
												<tr>
													<td>
														<form name ="form4" method="POST" action="cargoNVAMR.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>">
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
														<form name ="form4" method="POST" action="cargoNVAMR.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>">
															<center><input type="submit" class="boton_3" name="cargo" value="CARGO"></center>
														</form>	
													</td>
													<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
													<td>
														<form name ="form4" method="POST" action="descargoNVAMR.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&desde=<%=trim(rs2("remitente"))%>&hasta=<%=trim(rs2("destinatario"))%>&tipo=<%=request.querystring("tipo")%>&cargad=<%=rs2("cargaorigen")%>&remitocargad=<%=trim(rs2("remitocargaorigen"))%>&comoviajad=<%=trim(rs2("quellevaorigen"))%>">
															<center><input type="submit" class="boton_3" name="descargo" value="DESCARGO"></center>
														</form>	
													</td>
												</tr>
											</table>
										<%end if
										if rs2("carga")=0 and rs2("descarga")=0 then%>
											<form name="form1" action="validasalidahrNVAMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
												<center><font face="verdana" color="black">Observaciones:</font><input type="text" name="obser" maxlength="500"></center>
												<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
											</form>
										<%else%>
											<form name="form1" action="registrosalidaNVAMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
												<center><font face="verdana" color="black">Observaciones:</font><input type="text" name="obser" maxlength="500"></center>
												<center><input type="submit" class="boton_2" name="btnllegada" value="                  SALIDA                  "></center>
											</form>
										<%end if%>
									<%end if%>
								<%end if%>


			
							<br>
							<form name="form2" action="loginhrnva.html" method="POST">
								<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
							</form>
						<%end if
						tipo=trim(rs2("tipo"))						
						if not isnull(rs2("ruta")) then%>
							<center><font face="verdana" size="3"> <a href="<%=rs2("ruta")%>">Mostrar Mapa del Recorrido</a></font></center>
						<%end if
					else
						if movtipo="S" then
							cade="select * from hojaruta where numviaje=" & numviaje & "and tipo='" & request.querystring("tipo") & "' and horasalida is null"
							set rs7=cn.execute(cade)
							if rs7.eof=true then
								cad="select top (1) * from hojaruta where numviaje=" & numviaje & " and tipo='" & request.querystring("tipo") & "' order by idhojaruta desc"
								set rs8=cn.execute(cad)
								fechac=trim(rs8("fechasalida"))
								horac=trim(rs8("horasalida"))
								rs9.CursorType = adOpenKeyset
								rs9.LockType = adLockOptimistic
								rs9.open"select * from viajes where numviaje=" & numviaje,cn
								rs9("fechallegada")=fechac
								rs9("horallegada")=horac
								rs9("etapa")="C"
								rs9.update
								rs9.close
								response.redirect("loginhrnva.html")
							else
								cadena="select  top(1) * from hojaruta where numviaje=" & numviaje & " and tipo='" & request.querystring("tipo") & "' order by idhojaruta desc"	
								set rs5=cn.execute(cadena)			
								idhr=rs5("idhojaruta")
								destino=trim(rs5("destinatario"))%>
								<body background="fondo_ondulado.jpg">
								<table width="100%" bgcolor="white" border="0">
									<tr>
										<td width="20%"><img src="LogoCATNUEVO.gif"></td>
										<td width="10%"><a href="cambiodeclave.asp?dni=<%=dni%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img src="cambioclave.jpg"width="70" height="70"></a></td>
										<form name="form10" action="consultainformeshrweb.asp?idhr=<%=rs5("idhojaruta")%>&tipo=<%=request.querystring("tipo")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
											<td width="60%" align="right"><font face="verdana"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
										</form>
										<td width="20%" align="center">
											<form name="form2" action="validacerrarhrnvaMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
												<center><a href="validacerrarhrnvaMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>"><img style="width: 70px; height: 50px;" src="cerrarhr.jpg"></a></center>
											</form>
										</td>
									</tr>
								</table>
								<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
								<br>	
								<br>
								<center><font face="verdana" size="20" color="white"><b><u>GENERACI&Oacute;N DE NUEVO TRAMO</u></b></font></center>
								<br>
								<br>
								<br>
								<br>
								<br>
								<font face="verdana" size="12" color="white"><b>Origen:&nbsp;<%=destino%></b></font>
								<br>
								<br>
								<form name="form6" action="generarnuevotramomr.asp?idhr=<%=idhr%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&tipo=<%=request.querystring("tipo")%>" method="POST">
									<table border="0">
										<tr>
											<td align="left"><font face="verdana" size="12" color="white"><b>Pr&oacute;ximo Destino:</b></font></td>
											<td>
												<%cad="select rstercero from terceros where estado='" & "A" & "' and milkrun='" & "M" & "' and perfil is null order by rstercero"
												set rs=cn.execute(cad)%>
												<select name="destino" required>
													<%do while rs.eof=false
														if trim(rs("rstercero"))<>destino then%>
															<option><%=rs("rstercero")%></option>
														<%end if
														rs.movenext
													loop
													rs.close%>
												</select>
											</td>
										</tr>
										<tr><td colspan="2">&nbsp;</td></tr>
										<tr><td colspan="2">&nbsp;</td></tr>
										<tr><td align="center" colspan="2"><img src="camionnuevotramo.png" width="250" height="200"></td></tr>
										<tr><td colspan="2">&nbsp;</td></tr>
										<tr><td colspan="2">&nbsp;</td></tr>
										<tr><td align="center" colspan="2"><font face="verdana" size="20" color="white"><b>Contenido actual</b></td></tr>
										<tr><td align="center" colspan="2">
										<%if trim(rs5("quellevadestino"))="VACIO"then%>
											<font face="verdana" size="20" color="white"><b><%=(rs5("quellevadestino"))%></b>
										<%else%>
											<font face="verdana" size="20" color="white"><b><%=rs5("cargadestino") & " " & trim(rs5("quellevadestino"))%></b>
										<%end if%></td>
										</tr>
										<tr><td colspan="2">&nbsp;</td></tr>
										<tr><td colspan="2">&nbsp;</td></tr>
										<tr>
											<td colspan="2" align="center">
												<input type="submit" class="boton_2" value="Generar Nuevo Tramo">
											</td>
										</tr>
										<tr><td colspan="2">&nbsp;</td></tr>
									</table>
								</form>
								<br>
								<br>
								<br>
								<br>
								<br>	
								<br>
								<center>
								<form name="form2" action="loginhrnva.html" method="POST">
									<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
								</form>
								</center>
							<%end if
						else
							cadena="select  top(1) * from hojaruta where numviaje=" & numviaje & " and tipo='" & request.querystring("tipo") & "' order by idhojaruta desc"	
							set rs5=cn.execute(cadena)			
							idhr=rs5("idhojaruta")
							destino=trim(rs5("destinatario"))%>
							<body background="fondo_ondulado.jpg">
							<table width="100%" bgcolor="white" border="0">
								<tr>
									<td width="20%"><img src="LogoCATNUEVO.gif"></td>
									<td width="10%"><a href="cambiodeclave.asp?dni=<%=dni%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img src="cambioclave.jpg"width="70" height="70"></a></td>
									<form name="form10" action="consultainformeshrweb.asp?idhr=<%=rs5("idhojaruta")%>&tipo=<%=request.querystring("tipo")%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">					
										<td width="60%" align="right"><font face="verdana"><b>Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenadesde%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenadesde%>" required></b></font>&nbsp;&nbsp;<input type="submit" class="boton_4" value="Ver Actividad"></td>
									</form>
									<td width="20%" align="center">
										<form name="form2" action="validacerrarhrnvaMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>" method="POST">
											<center><a href="validacerrarhrnvaMR.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=request.querystring("tipo")%>&clave=<%=request.querystring("clave")%>"><img style="width: 70px; height: 50px;" src="cerrarhr.jpg"></a></center>
										</form>
									</td>
								</tr>
							</table>
							<center><font face="verdana" size="20" color="white"><%=request.querystring("tipo")%></font></center>
							<br>	
							<br>
							<center><font face="verdana" size="20" color="white"><b><u>GENERACI&Oacute;N DE NUEVO TRAMO</u></b></font></center>
							<br>
							<br>
							<br>
							<br>
							<br>
							<font face="verdana" size="12" color="white"><b>Origen:&nbsp;<%=destino%></b></font>
							<br>
							<br>
							<form name="form6" action="generarnuevotramomr.asp?idhr=<%=idhr%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>&tipo=<%=request.querystring("tipo")%>" method="POST">
								<table border="0">
									<tr>
										<td align="left"><font face="verdana" size="12" color="white"><b>Pr&oacute;ximo Destino:</b></font></td>
										<td>
											<%cad="select rstercero from terceros where estado='" & "A" & "' and milkrun='" & "M" & "' and perfil is null order by rstercero"
											set rs=cn.execute(cad)%>
											<select name="destino" required>
												<%do while rs.eof=false
													if trim(rs("rstercero"))<>destino then%>
														<option><%=rs("rstercero")%></option>
													<%end if
													rs.movenext
												loop
												rs.close%>
											</select>
										</td>
									</tr>
									<tr><td colspan="2">&nbsp;</td></tr>
									<tr><td colspan="2">&nbsp;</td></tr>
									<tr><td align="center" colspan="2"><img src="camionnuevotramo.png" width="250" height="200"></td></tr>
									<tr><td colspan="2">&nbsp;</td></tr>
									<tr><td colspan="2">&nbsp;</td></tr>
									<tr><td align="center" colspan="2"><font face="verdana" size="20" color="white"><b>Contenido actual</b></td></tr>
									<tr><td align="center" colspan="2">
									<%if trim(rs5("quellevadestino"))="VACIO"then%>
										<font face="verdana" size="20" color="white"><b><%=(rs5("quellevadestino"))%></b>
									<%else%>
										<font face="verdana" size="20" color="white"><b><%=rs5("cargadestino") & " " & trim(rs5("quellevadestino"))%></b>
									<%end if%></td>
									</tr>
									<tr><td colspan="2">&nbsp;</td></tr>
									<tr><td colspan="2">&nbsp;</td></tr>
									<tr>
										<td colspan="2" align="center">
											<input type="submit" class="boton_2" value="Generar Nuevo Tramo">
										</td>
									</tr>
									<tr><td colspan="2">&nbsp;</td></tr>
								</table>
							</form>
							<br>
							<br>
							<br>
							<br>
							<br>	
							<br>
							<center>
							<form name="form2" action="loginhrnva.html" method="POST">
								<center><a href="loginhrnva.html"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
							</form>
							</center>
						<%end if
					end if
				end if%>
			</html>			
			<%end if
		end if
	end if
end if%>