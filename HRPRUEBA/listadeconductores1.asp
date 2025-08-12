<html>

<head>
<meta http-equiv="refresh" content="30">
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
    padding: 10px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 100;
    font-size: 25px;
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
    padding: 5px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 50;
    font-size: 20px;
    color: #FFFFFF;
    background-color:#05C60E;
    border-radius: 15px;
    border: 2px solid #27AE60;
  }
  .boton_3:hover{
    color: #9b0e0e;
    background-color:  #AED581;
    text-decoration: none;
  }


  .boton_5{
    text-decoration: none;
    padding: 5px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 50;
    font-size: 20px;
    color: #FFFFFF;
    background-color:#8B97F1;
    border-radius: 15px;
    border: 2px solid #27AE60;
  }
  .boton_5:hover{
    color: #9b0e0e;
    background-color:  #AED581;
    text-decoration: none;
  }


  .boton_6{
    text-decoration: none;
    padding: 5px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 50;
    font-size: 25px;
    color: #FFFFFF;
    background-color:#5007DA;
    border-radius: 15px;
    border: 2px solid #27AE60;
  }
  .boton_6:hover{
    color: #9b0e0e;
    background-color:  #94CCF8;
    text-decoration: none;
  }


  .boton_7{
    text-decoration: none;
    padding: 5px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 50;
    font-size: 20px;
    color: #000000;
    background-color:#F5F506;
    border-radius: 15px;
    border: 2px solid #27AE60;
  }
  .boton_7:hover{
    color: #9b0e0e;
    background-color:  #94CCF8;
    text-decoration: none;
  }

  .boton_8{
    text-decoration: none;
    padding: 5px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 50;
    font-size: 20px;
    color: #FFFFFF;
    background-color:#F50522;
    border-radius: 15px;
    border: 2px solid #27AE60;
  }
  .boton_8:hover{
    color: #9b0e0e;
    background-color:  #94CCF8;
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
   font-size: 13px;
   font-face: "Century Gothic"
   height: 30px;
   padding: 0px;
   width: 170px;
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

input[type=radio] { border: 0px; width: 5%; height: 1em; }


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

input[type=password]:focus{
border-color:dodgerBlue;
box-shadow:0 0 8px 0 dodgerblue;
}

input[type=password]{
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
set rs7=server.createobject("ADODB.Recordset")
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
<%usuario=request("usuario")%>
<%clave=request("clave")
encontrado=0
sitio=""
cade="select * from usuariosrasa where nombre='" & request("usuario") & "' and clave='" & request("clave") & "'"
set rs7=cn.execute(cade)
if rs7.eof=false then
	sitio=rs7("sitio")
end if
if sitio<>"" then%>
<body background="fondo_ondulado.jpg">
<table width="100%" border="0">
	<tr>
		<td rowspan="2"><img src="logocatnuevo.gif" width="100" height="100"></td>
		<td rowspan="2"><a href="cambioclavecliente.asp?usuario=<%=usuario%>&clave=<%=clave%>"><img src="cambioclave.jpg"width="60" height="60"></a></td>
		<td rowspan="2" align="center" width="60%"><font face="verdana" size="6" color="white">Bienvenido <%=request("usuario")%></font></td>
		<td rowspan="2" align="center">
			<form name="form3" action="hrabiertasMI.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<input type="submit" class="boton_3" value="HR ABIERTAS MOV. INT.">
			</form>		
			<form name="form3" action="hrabiertasMR.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<input type="submit" class="boton_3" value="HR ABIERTAS MILK RUN">
			</form>	
			<form name="form3" action="hrabiertasSP.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<input type="submit" class="boton_3" value="HR ABIERTAS SPOT">
			</form>		
		</td>
		<td rowspan="2" align="center">			
			<form name="form4" action="hrvalidadas.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<br>
				<input type="submit" class="boton_5" value="HR VALIDADAS">
			</form>
		</td>
		<td rowsapn="2" align="center">			
			<form name="form5" action="informe.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<select name="movimi" required>
  					<option>MILK RUN</option>
  					<option>MOVIMIENTOS INTERNOS</option>
  					<option>SPOT</option>
				</select>
				<input type="submit" class="boton_5" value="INFORMES">
			</form>
		</td>
		<td rowsapn="2" align="center">			
			<form name="form5" action="tablerodiariomi.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<input type="submit" class="boton_3" value="TABLERO DIARIO MOV. INT.">
			</form>
			<form name="form6" action="tablerodiariomr.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<input type="submit" class="boton_3" value="TABLERO DIARIO MILK RUN">
			</form>
			<form name="form6" action="tablerodiariosp.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
				<input type="submit" class="boton_3" value="TABLERO DIARIO SPOT">
			</form>
		</td>

	</tr>
	<tr>
	</tr>
</table>
<%if sitio="R" then%>
	<hr>
	<center>

	<table border="0" width="100%">
		<tr bgcolor="white">
			<td colspan="4" align="center"><font face="verdana" size="5">Seleccione Conductor para validar LLEGADA y SALIDA</font></td>
		</tr>
		<tr>
			<td colspan="4">&nbsp;</td>
		</tr>
		<tr bgcolor="#0E6289">
			<td><font face="verdana" size="10" color="white"><b>Conductor</b></font></td>
			<td><font face="verdana" size="10" color="white"><b>Patente</font></b></td>
			<td><font face="verdana" size="10" color="white"><b>Primer Destino</font></b></td>			
			<td align="center"><font face="verdana" size="10" color="white"><b>Acción</b></font></td>
		</tr>
		<tr>
			<td colspan="4">&nbsp;</td>
		</tr>

		<%cadena="select v.numviaje,c.nombre,m.patente from viajes v, conductores c, medios m where v.idconductor=c.idconductor and m.idmedio=v.idmedio and (v.tipomov='" & "I" & "' or v.tipomov='" & "S" & "') and v.etapa='" & "A" & "' and web='" & "S" & "' and v.horainiciovalidada is null and v.horafinvalidada is null order by c.nombre"
		set rs=cn.execute(cadena)
		if rs.eof=false then
			encontrado=1%>
			<%do while rs.eof=false%>
				<tr bgcolor="#AAD2E5">
					<td width="45%" valign="middle"><font face="verdana" size="6">					
						<%=rs("nombre")%></font>
					</td>					
					<td width="13%"><font face="verdana" size="6">
						<%=rs("patente")%></font>
					</td>
					<%cadena="select top (1) destinatario from hojaruta where numviaje=" & rs("numviaje") & " order by idhojaruta"
					set rs5=cn.execute(cadena)
					primer=trim(rs5("destinatario"))%>
					<td width="28%"><font face="verdana" size="6">
						<%=primer%></font>
					</td>				
					<%caden="select top (1) destinatario from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta"
					set rs7=cn.execute(caden)
					remit=trim(rs7("destinatario"))%>
					<td width="14%" align="center">
						<a href="validarllegada.asp?conductor=<%=rs("nombre")%>&idhr=<%=rs("numviaje")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>"><input type="button" class="boton_7" value="Validar Inicio del Servicio"></a>
						<center><font face="verdana" size="3"><b><%=remit%></b></font></center>
					</td>
				</tr>
				<tr>
					<td colspan="4" bgcolor="#AAD2E5"><hr></td>
				</tr>
				<%rs.movenext
			loop%>
		<%end if%>	



		<%cadena="select v.numviaje,c.nombre,m.patente from viajes v, conductores c, medios m where v.idconductor=c.idconductor and m.idmedio=v.idmedio and (v.tipomov='" & "I" & "' or v.tipomov='" & "S" & "') and v.etapa='" & "C" & "' and web='" & "S" & "' and v.horainiciovalidada is not null and v.horafinvalidada is null order by c.nombre"
		set rs=cn.execute(cadena)
		if rs.eof=false then
			encontrado=1%>
			<%do while rs.eof=false%>
				<tr bgcolor="#AAD2E5">
					<td width="45%" valign="middle"><font face="verdana" size="6">					
						<%=rs("nombre")%></font>
					</td>					
					<td width="13%"><font face="verdana" size="6">
						<%=rs("patente")%></font>
					</td>
					<%cadena="select top (1) destinatario from hojaruta where numviaje=" & rs("numviaje") & " order by idhojaruta"
					set rs5=cn.execute(cadena)
					primer=trim(rs5("destinatario"))%>
					<td width="28%"><font face="verdana" size="6">
						<%=primer%></font>
					</td>				
					<%caden="select top (1) destinatario from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta desc"
					set rs7=cn.execute(caden)
					desti=trim(rs7("destinatario"))%>
					<td width="14%" align="center">
						<a href="validarsalida.asp?conductor=<%=rs("nombre")%>&idhr=<%=rs("numviaje")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>"><input type="button" class="boton_3" value="Validar Fin del Servicio"></a>
						<center><font face="verdana" size="3"><b><%=desti%></b></font></center>
					</td>
				</tr>
				<tr>
					<td colspan="4" bgcolor="#AAD2E5"><hr></td>
				</tr>
				<%rs.movenext
			loop%>
		<%end if%>	

		<%cadena="select v.numviaje,c.nombre,m.patente from viajes v, conductores c, medios m where v.idconductor=c.idconductor and m.idmedio=v.idmedio and (v.tipomov='" & "I" & "' or v.tipomov='" & "S" & "') and v.etapa='" & "C" & "' and web='" & "S" & "' and v.horainiciovalidada is null and v.horafinvalidada is null order by c.nombre"
		set rs=cn.execute(cadena)
		if rs.eof=false then
			encontrado=1%>
			<%do while rs.eof=false%>
				<tr bgcolor="#AAD2E5">
					<td width="45%" valign="middle"><font face="verdana" size="6">					
						<%=rs("nombre")%></font>
					</td>					
					<td width="13%"><font face="verdana" size="6">
						<%=rs("patente")%></font>
					</td>
					<%cadena="select top (1) destinatario from hojaruta where numviaje=" & rs("numviaje") & " order by idhojaruta"
					set rs5=cn.execute(cadena)
					primer=trim(rs5("destinatario"))%>
					<td width="28%"><font face="verdana" size="6">
						<%=primer%></font>
					</td>				
					<%caden="select top (1) remitente from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta"
					set rs7=cn.execute(caden)
					remit=trim(rs7("remitente"))%>
					<%caden="select top (1) destinatario from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta desc"
					set rs7=cn.execute(caden)
					desti=trim(rs7("destinatario"))%>
					<td width="14%" align="center">
						<a href="validarllegadaysalida.asp?conductor=<%=rs("nombre")%>&idhr=<%=rs("numviaje")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>"><input type="button" class="boton_8" value="Validar Inicio y Fin del servicio"></a>
						<center><font face="verdana" size="3"><b><%=remit%> - <%=desti%></b></font></center>
					</td>
				</tr>
				<tr>
					<td colspan="4" bgcolor="#AAD2E5"><hr></td>
				</tr>
				<%rs.movenext
			loop%>
		<%end if%>	

	</table>

	<%if encontrado=1 then%>
		<br>
		<br>
		<br>
		<br>
		<table border="0">
			<form name="form3" action="nuevotramo.html">
				<td align="center">
					<button class="boton_2" onclick="location.href='nuevotramo.html'">REGRESAR</button>
			</form>
			</td>
			<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<div>
					<form name="form5" action="listadeconductores1.asp?usuario=<%=request("usuario")%>&clave=<%=request("clave")%>" method="POST">
						<td align="center">
						<center><input type="submit" class="boton_2" value="REFRESCAR"></center>
					</form>
				</div> 
			</td>
			</tr>
		</table>
	<%else%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
		<head>  
		<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
		<title>Sin Conductores para Validar</title>
		</head>
		<body background="fondo_ondulado.jpg">
			<br>
			<br>
			<br>
			<div class="box-green">
					<center><font face="verdana" size="12" color="white">No existen conductores para validar LLEGADA y SALIDA.</font></center>
			</div>
			<br>
			<br>
			<table border="0">
				<tr>
					<td>
						<div>
							<form name="form3" action="nuevotramo.html">
								<td align="center">
								<center><button class="boton_2" onclick="location.href='nuevotramo.html'">REGRESAR</button></center>
							</form>
						</div> 
					</td>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<div>
						<form name="form5" action="listadeconductores1.asp?usuario=<%=request("usuario")%>&clave=<%=request("clave")%>" method="POST">
							<td align="center">
							<center><input type="submit" class="boton_2" value="REFRESCAR"></center>
						</form>
					</div> 
					</td>
				</tr>
			</table>
		</body>
		</html>	
	<%end if%>

<%else%>
	<hr>
	<center>
	<table border="0" width="100%">
		<tr bgcolor="white">
			<td colspan="6" align="center"><font face="verdana" size="5">Seleccione Conductor para validar LLEGADA y SALIDA</font></td>
		</tr>
		<tr>
			<td colspan="6">&nbsp;</td>
		</tr>
		<tr bgcolor="#0E6289">
			<td><font face="verdana" size="10" color="white"><b>H.R.</b></font></td>
			<td><font face="verdana" size="10" color="white"><b>Fecha</b></font></td>
			<td><font face="verdana" size="10" color="white"><b>Conductor</b></font></td>
			<td><font face="verdana" size="10" color="white"><b>Patente</font></b></td>
			<td><font face="verdana" size="10" color="white"><b>Primer Destino</font></b></td>
			<td align="center"><font face="verdana" size="10" color="white"><b>Acción</b></font></td>
		</tr>
		<tr>
			<td colspan="6">&nbsp;</td>
		</tr>


		<%cadena="select v.numviaje,c.nombre,m.patente,v.fechasalida from viajes v, conductores c, medios m where v.idconductor=c.idconductor and m.idmedio=v.idmedio and (v.tipomov='" & "I" & "' or v.tipomov='" & "S" & "') and v.etapa='" & "A" & "' and web='" & "S" & "' and v.horainiciovalidada is null and v.horafinvalidada is null order by c.nombre"
		set rs=cn.execute(cadena)
		if rs.eof=false then
			encontrado=1%>
			<%do while rs.eof=false%>
				<tr bgcolor="#AAD2E5">
					<td width="10%" valign="middle"><font face="verdana" size="6">	
						<a href="javascript:Abrir_iqnet('consultalchr.asp?numviaje=<%=rs("numviaje")%>&usuario=<%=usuario%>&clave=<%=clave%>')"><%=rs("numviaje")%></a>
					</td>
					<td width="10%" valign="middle"><font face="verdana" size="6">					
						<%=rs("fechasalida")%></font>
					</td>															
					<td width="40%" valign="middle"><font face="verdana" size="6">					
						<%=rs("nombre")%></font>
					</td>					
					<td width="20%"><font face="verdana" size="6">
						<%=rs("patente")%></font>
					</td>
					<%cadena="select top (1) destinatario from hojaruta where numviaje=" & rs("numviaje") & " order by idhojaruta"
					set rs5=cn.execute(cadena)
					primer=trim(rs5("destinatario"))%>
					<td width="28%"><font face="verdana" size="6">
						<%=primer%></font>
					</td>				
					<%caden="select top (1) destinatario from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta"
					set rs7=cn.execute(caden)
					remit=trim(rs7("destinatario"))%>
					<td width="20%" align="center">
						<a href="validarllegada.asp?conductor=<%=rs("nombre")%>&idhr=<%=rs("numviaje")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>"><input type="button" class="boton_7" value="Validar Inicio del Servicio"></a>
						<center><font face="verdana" size="3"><b><%=remit%></b></font></center>
					</td>
				</tr>
				<tr>
					<td colspan="6" bgcolor="#AAD2E5"><hr></td>
				</tr>
				<%rs.movenext
			loop%>
		<%end if%>	



		<%cadena="select v.numviaje,c.nombre,m.patente,v.fechasalida from viajes v, conductores c, medios m where v.idconductor=c.idconductor and m.idmedio=v.idmedio and (v.tipomov='" & "I" & "' or v.tipomov='" & "S" & "') and v.etapa='" & "C" & "' and web='" & "S" & "' and v.horainiciovalidada is not null and v.horafinvalidada is null order by c.nombre"
		set rs=cn.execute(cadena)
		if rs.eof=false then
			encontrado=1%>
			<%do while rs.eof=false%>
				<tr bgcolor="#AAD2E5">
					<td width="10%" valign="middle"><font face="verdana" size="6">	
						<a href="javascript:Abrir_iqnet('consultalchr.asp?numviaje=<%=rs("numviaje")%>&usuario=<%=usuario%>&clave=<%=clave%>')"><%=rs("numviaje")%></a>
					</td>
					<td width="10%" valign="middle"><font face="verdana" size="6">					
						<%=rs("fechasalida")%></font>
					</td>															
					<td width="40%" valign="middle"><font face="verdana" size="6">					
						<%=rs("nombre")%></font>
					</td>					
					<td width="20%"><font face="verdana" size="6">
						<%=rs("patente")%></font>
					</td>
					<%cadena="select top (1) destinatario from hojaruta where numviaje=" & rs("numviaje") & " order by idhojaruta"
					set rs5=cn.execute(cadena)
					primer=trim(rs5("destinatario"))%>
					<td width="28%"><font face="verdana" size="6">
						<%=primer%></font>
					</td>				
					<%caden="select top (1) destinatario from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta desc"
					set rs7=cn.execute(caden)
					desti=trim(rs7("destinatario"))%>
					<td width="20%" align="center">
						<a href="validarsalida.asp?conductor=<%=rs("nombre")%>&idhr=<%=rs("numviaje")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>"><input type="button" class="boton_3" value="Validar Fin del Servicio"></a>
						<center><font face="verdana" size="3"><b><%=desti%></b></font></center>
					</td>
				</tr>
				<tr>
					<td colspan="6" bgcolor="#AAD2E5"><hr></td>
				</tr>
				<%rs.movenext
			loop%>
		<%end if%>	

		<%cadena="select v.numviaje,c.nombre,m.patente,v.fechasalida from viajes v, conductores c, medios m where v.idconductor=c.idconductor and m.idmedio=v.idmedio and (v.tipomov='" & "I" & "' or v.tipomov='" & "S" & "') and v.etapa='" & "C" & "' and web='" & "S" & "' and v.horainiciovalidada is null and v.horafinvalidada is null order by c.nombre"
		set rs=cn.execute(cadena)
		if rs.eof=false then
			encontrado=1%>
			<%do while rs.eof=false%>
				<tr bgcolor="#AAD2E5">
					<td width="10%" valign="middle"><font face="verdana" size="6">	
						<a href="javascript:Abrir_iqnet('consultalchr.asp?numviaje=<%=rs("numviaje")%>&usuario=<%=usuario%>&clave=<%=clave%>')"><%=rs("numviaje")%></a>
					</td>
					<td width="10%" valign="middle"><font face="verdana" size="6">					
						<%=rs("fechasalida")%></font>
					</td>										
					<td width="40%" valign="middle"><font face="verdana" size="6">					
						<%=rs("nombre")%></font>
					</td>					
					<td width="20%"><font face="verdana" size="6">
						<%=rs("patente")%></font>
					</td>
					<%cadena="select top (1) destinatario from hojaruta where numviaje=" & rs("numviaje") & " order by idhojaruta"
					set rs5=cn.execute(cadena)
					primer=trim(rs5("destinatario"))%>
					<td width="28%"><font face="verdana" size="6">
						<%=primer%></font>
					</td>				
					<%caden="select top (1) destinatario from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta"
					set rs7=cn.execute(caden)
					primer=trim(rs7("destinatario"))%>
					<%caden="select top (1) destinatario from hojaruta where numviaje=" &  rs("numviaje") & " order by idhojaruta desc"
					set rs7=cn.execute(caden)
					desti=trim(rs7("destinatario"))%>
					<td width="20%" align="center">
						<a href="validarllegadaysalida.asp?conductor=<%=rs("nombre")%>&idhr=<%=rs("numviaje")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>"><input type="button" class="boton_8" value="Validar Inicio y Fin del servicio"></a>
						<center><font face="verdana" size="3"><b><%=primer%> - <%=desti%></b></font></center>
					</td>
				</tr>
				<tr>
					<td colspan="6" bgcolor="#AAD2E5"><hr></td>
				</tr>
				<%rs.movenext
			loop%>
		<%end if%>	

	</table>

	<%if encontrado=1 then%>
		<br>
		<br>
		<br>
		<br>
		<table border="0">
			<form name="form3" action="nuevotramo.html">
			<td align="center">
				<button class="boton_2" onclick="location.href='nuevotramo.html'">REGRESAR</button>
			</form>
			</td>
			<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<div>
					<form name="form5" action="listadeconductores1.asp?usuario=<%=request("usuario")%>&clave=<%=request("clave")%>" method="POST">
						<td align="center">
						<center><input type="submit" class="boton_2" value="REFRESCAR"></center>
					</form>
				</div> 
			</td>
			</tr>
		</table>
	<%else%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
		<head>  
		<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
		<title>Sin Conductores para Validar</title>
		</head>
		<body background="fondo_ondulado.jpg">
			<br>
			<br>
			<br>
			<div class="box-green">
				<center><font face="verdana" size="12" color="white">No existen conductores para validar LLEGADA y SALIDA.</font></center>
			</div>
			<br>
			<br>
			<table border="0">
				<tr>
					<td>
						<div>
							<form name="form3" action="nuevotramo.html">
								<td align="center">
								<center><button class="boton_2" onclick="location.href='nuevotramo.html'">REGRESAR</button></center>
							</form>
						</div> 
					</td>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<div>
						<form name="form5" action="listadeconductores1.asp?usuario=<%=request("usuario")%>&clave=<%=request("clave")%>" method="POST">
							<td align="center">
							<center><input type="submit" class="boton_2" value="REFRESCAR"></center>
						</form>
					</div> 
					</td>
				</tr>
			</table>
		</body>
		</html>	
	<%end if%>
<%end if%>
<%end if%>