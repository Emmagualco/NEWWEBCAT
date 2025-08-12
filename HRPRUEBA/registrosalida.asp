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
rs.open "Select * from hojaruta where numviaje=" & request.querystring("hr") & " and fechallegada is not null and fechasalida is null",cn
rs("fechasalida")=date()
rs("horasalida")=time()
rs.update
rs.close
%>
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
tipo=request.querystring("tipo")
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARGPRUEBA"
cn.open(cad)
productos_Sql = "Select * from conductores  where documento='" & dni & "' and estado='" & "A" & "'"
set rs = CreateObject("ADODB.Recordset")
rs.Open productos_Sql, Cn
if rs.eof=false then%>
	<%idconductor=rs("idconductor")%>
	<center>	
	<table border="0">
		<tbody>
			<tr>
				<td><img src="<%=rs("foto")%>" width="150" height="150"></td>				
				<td><font face="verdana" size="10"><b>Conductor:&nbsp;<%=rs("nombre")%></font></b></td>
				<td><a href="javascript:Abrir_iqnet('consultalc.asp?idconductor=<%=idconductor%>&tipo=<%=tipo%>')"><img src="treslineas.png"width="70" height="50"></a></td>
			</tr>
		</tbody>
	</table>
	</center>
	<br>
	<%hoy=month(now())&"/"&day(now())&"/"&year(now())%>
	<%cadena="select * from viajes where numviaje=" & request.querystring("hr")
	set rs1=cn.execute (cadena)
	if rs1.eof=false then
		hojaruta=trim(rs1("numviaje"))
		cad="select patente,idtipomedio from medios where idmedio="& rs1("idmedio")
		set rs3=cn.execute(cad)
		patente=trim(rs3("patente"))
		idtipo=rs3("idtipomedio")
		cad="select desctipomedio from tipomedio where idtipomedio=" & idtipo
		set rs4=cn.execute(cad)%>
		<center><font face="verdana" size="12" color="green"><b>Dominio:&nbsp;</b></font><font face="verdana" size="7" color="green"><%=patente%>-<%=trim(rs4("desctipomedio"))%></font></center><br>
		<center><font face="verdana" size="12"><b>Hoja de Ruta: &nbsp;<%=rs1("numviaje")%></b></center><br>
		<%cade="select * from hojaruta where numviaje=" & rs1("numviaje") & " and (horallegada is null or horasalida is null) and tipo='" & request.querystring("tipo") & "' order by idhojaruta desc"
		set rs2=cn.execute(cade)
		if rs2.eof=false then%>
			<center>
			<%tipo=rs2("tipo")%>
			<font face="verdana" size="10"><b>Tipo:&nbsp;<%=tipo%><b></font>
			CLIENTE:<%=rs2("cliente")%>
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
			<br>
			<%if isnull(rs2("horallegada")) and isnull(rs2("horasalida")) then%>
				<form name="form5" action="modihr.asp" method="POST">
				<table style="text-align: left; width: 800px;" border="0" cellpadding="2" cellspacing="2">
					<tbody>
						<tr>
							<td><font size="5"><span style="font-family: Verdana;">Para Modificar el Origen o el Destino, haga Click en la imágen-></span></font></td>
							<td><a href="modihr.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&desde=<%=rs2("remitente")%>&hasta=<%=rs2("destinatario")%>&tipo=<%=tipo%>"><img style="width: 100px; height: 100px;" src="editar.jpg"></a></td>
						</tr>
					</tbody>
				</table>					
				</form>
				<%llega="Llego a " & trim(rs2("destinatario"))%>
				<form name="form1" action="registrollegada.asp?hr=<%=rs1("numviaje")%>&dni=<%=dni%>&tipo=<%=tipo%>" method="POST">
					<center><input type="submit" class="boton_2" name="btnllegada" value="<%=llega%>"></center>
				</form>
				<form name="form2" action="loginhr.html">
					<center><input type="submit" class="boton_2" value="Ir a la Página Principal"></center>
				</form>
			<%end if%>

			<%if isnull(rs2("horasalida")) and not isnull(rs2("horallegada"))then%>
				<font face="verdana" size="6" color="red"><b>Fecha Llegada: &nbsp;<%=rs2("fechallegada")%><br>
				&nbsp;&nbsp;Hora Llegada:&nbsp;<%=rs2("horallegada")%></b></font>
				<br><br>
				<%sale="Salgo de " & trim(rs2("destinatario"))%>
				<form name="form1" action="registrosalida.asp?hr=<%=rs1("numviaje")%>&dni=<%=dni%>&tipo=<%=tipo%>"method="POST">
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
					<%tipo=request.querystring("tipo")%>
					<font face="verdana" size="10"><b>Tipo:&nbsp;<%=tipo%><b></font>
					<br>
					<font face="verdana" size="10" color="blue"><b><u>GENERACIÓN DE TRAMO</u></b></font><br><br>
					<%cade="select top(1) destinatario,cliente from hojaruta where numviaje=" & rs1("numviaje") & " and tipo='" & request.querystring("tipo")& "' order by idhojaruta desc"
					set rs4=cn.execute(cade)
					oriori=rs4("destinatario")
					clicli=rs4("cliente")
					rs4.close%>
					<form name="form1" action="generarhr.asp?hr=<%=rs1("numviaje")%>&dni=<%=dni%>&hojaruta=<%=hojaruta%>&tipo=<%=tipo%>" method="POST">
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
				</body>
			</html>
		<%end if%>
	<%end if%>		 
<%end if%>
</body>
</html>
