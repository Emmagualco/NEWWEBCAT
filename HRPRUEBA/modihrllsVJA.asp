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
   width: 460px;
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
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
productos_Sql = "Select * from conductores  where documento='" & request.querystring("dni") & "' and estado='" & "A" & "'"
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
	<%mes=month(now)
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
	<%cadena="select * from viajes where numviaje=" & request.querystring("hr")
	set rs1=cn.execute (cadena)
	if rs1.eof=false then
		hojaruta=trim(rs1("hojaderuta"))
		cad="select patente,idtipomedio from medios where idmedio="& rs1("idmedio")
		set rs3=cn.execute(cad)
		patente=trim(rs3("patente"))
		idtipo=rs3("idtipomedio")
		cad="select desctipomedio from tipomedio where idtipomedio=" & idtipo
		set rs4=cn.execute(cad)%>
		<center><font face="verdana" size="12" color="green"><b>Dominio:&nbsp;</b></font><font face="verdana" size="7" color="green"><%=patente%>-<%=trim(rs4("desctipomedio"))%></font></center><br>
		<center><font face="verdana" size="12"><b>Hoja de Ruta: &nbsp;<%=rs1("numviaje")%></b></center>
		<%cade="select * from hojaruta where numviaje=" & rs1("numviaje") & " and (horallegada is null or horasalida is null)"
		set rs2=cn.execute(cade)
		if rs2.eof=false then%>
			<%contenidoo=trim(rs2("comoviaja"))
			contenidod=trim(rs2("comoviajad"))%>
			<center>
			<%tipo=rs2("tipo")%>
			<font face="verdana" size="10"><b>Tipo:&nbsp;<%=tipo%><b></font><br>
			<font face="verdana" size="10"><b>Cliente:&nbsp;<%=rs2("cliente")%><b></font>			
			<br>
			<table border="0">
				<tr>
					<td colspan="2"><hr></td>
				</tr>
				<tr>
					<td align="right"><font face="verdana" size="12"><b>Origen:</b></font></td>
					<td><font face="verdana" size="12" color="blue"><%=rs2("remitente")%></font></td>
				</tr>
				<tr>
					<td align="right"><font face="verdana" size="6"><b>Cant. Carga:&nbsp;<span class="blanco-violeta"><%=rs2("cargao")%></span></b></font></td>
					<td><font face="verdana" size="6">&nbsp;&nbsp;&nbsp;<b>Cant. Descarga:&nbsp;<span class="blanco-naranja"><%=rs2("descargao")%></span></b></font></td>
				</tr>
				<tr>
					<td align="left" colspan="2"><font face="verdana" size="6"><b>N° Remitos Carga:&nbsp;<span class="blanco-violeta"><%=rs2("remitocargao")%></span></b></font></td>
				</tr>
				<tr>
					<td align="left" colspan="2"><font face="verdana" size="6"><b>N° Remitos Descarga:&nbsp;<span class="blanco-naranja"><%=rs2("remitodescargao")%></span></b></font></td>
				</tr>
				<tr>
					<td colspan="2"><font face="verdana" size="10" color="red"><b>Contenido Origen:&nbsp;<%=rs2("comoviaja")%><b></font></td>
				</tr>
				<tr>
					<td colspan="2"><hr></td>
				</tr>
				<tr>
					<td align="right"><font face="verdana" size="12"><b>Destino:</b></font></td>
					<td><font face="verdana" size="12" color="blue"><%=rs2("destinatario")%></font></td>
				</tr>
				<tr>
					<td align="right"><font face="verdana" size="6"><b>Cant. Carga:&nbsp;<span class="blanco-violeta"><%=rs2("cargad")%></span></b></font></td>
					<td><font face="verdana" size="6">&nbsp;&nbsp;&nbsp;<b>Cant. Descarga:&nbsp;<span class="blanco-naranja"><%=rs2("descargad")%></span></b></font></td>
				</tr>
				<tr>
					<td align="left" colspan="2"><font face="verdana" size="6"><b>N° Remitos Carga:&nbsp;<span class="blanco-violeta"><%=rs2("remitocargad")%></span></b></font></td>
				</tr>
				<tr>
					<td align="left" colspan="2"><font face="verdana" size="6"><b>N° Remitos Descarga:&nbsp;<span class="blanco-naranja"><%=rs2("remitodescargad")%></span></b></font></td>
				</tr>
				<tr>
					<td colspan="2"><font face="verdana" size="10" color="red"><b>Contenido Destino:&nbsp;<%=rs2("comoviajad")%><b></font></td>
				</tr>

				<tr>
					<td colspan="2"><hr></td>
				</tr>
			</table>
			<br>
		<%end if%>
	<%end if%>
	<center>
	<font face="verdana" size="10" color="blue"><b><u>MODIFICACIÓN DE TRAMO</u></b></font><br><br>
	<form name="form1" action="modificarhr.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&hojaruta=<%=request.querystring("hr")%>&vdesde=<%=trim(request.querystring("desde"))%>&vhasta=<%=trim(request.querystring("hasta"))%>&tipo=<%=trim(request.querystring("tipo"))%>" method="POST">
		<%cade="select top(1) destinatario,cliente from hojaruta where numviaje=" & request.querystring("hr")  & " and tipo='" & request.querystring("tipo")& "' order by idhojaruta desc"
		set rs4=cn.execute(cade)
		oriori=rs4("destinatario")
		clicli=rs4("cliente")
		rs4.close%>
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
		<table style="width:70%" border=0>
			<tr>
				<td style="width:50%" align="right"><font face="verdana" size="12">Origen:</font></td>
					<td style="width:50%">
						<select name="desde" disabled>
							<option selected><%=request.querystring("desde")%></option>
						</select>
					</td>
					</tr>
					<tr>
						<td style="width:50%" align="right"><font face="verdana" size="6">Cant. Carga:&nbsp;<input type="number" min="0" style="width:110px;" maxlength="5" name="cargao" value="<%=request.querystring("cargao")%>" required></font></td>
						<td style="width:50%" align="right"><font face="verdana" size="6">Cant. Descarga:&nbsp;<input type="number" min="0" style="width:110px;" maxlength="5" name="descargao" value="<%=request.querystring("descargao")%>" required></font></td>
					</tr>

					<tr>
						<td align="center"><font face="verdana" size="6"><b><u>N° Remitos Carga</u></b></font></td>
						<td align="center"><font face="verdana" size="6"><b><u>N° Remitos Descarga</u></b></font></td>
					</tr>

					<tr>
						<td align="center"><font face="verdana" size="6"><input type="text" style="width:460px;" name="remitocargao" value="<%=request.querystring("remitocargao")%>"></font></td>
						<td align="center"><font face="verdana" size="6"><input type="text" style="width:460px;" name="remitodescargao" value="<%=request.querystring("remitodescargao")%>"></font></td>
					</tr>
					<tr>
						<td align="right"><font face="verdana" size="12">Contenido Origen:</font></td>
						<td>
							<select name="contenidoo">
								<option selected><%=contenidoo%></option>
								<option>PIEZAS</option>
								<option>PIEZAS Y RACKS</option>
								<option>RACKS</option>
								<option>VACÍO</option>
							</select>
						</td>
					</tr>
					<tr>
						<td colspan="2"><hr></td>
					</tr>
					<tr>
						<td style="width:50%" align="right"><font face="verdana" size="12">Destino:</font></td>
						<td style="width:50%">
							<select name="hasta" disabled>
							<option selected><%=request.querystring("hasta")%></option>
							</select>
						</td>
					</tr>
					<tr>
						<td style="width:50%" align="right"><font face="verdana" size="6">Cant. Carga:&nbsp;<input type="number" min="0" style="width:110px;" maxlength="5" name="cargad" value="<%=request.querystring("cargad")%>" required></font></td>
						<td style="width:50%" align="right"><font face="verdana" size="6">Cant. Descarga:&nbsp;<input type="number" min="0" style="width:110px;" maxlength="5" name="descargad" value="<%=request.querystring("descargad")%>" required></font></td>
					</tr>

					<tr>
						<td align="center"><font face="verdana" size="6"><b><u>N° Remitos Carga</u></b></font></td>
						<td align="center"><font face="verdana" size="6"><b><u>N° Remitos Descarga</u></b></font></td>
					</tr>

					<tr>
						<td align="center"><font face="verdana" size="6"><input type="text" style="width:460px;" name="remitocargad" value="<%=request.querystring("remitocargad")%>"></font></td>
						<td align="center"><font face="verdana" size="6"><input type="text" style="width:460px;" name="remitodescargad" value="<%=request.querystring("remitodescargad")%>"></font></td>
					</tr>

					<tr>
						<td align="right"><font face="verdana" size="12">Contenido Destino:</font></td>
						<td>
							<select name="contenidod">
								<option selected><%=contenidod%></option>
								<option>PIEZAS</option>
								<option>PIEZAS Y RACKS</option>
								<option>RACKS</option>
								<option>VACÍO</option>
							</select>
						</td>
					</tr>
				</table>
				<br>
				<center><input type="submit" class="boton_2" value="    Modificar Tramo    "></center>
			</form>
			<form name="form2" action="loginhr.html">
				<center><input type="submit" class="boton_2" value="Ir a la Página Principal"></center>
			</form>
			</center>
		</body>
	</html>
<%end if%>
</body>
</html>



