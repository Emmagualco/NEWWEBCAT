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
		<%cade="select * from hojaruta where numviaje=" & rs1("numviaje") & " and horallegada is not null and horasalida is null order by idhojaruta desc"
		set rs2=cn.execute(cade)
		if rs2.eof=false then
			
			contenidod=trim(rs2("quellevaorigen"))%>
			<center>
			<%tipo=rs2("tipo")%>
			<%cliente=rs2("cliente")%>
			<font face="verdana" size="10"><b>Tipo:&nbsp;<%=tipo%><b></font>
			<br>
			<font face="verdana" size="10"><b>CLIENTE:&nbsp;<%=cliente%><b></font>
			<br>
		<%end if%>
	<%end if%>
	<center>
<%quellevacarga=rs2("quellevacarga")%>
	<font face="verdana" size="12" color="blue">Descargo en&nbsp;<%=request.querystring("hasta")%></font>
	<form name="form1" action="modificarhrdescarga.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&hojaruta=<%=request.querystring("hr")%>&vdesde=<%=trim(request.querystring("desde"))%>&vhasta=<%=trim(request.querystring("hasta"))%>&tipo=<%=trim(request.querystring("tipo"))%>&quellevacarga=<%=quellevacarga%>" method="POST">
		<table style="width:80%" border=0>
			<tr>
				<td style="width:50%" align="right"><font face="verdana" size="6"><b>Cantidad Descarga:</b></font></td>
				<td style="width:50%" align="left"><font face="verdana" size="6"><input type="number" min="1" style="width:110px;" maxlength="5" name="descargad" value=<%=rs2("carga")%> required></font></td>
			</tr>
			<tr>
				<td style="width:50%" align="right"><font face="verdana" size="6"><b>N° Remitos Descarga:</b></font></td>
				<td style="width:50%" align="left"><font face="verdana" size="6"><input type="text" style="width:460px;" name="remitodescargad" value=<%=trim(rs2("remitocarga"))%>></font></td>
			</tr>
			<tr>
				<td align="right"><font face="verdana" size="6"><b>Contenido:</b></font></td>
				<td>
					<select name="contenidod" required disabled>
						<option>PIEZAS</option>
						<option>PIEZAS Y RACKS</option>
						<option>RACKS</option>
						<option selected><%=rs2("quellevacarga")%></option>
					</select>
				</td>
			</tr>
		</table>
		<br>
		<center><input type="submit" class="boton_2" value="   Confirmar Descarga  "></center>
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



