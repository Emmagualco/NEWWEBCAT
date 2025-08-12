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
	background: url(imagenruta.png) "CCC no-repeat";
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
   font-style: "Century Gothic";
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
font-size: 40px;
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
tipo=request.querystring("tipo")
numviaje=request.querystring("hr")
clave=request.querystring("clave")
dni=request.querystring("dni")
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)%>
<body background="fondo_ondulado.jpg">
	<table width="100%" bgcolor="white" border="0">
		<tr>
			<td width="20%"><img src="LogoCATNUEVO.gif"></td>
			<td width="60%">&nbsp;</td>
			<td width="20%" align="center"><a href="javascript:Abrir_iqnet('consultalc.asp?numviaje=<%=numviaje%>&clave=<%=clave%>&tipo=<%=tipo%>')"><img src="treslineas.png"width="70" height="50"></a></td>
		<tr>
	</table>
	<center>
	<br>
	<br>
	<font face="verdana" size="20" color="white"><b>CARGO EN&nbsp;<%=request.querystring("hasta")%></b></font>
	<br>
	<br>
	<br>
	<br>
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
		<%cade="select * from hojaruta where numviaje=" & rs1("numviaje") & " and horallegada is not null and horasalida is null order by idhojaruta"
		set rs2=cn.execute(cade)
		if rs2.eof=false then
			contenidod=trim(rs2("quellevaorigen"))%>
			<center>
			<%tipo=rs2("tipo")%>
			<%cliente=rs2("cliente")%>
			<%destinatario=trim(rs2("destinatario"))%>
		<%end if%>
	<%end if%>
	<center>
	

	<form name="form1" action="modificarhrcargaNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&hojaruta=<%=request.querystring("hr")%>&vdesde=<%=trim(request.querystring("desde"))%>&vhasta=<%=trim(request.querystring("hasta"))%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>" method="POST">
		<table style="width:80%" border=0>
			<tr>
				<td style="width:50%" align="right"><font face="verdana" size="6" color="white"><b>Cantidad Carga:</b></font></td>
				<td style="width:50%" align="left"><font face="verdana" size="6" color="white"><input type="number" min="1" style="width:110px;" maxlength="5" name="cargad" required></font></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>

			<tr>
				<td style="width:50%" align="right"><font face="verdana" size="6" color="white"><b>N� Remitos Carga:</b></font></td>
				<%if destinatario="MONTICH" then%>
					<td style="width:50%" align="left"><font face="verdana" size="6"><input type="text" style="width:460px;" name="remitocargad" required></font></td>
				<%else%>
					<td style="width:50%" align="left"><font face="verdana" size="6"><input type="text" style="width:460px;" name="remitocargad"></font></td>
				<%end if%>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>

			<tr>
				<td align="right"><font face="verdana" size="6" color="white"><b>Contenido:</b></font></td>
				<td>
					<select name="contenidod" required>
						<option>PIEZAS</option>
						<option>PIEZAS Y RACKS</option>
						<option>RACKS</option>
					</select>
				</td>
			</tr>
		</table>
		<br><br>
		<center><input type="submit" class="boton_2" value="    Confirmar Carga   "></center>
	</form>
	<form name="form2" action="listarhr1nva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>" method="POST"><br><br>
		<center><a href="listarhr1nva.asp?hr=<%=numviaje%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(tipo)%>&clave=<%=request.querystring("clave")%>"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
	</form>
	</center>
</body>
</html>