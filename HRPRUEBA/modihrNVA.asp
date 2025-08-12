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
    font-size: 45px;
    color: #ffffff;
    background-color: #F5071D;
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

<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
tipo=request.querystring("tipo")
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
numviaje=request.querystring("hr")%>
<body background="fondo_ondulado.jpg">
<center>
<table width="100%" bgcolor="white" border="0">
	<tr>
		<td width="20%"><img src="LogoCATNUEVO.gif"></td>
		<td width="60%">&nbsp;</td>
		<td width="20%" align="center"><a href="javascript:Abrir_iqnet('consultalc.asp?numviaje=<%=numviaje%>&clave=<%=clave%>&tipo=<%=tipo%>')"><img src="treslineas.png"width="70" height="50"></a></td>
	<tr>
</table>
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
	<%cade="select * from hojaruta where numviaje=" & rs1("numviaje") & " and (horallegada is null or horasalida is null) order by idhojaruta desc"
	set rs2=cn.execute(cade)
	if rs2.eof=false then%>
		<center>
		<%tipo=rs2("tipo")%>
		<br>
	<%end if%>
<%end if%>
<center>
<font face="verdana" size="10" color="white"><b><u>MODIFICACIÓN DE DESTINO</u></b></font><br><br><br>
<form name="form1" action="modificarhrNVA.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&hojaruta=<%=request.querystring("hr")%>&vdesde=<%=trim(request.querystring("desde"))%>&vhasta=<%=trim(request.querystring("hasta"))%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>" method="POST">
	<%cade="select top(1) destinatario,remitente,cliente from hojaruta where numviaje=" & request.querystring("hr")  & " and tipo='" & request.querystring("tipo")& "' order by idhojaruta desc"
	set rs4=cn.execute(cade)
	oriori=rs4("remitente")
	destidesti=rs4("destinatario")
	clicli=rs4("cliente")
	rs4.close%>
	<%cade1="select rscliente from cliente where hrweb='" & "S" & "' order by rscliente"
	set rs4=cn.execute(cade1)%>
	<table style="width:80%" border=0>
		<tr>
			<td style="width:50%" align="right"><font face="verdana" size="12" color="white">DESTINO ACTUAL:&nbsp;</font></td>
			<td style="width:50%"><font face="verdana" size="12" color="white"><%=destidesti%></font></td>
		</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
		</tr>
		<tr>
			<td style="width:50%" align="right"><font face="verdana" size="12" color="white">NUEVO DESTINO:&nbsp;</font></td>
			<td style="width:50%">
				<%cad="select rstercero from terceros where estado='" & "A" & "' and tipomov='" & "I" & "' order by rstercero"
				set rs4=cn.execute(cad)%>
				<select name="hasta">
					<%do while rs4.eof=false
						if trim(rs4("rstercero"))<>trim(destidesti) then%>
							<option><%=rs4("rstercero")%></option>
						<%end if
						rs4.movenext
					loop
					rs4.close%>
				</select>
			</td>
		</tr>
	</table>
	<br>
	<br>
	<br>
	<center><input type="submit" class="boton_2" value="   MODIFICAR DESTINO   "></center>
</form>
<br>
<br>
<br>
<form name="form2" action="listarhr1nva.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&hojaruta=<%=request.querystring("hr")%>&vdesde=<%=trim(request.querystring("desde"))%>&vhasta=<%=trim(request.querystring("hasta"))%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>" method="POST">
	<center><a href="listarhr1nva.asp?hr=<%=rs1("numviaje")%>&dni=<%=request.querystring("dni")%>&hojaruta=<%=request.querystring("hr")%>&vdesde=<%=trim(request.querystring("desde"))%>&vhasta=<%=trim(request.querystring("hasta"))%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>"><img style="width: 150px; height: 150px;" src="flechaatras.jpg"></a></center>
</form>
</center>
</body>
</html>


