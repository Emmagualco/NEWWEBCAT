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
    font-size: 15px;
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


  .boton_5{
    text-decoration: none;
    padding: 5px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 50;
    font-size: 10px;
    color: #1C2833;
    background-color:#D4EFDF;
    border-radius: 10px;
    border: 2px solid #27AE60;
  }
  .boton_5:hover{
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
   font-size: 20px;
   font-face: "Century Gothic"
   height: 30px;
   padding: 0px;
   width: 300px;
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

<%mes=month(now())%>
<%if mes<10 then
	mes="0" & month(now())
end if%>
<%dia=day(now())%>
<%if dia <10 then
	dia="0" & day (now())
end if%>					
<%anio=year(now())%>
<%cadenahoy=anio & "-" & mes & "-" & dia%>

<%usuario=trim(request("usuario"))%>
<%clave=trim(request("clave"))%>
<center>
<%cad="select * from proveedores where usuarioweb='" & usuario & "'and clave='" & clave & "'"
set rs=cn.execute(cad)
if rs.eof=false then%>
	<%idproveedor=rs("idproveedor")%>
	<body background="fondo_ondulado.jpg">
	<table width="100%" border="0">
		<tr>
			<td width="20%"><img src="logocatnuevo.gif" width="100" height="100"></td>
			<td width="80%"><font face="verdana" size="6" color="white">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Bienvenido <%=request("usuario")%></font></td>
		</tr>
	</table>
	<hr>
	<form name="form2" action="consultarecorridosproveedores.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="post">
	<table border="1" width="100%">
		<tr bgcolor="white">
			<td align="right" width="50%"><font face="verdana" size="5">Seleccione Conductor:</font></td>
		<td align="left" width="50%">
			<select name="conductores">
				<option>TODOS</option>
				<%cade="select nombre from conductores where idproveedor=" & idproveedor & " and estado='" & "A" & "' order by nombre"
				set rs1=cn.execute (cade)
				do while rs1.eof=false%>
					<option><%=rs1("nombre")%></option>
					<%rs1.movenext
				loop%>
			</select>
		</td>
		<tr bgcolor="white">
			<td colspan="2" align="center">
				<font face="verdana"  size="5">Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenahoy%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenahoy%>" required></font>
				<input value="   Consultar   " type="submit" class="boton_3"><br>
			</td>
		</tr>
	</table>
	</form>
<br>
<br>
<br>
<br>
<br>
<br>
	<center>
	<form name="form3" action="proveedores.html">
		<button class="boton_2" onclick="location.href='proveedores.html'">REGRESAR</button>
	</form>
	

<%else%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>  
	<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
	<title>Sin Conductores para Validación</title>
	</head>
	<body background="fondodepantalla.jpg">
		<br>
		<br>
		<br>
<%=request("usuario")%>
<%=request("clave")%>
		<div class="box-green">
			<center><font face="verdana" size="12">No existe conmbinación Usuario/Clave.</font></center>
		</div>
		<br>
		<br>
		<br>
		<table>
			<tr>
				<td>
					<div>	
						<form name="form3" action="proveedores.html">
							<td align="center"><center><button class="boton_2" onclick="location.href='proveedores.html'">REGRESAR</button></center>
						</form>
					</div> 
				</td>
			<tr>
		</table>
	</body>
	</html>	
<%end if%>
