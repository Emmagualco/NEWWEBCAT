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
ceco=trim(request.form("ceco"))
ca="select c.rscliente from cliente c, cecos e where c.idcliente=e.idcliente and e.nombre='" & ceco & "'"
set rs4=cn.execute(ca)
cliente=trim(rs4("rscliente"))
rs4.close

cadena="select numviaje from hojaruta where idhojaruta=" & request.querystring("idhr")
set rs2=cn.execute(cadena)
numviaje=rs2("numviaje")
rs2.close

Const adOpenKeyset = 1
Const adLockOptimistic = 3
cadena="select * from viajes where numviaje=" & numviaje
set rs2=cn.execute(cadena)
if isnull(rs2("fechallegada")) then
	cade="select * from hojaruta where idhojaruta=" & request.querystring("idhr")
	set rs1=cn.execute(cade)
	numviaje=rs1("numviaje")
	remitente=rs1("destinatario")
	hojaderuta=trim(rs1("hojaderuta"))
	tipo=trim(rs1("tipo"))
	cliente=cliente
	cargad=rs1("cargadestino")
	remitocargad=trim(rs1("remitocargadestino"))
	comoviajas=trim(rs1("quellevadestino"))
	rs1.close
	rs.CursorType = adOpenKeyset
	rs.LockType = adLockOptimistic
	rs.open "hojaruta",cn
	rs.addnew
	rs("numviaje")=numviaje
	rs("remitente")=remitente
	rs("destinatario")=request.form("destino")
	rs("origen")="PFCBA"
	rs("destino")="PFCBA"
	rs("plataformaorigen")="CORDOBA"
	rs("plataformadestino")="CORDOBA"
	rs("hojaderuta")=hojaderuta
	rs("tipo")=tipo
	rs("cliente")=cliente
	rs("cargaorigen")=cargad
	rs("remitocargaorigen")=remitocargad
	rs("quellevaorigen")=comoviajas
	rs("cargadestino")=cargad
	rs("remitocargadestino")=remitocargad
	rs("quellevadestino")=comoviajas
	rs("ceco")=request.form("ceco")
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
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARGPRUEBA"
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
<%usuario=request.querystring("usuario")
clave=request.querystring("clave")%>
<body background="fondodepantalla.jpg">
<center>
<font face="verdana" size="6">Bienvenido <%=usuario%><br></font><br>
<hr>
<form name="form1" action="conductoresnuevotramo.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
	<%cad="select * from usuariosrasa where nombre='" & usuario & "' and clave='" & clave & "'" 
	set rs=cn.execute(cad)
	if rs.eof=false then%>
		<font face="verdana" size="6">Seleccione Origen:</font>
		<%idusuario=rs("idusuario")
		cad="select t.rstercero from usuariosrasa u, usuariosxsitio x, terceros t where t.IDTercero =x.idsitio and x.idusuario=u.idusuario and u.idusuario=" & idusuario & " and t.tipomov='" & "I" & "' order by t.RSTercero" 
		set rs1=cn.execute(cad)%>
		<select name="sitio">
			<%do while rs1.eof=false%>
				<option><%=rs1("rstercero")%></option>
				<%rs1.movenext
			loop%>
		</select>
		<br>
		<br>		
		<table>
			<tr>
				<td>
					<center><input type="submit" class="boton_3"  value="Aceptar"></center>
				</td>
				</form>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<form name="form5" action="nuevotramo.html">
				<td>
					<center><input type="submit" class="boton_3"  value="Regresar"></center>
				</td>
				</form>
			</tr>
		</table>
	<%else%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
		<head>  
			<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
			<title>Error de Ingreso</title>
				<link rel="stylesheet" href="estilohr.css">
			</head>
			<body background="fondodepantalla.jpg">
				<br>
				<br>
				<br>
				<br>
				<div class="box-green">
					<center><font face="verdana" size="6">No existe combinación Usuario/Password. Por favor, verifique los datos ingresados.</font></center>
				</div>
				<br>
				<br>
				<div>	
					<form name="form4" action="nuevotramo.html">
						<center><input type="submit" class="boton_3" value="Regresar"></center>
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
	<body background="fondodepantalla.jpg">
		<br>
		<br>
		<br>
		<br>
		<div class="box-green">
			<center><font face="verdana" size="12">La Hoja de Ruta se ha cerrado, ya no se puede generar un nuevo tramo</font></center>
		</div>
		<br>
		<br>
		<div>	
			<form name="form1" action="loginnuevotramo1.asp?sitio=<%=request.form("sitio")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
				<center><input type="submit" class="boton_2" value="Regresar"></center>
			</form>
		</div> 
	</body>
	</html>	
<%end if%>





	
