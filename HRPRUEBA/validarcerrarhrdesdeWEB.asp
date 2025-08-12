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
<body bgcolor="#9091A0">
<table style="text-align: left; width: 100%;" border="0">
	<tbody>
		<tr bgcolor="white">
			<td colspan="4" rowspan="1" style="vertical-align: middle;"><img style="width: 118px; height: 118px;" alt="" src="LogoCATNUEVO.gif"></td>
			<td><big style="font-weight: bold; color: rgb(0, 0, 153);"><font face="verdana" size="12">CERRAR HOJA DE RUTA</big></font></td>
		</tr>
</tbody>

<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
set rs6=server.createobject("ADODB.Recordset")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
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
<%usuario=request.querystring("usuario")%>
<%clave=request.querystring("clave")
encontrado=0
cad="select * from hojaruta where numviaje=" & request.querystring("numviaje") & " order by idhojaruta desc"
set rs=cn.execute(cad)
if trim(rs("quellevadestino"))<>"VACÍO" then%>
	</table>
	<br>
	<br>
	<br>
	<br>
	<br>
	<center><font face="verdana" size="14" color="white"><b>ESTA HOJA DE RUTA TIENE CARGA.<br>
	DESEA CERRARLA?</font></b>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
	<table style="text-align: left; width: 100%;" border="0">
		<tr>
		<td align="right">
			<form name="form1" action="cerrarhrdesdeWEB.asp?numviaje=<%=request.querystring("numviaje")%>&usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
			<a href="cerrarhrdesdeWEB.asp?numviaje=<%=request.querystring("numviaje")%>&usuario=<%=usuario%>&clave=<%=clave%>"><img style="width: 225px; height: 225px;" alt="" src="aceptar.jpg"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</form>
		</td>
		<td align="left">
			<form name="form1" action="listadeconductores.asp?usuario=<%=usuario%>&clave=<%=clave%>" method="POST">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="listadeconductores.asp?usuario=<%=usuario%>&clave=<%=clave%>"><img style="width: 225px; height: 225px;" alt="" src="cancelar.jpg"></a>
			</form>
		</td>
		</tr>
	</table>
<%else%>
	<%cade="select * from hojaruta where numviaje=" & request.querystring("numviaje") & " order by idhojaruta desc"
	set rs5=cn.execute(cade)
	fechallegada=rs5("fechasalidarasa")
	horallegada=rs5("horasalidarasa")
	rs1.CursorType = adOpenKeyset
	rs1.LockType = adLockOptimistic
	rs1.open"select * from hojaruta where numviaje=" & request.querystring("numviaje") & " order by idhojaruta desc",cn
	rs1("quellevadestino")="VACÍO"
	rs1("cargadestino")=0
	if rs1("carga")>0 then
		rs1("descarga")=rs1("carga")
		rs1("quellevadescarga")=rs1("quellevacarga")
		rs1("remitodescarga")=trim(rs1("remitocarga"))
	else
		if rs1("cargaorigen")>0 then
			rs1("descarga")=rs1("cargaorigen")
			rs1("quellevadescarga")=rs1("quellevaorigen")
			rs1("remitodescarga")=trim(rs1("remitocargaorigen"))
		end if		
	end if
	rs1.update
	rs1.close
	rs6.CursorType = adOpenKeyset
	rs6.LockType = adLockOptimistic
	rs6.open "Select * from viajes where numviaje=" & request.querystring("numviaje"),cn
	rs6("fechallegada")=fechallegada
	rs6("horallegada")=trim(horallegada)
	rs6("etapa")="C"
	rs6.update
	rs6.close
	response.redirect("listadeconductores.asp" & "?usuario="& usuario & "&clave=" & clave)%> 
<%end if%>
</body>
</html>