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
			<td><big style="font-weight: bold; color: rgb(0, 0, 153);"><font face="verdana" size="12">SIN REGISTRO DE CARGA Y DESCARGA</big></font></td>
		</tr>
</tbody>
</table>
<br>
<br>
<br>
<br>
<br>
<center><font face="verdana" size="14" color="white"><b>NO HA REGISTRADO<br>
CARGA NI DESCARGA<br>
EN ESTE PUNTO.<br><br>
CONFIRMA SALIDA?</font></b>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<%obser=ucase(trim(request.form("obser")))%>
<table style="text-align: left; width: 100%;" border="0">
	<tr>
	<td align="right">
		<form name="form1" action="registrosalidaNVAMR.asp?hr=<%=request.querystring("hr")%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>" method="POST">
		<a href="registrosalidaNVAMR.asp?obser=<%=obser%>&hr=<%=request.querystring("hr")%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>"><img style="width: 225px; height: 225px;" alt="" src="aceptar.jpg"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</form>
	</td>
	<td align="left">
		<form name="form1" action="listarhr1NVA.asp?hr=<%=request.querystring("hr")%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>" method="POST">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="listarhr1NVA.asp?hr=<%=request.querystring("hr")%>&dni=<%=request.querystring("dni")%>&tipo=<%=trim(request.querystring("tipo"))%>&clave=<%=request.querystring("clave")%>"><img style="width: 225px; height: 225px;" alt="" src="cancelar.jpg"></a>
		</form>
	</td>
	</tr>
</table>
</body>
</html>