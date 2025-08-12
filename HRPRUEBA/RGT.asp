<%@Language="VBScript"%>
<%Response.Buffer = False%>
<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
dni=request.querystring("dni")
dim horas
dim todo
dim entero
dim partedec
dim horast 
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>  
	<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
  	<title>Consulta de Recorridos por Chofer</title>

<style>
img{
	border-radius: 50%;
}

/* Cajas con contenido de color */

.box-blue,
.box-gray,
.box-green,
.box-grey,
.box-red,
.box-yellow {
	margin:0 0 25px;
	overflow:hidden;
	padding:20px;
	-webkit-border-radius: 10px;
        border-radius: 30px;
}

.box-blue {
	background-color:#d8ecf7;
	border:1px solid #afcde3;
	height: 70px;
  	width: 950px;
	margin-left: 30px;
	margin-top: 0px;
}

.box-gray {
	background-color:#0288D1;
	border:1px solid #bdbdbd;
	height: 30px;
  	width: 950px;
	margin-left: 30px;
	margin-top: 0px;
}

.box-gray_chico {
        border-radius: 30px;
	background-color:#0288D1;
	border:1px solid #bdbdbd;
  	width: 980px;
	margin-left: 30px;
	margin-top: 0px;
}


.box-green {
	background-color:#d9edc2;
	border:1px solid #b2ce96;
}

.box-grey {
	background-color:#F5F5F5;
	border:1px solid #DDDDDD;
}

.box-red {
	background-color:#f9dbdb;
	border:1px solid #e9b3b3;
}

.box-yellow {
	background-color:#fef5c4;
	border:1px solid #fadf98;
}

#primero {
width: 100px;
float:left;
border: none ;

}
#segundo {
border: none;
overflow: hidden;
padding-top:2% !important;
padding-left: 3% !important;

}

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
    padding: 10px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 100;
    font-size: 12px;
    color: #ffffff;
    background-color:#071BE7;
    border-radius: 15px;
    border: 2px solid #0016b0;
  }
  .boton_2:hover{
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
   background-color:#EAE8E9;
   border: none;
   font-size: 20px;
   font-style: "Century Gothic";
   height: 20px;
   padding: 0px;
   width: 400px;
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

{
	document.body.style.zoom="80%"
}
</style>  	
</head>
	<body background="fondo_ondulado.jpg">
	<center>
	<div class="box-blue">
		<div id="primero"><img src="logoCATNUEVO.gif" width="80" heigth="80"></div>
		<div id="segundo"><big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span
		style="font-family: Helvetica,Arial,sans-serif;">Consulta de RGT</span></big></big><br></big></div>
	</div>
	<div class="box-gray">
		<table width="800px" border="0">
			<tr align="center">
				<%mes=month(request.querystring("desde"))
				if mes<10 then
					mes="0" & month(request.querystring("desde"))
				end if
				dia=day(request.querystring("desde"))
				if dia <10 then
					dia="0" & day (request.querystring("desde"))
				end if					
				anio=year(request.querystring("desde"))
				cadenadesde=dia & "/" & mes & "/" & anio
				mes=month(request.querystring("hasta"))
				if mes<10 then
					mes="0" & month(request.querystring("hasta"))
				end if
				dia=day(request.querystring("hasta"))
				if dia <10 then
					dia="0" & day (request.querystring("hasta"))
				end if					
				anio=year(request.querystring("hasta"))
				cadenahasta=dia & "/" & mes & "/" & anio%>
				<td style="text-align: center;"><big style="font-weight: bold;"><big style="color: rgb(255, 255, 255);"><big><span
				style="font-family: Helvetica,Arial,sans-serif;">Desde:&nbsp; <%=cadenadesde%>&nbsp;&nbsp;&nbsp;Hasta:&nbsp;<%=cadenahasta%></span></big></big></big></td>
			</tr>
		</table>
	</div>
	<center>
	<table border="1">
		<tr align="center">
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Fecha</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Dominio</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Tipo Cami�n</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Km Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Km MR</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Unidad</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Origen</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(248, 252, 158);">ML Tot</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">ML R</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Capac. R</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Carga R</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(75, 251, 248);">TLLR</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">ML N</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Capac. N</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Carga N</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(75, 251, 248);">TLLN</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(75, 251, 248);">Circuito</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Proveedor</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Tipo Flujo</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Cargo</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Observac.</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Incidente RASA</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Incidente NARG</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Otros RASA</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Otros NARG</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Rep. RASA</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Rep. UTE</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">ILN</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">No Prod.</td>

		</tr>
		<%totalhs=0%>

		<%cade="select * from rgtdef where fecha between '" & request.querystring("desde")  & "' and '" & request.querystring("hasta") & "' order by numviaje"
		set rs=cn.execute(cade)
		do while rs.eof=false%>
			<tr bgcolor="white" align="center">
				<td><font face="verdana" size="2"><%=rs("fecha")%></font></td>
				<td><font face="verdana" size="2"><%=rs("dominio")%></font></td>
				<td><font face="verdana" size="2"><%=rs("tipocamion")%></font></td>
				<td><font face="verdana" size="2"><%=rs("kmruta")%></font></td>
				<td><font face="verdana" size="2"><%=rs("kmmr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("unidad")%></font></td>
				<td><font face="verdana" size="2"><%=rs("origen")%></font></td>
				<td><font face="verdana" size="2"><%=rs("mltot")%></font></td>
				<td><font face="verdana" size="2"><%=rs("mlr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("capacr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("cargar")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("tllr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("mln")%></font></td>					
				<td><font face="verdana" size="2"><%=rs("capacn")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("cargan")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("tlln")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("circuito")%></font></td>
				<td><font face="verdana" size="2"><%=rs("proveedor")%></font></td>
				<td><font face="verdana" size="2"><%=rs("tipoflujo")%></font></td>
				<td><font face="verdana" size="2"><%=rs("cargo")%></font></td>
				<td><font face="verdana" size="2"><%=rs("observac")%></font></td>
				<td><font face="verdana" size="2"><%=rs("incidenterasa")%></font></td>
				<td><font face="verdana" size="2"><%=rs("incidentenarg")%></font></td>
				<td><font face="verdana" size="2"><%=rs("otrosrasa")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("otrosnarg")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("reprasa")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("repute")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("iln")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("noprod")%></font></td>
			</tr>
			<%rs.movenext
		loop%>
	</table>
<hr>
<center>
	<table>
		<tr>
			<td>
				<form name="form2" action="listadeconductores1.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
					<center><input type="submit" class="boton_2" value="  REGRESAR  "></center>
				</form>
			</td>
			<td>
				<form name="form2" action="excelrgt.asp?desde=<%=request.querystring("desde")%>&hasta=<%=request.querystring("hasta")%>&usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
					<center><input type="submit" class="boton_2" value="GENERAR EXCEL"></center>
				</form>
			</td>
		</tr>
	</table>
</center>
<hr>
	</body>
</html>