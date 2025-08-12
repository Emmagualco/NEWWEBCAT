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
	<meta http-equiv="refresh" content="30">

	<script language="JavaScript">
	function Abrir_iqnet (pagina) {
	var opciones="toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, width=800, height=255, top=85, left=200";
	window.open(pagina,"",opciones);
	}
</script>
  	<title>Consulta de Hojas de Ruta Abiertas</title>

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
	height: 100px;
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
   font-face: "Century Gothic"
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

</style>  	
</head>
<body background="fondo_ondulado.jpg">
	<center>
	<div class="box-blue">
		<div id="primero"><img src="logoCATNUEVO.gif" width="80" heigth="80"></div>
		<div id="segundo"><big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span
		style="font-family: Helvetica,Arial,sans-serif;">Hojas de Ruta Abiertas - MILK RUN</span></big></big><br></big></div>
		<center><font face="verdana" size="4"><b>Fecha:&nbsp;<%=date()%>&nbsp;Hora:&nbsp;<%=time()%></b></font></center>
	</div>
	<center>
	<table width="60%">
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3><center>N° H.R.</center></h3></td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3>Fecha H.R.</h3></td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3>Conductor</h3></td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3>Patente</h3></td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3><center>Ubicación Actual</center></h3></td>	
		</tr>
		<%a=0
		g=0
		viajes=0
		hs=0
		cerrar=0
		cade="select v.numviaje,v.fechasalida,m.patente,c.nombre from viajes v, conductores c, medios m where v.idmedio=m.idmedio and v.idconductor=c.idconductor and  v.etapa='" & "A" & "' and v.tipomov='" & "M" & "' and shipment is null order by c.nombre"
		set rs4=cn.execute(cade)
		do while rs4.eof=false
			cad="select * from hojaruta where numviaje=" & rs4("numviaje") & " order by idhojaruta desc"
			set rs1=cn.execute(cad)
			if not isnull(rs1("horallegada")) and isnull(rs1("horasalida")) then		
				ubicacion="LLEGADA A " & trim(rs1("destinatario")) & " " & trim(rs1("horallegada"))
 				horaserver=rs1("horallegada")
				hora=hour(horaserver)
				if hora<10 then
					hora="0" & hour(horaserver)
				end if
				minu=minute(horaserver)
				if minu<10 then
					minu="0" & minute(horaserver)
				end if
				cadhorall=hora & ":" & minu

				horaserver=time()
				hora=hour(horaserver)
				if hora<10 then
					hora="0" & hour(horaserver)
				end if
				minu=minute(horaserver)
				if minu<10 then
					minu="0" & minute(horaserver)
				end if
				cadhorahoy=hora & ":" & minu
				hs = DateDiff("n", cadhorall, cadhorahoy)
				if rs1("fechallegada")<>date then
					hs=hs+1440
				end if
			end if
			if not isnull(rs1("horallegada")) and not isnull(rs1("horasalida")) then		
				ubicacion="SALIDA DE " & trim(rs1("destinatario")) & " " & trim(rs1("horasalida"))
				cerrar=1
			end if
			if isnull(rs1("horallegada")) and isnull(rs1("horasalida")) then		
				ubicacion="EN CAMINO A " & trim(rs1("destinatario")) & " DESDE " & trim(rs1("remitente"))
			end if
			cad="select * from hojaruta where numviaje=" & rs4("numviaje") & " order by idhojaruta desc"
			set rs3=cn.execute(cad)
			if hs<60 then%>
				<tr bgcolor="#49F66B">
					<td style="font-family: Verdana;"><h4><center><a href="javascript:Abrir_iqnet('consultalchr.asp?numviaje=<%=rs4("numviaje")%>')"><%=rs4("numviaje")%></a></center></h4></td>
					<td style="font-family: Verdana;"><h4><%=rs4("fechasalida")%></h4></td>
					<td style="font-family: Verdana;"><h4><%=rs4("nombre")%></h4></td>
					<td style="font-family: Verdana;"><h4><%=rs4("patente")%></h4></td>
					<td style="font-family: Verdana;"><h4><%=ubicacion%></h4></td>
				</tr>
			<%else%>
				<tr bgcolor="#FD0316">
					<td style="font-family: Verdana;"><h4><center><a href="javascript:Abrir_iqnet('consultalchr.asp?numviaje=<%=rs4("numviaje")%>')"><%=rs4("numviaje")%></a></center></h4></td>
					<td style="font-family: Verdana; color:#ffffff;"><h4><%=rs4("fechasalida")%></h4></td>
					<td style="font-family: Verdana; color:#ffffff;"><h4><%=rs4("nombre")%></h4></td>
					<td style="font-family: Verdana; color:#ffffff;"><h4><%=rs4("patente")%></h4></td>
					<td style="font-family: Verdana; color:#ffffff;"><h4><%=ubicacion%></h4></td>
				</tr>
			<%end if				
			viajes=viajes + 1
			hs=0
			cerrar=0
			rs4.movenext
		loop%>
		<tr>
			<td colspan="6">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="6"><hr></td>
		</tr>
		<tr>
			<td colspan="6" align="center"><font face="verdana" size="5" color="white"><b>Total de H.R. Abiertas:&nbsp;<%=viajes%></b></font></td>
		</tr>
		<tr>
			<td colspan="6"><hr></td>
		</tr>
		<tr>
			<td colspan="6">&nbsp;</td>
		</tr>
	</table>
	<form name="form3" action="listadeconductores1.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
		<center><input type="submit" class="boton_2" value="  REGRESAR  "></center>
	</form>	
	</center>
</body>
</html>
