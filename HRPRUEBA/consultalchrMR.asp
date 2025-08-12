<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
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
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
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
  	width: 650px;
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

</style>  	
</head>

<body>
	<center>
	<div class="box-blue">
		<center><font face="verdana" size="5" color="#223697"><b>N&ordm; H.R.:&nbsp;<%=request.querystring("numviaje")%></b></font><br>
		<font face="verdana" size="5" color="#223697"><b>Consulta de Recorridos</b></font><br>
		<%cade="select fechasalida from viajes where numviaje=" & request.querystring("numviaje")
		set rs=cn.execute(cade)
		mes=month(rs("fechasalida"))
		if mes<10 then
			mes="0" & month(rs("fechasalida"))
		end if
		dia=day(rs("fechasalida"))
		if dia <10 then
			dia="0" & day (rs("fechasalida"))
		end if					
		anio=year(rs("fechasalida"))
		cadenadesde=dia & "/" & mes & "/" & anio
		desdecadena=mes & "/" & dia & "/" & anio%>		
		<font face="verdana" size="5" color="#223697"><b>Fecha H.R.:&nbsp; <%=cadenadesde%></b></font></center>
	</div>
	<table>
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Cliente</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Origen</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Destino</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Prevista Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Prevista Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Cont. Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Carga</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Descarga</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Cont. Salida</td>

		</tr>
		<%totalhs=0%>
		<%cade="select * from hojaruta where numviaje=" & request.querystring("numviaje") & " order by idhojaruta"
		set rs2=cn.execute(cade)
		if rs2.eof=false then
			color1=0
			color2=0
			do while rs2.eof=false%>
				<%if color1=0 then%>
					<tr bgcolor="#D5F5E3">
						<%color1=1
						color2=0
				else
					if color2 =0 then%>
						<tr bgcolor="#EAF2F8">
							<%color2=1
							color1=0
					end if
				end if%>
				<td style="font-family: Verdana; size:7"><%=rs2("cliente")%></td>
				<td style="font-family: Verdana;"><%=rs2("remitente")%></td>
				<td style="font-family: Verdana;"><%=rs2("destinatario")%></td>
				<td style="font-family: Verdana;"><center><%=rs2("horallegadaRASA")%></center></td>
				<td style="font-family: Verdana;"><center><%=rs2("horallegada")%></center></td>
				<td style="font-family: Verdana;"><center><%=rs2("horasalidaRASA")%></center></td>
				<td style="font-family: Verdana;"><center><%=rs2("horasalida")%></center></td>
					<%if rs2("cargaorigen")=0 then%>
						<td style="font-family: Verdana;"><center><%=rs2("quellevaorigen")%></center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs2("cargaorigen") & " " & trim(rs2("quellevaorigen"))%></center></td>
					<%end if%>
					<%if rs2("carga")=0 then%>
						<td style="font-family: Verdana;"><center>NO</center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs2("carga") & " " & trim(rs2("quellevacarga"))%></center></td>
					<%end if%>
					<%if rs2("descarga")=0 then%>
						<td style="font-family: Verdana;"><center>NO</center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs2("descarga") & " " & trim(rs2("quellevadescarga"))%></center></td>
					<%end if%>
					<%if rs2("cargadestino")=0 then%>
						<td style="font-family: Verdana;"><center><%=rs2("quellevadestino")%></center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs2("cargadestino") & " " & trim(rs2("quellevadestino"))%></center></td>
					<%end if%>
				</tr>
				<%rs2.movenext
			loop
		end if
		rs2.close%>
	</table>
	<br>
	<%cad="select * from hojarutamr where numviaje=" & request.querystring("numviaje") & " order by idhojaruta"
	set rs=cn.execute(cad)
	if rs.eof=false then%>
		
		<center>
		<font face="verdana" size="14" color="#223697"><b>Destinos Internos</b></font>
		<table>
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Cliente</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Origen</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Destino</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Cont. Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Carga</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Descarga</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Cont. Salida</td>

		</tr>
			<%color1=0
			color2=0
			do while rs.eof=false%>

				<%if color1=0 then%>
					<tr bgcolor="#D2D6FB">
						<%color1=1
						color2=0
				else
					if color2 =0 then%>
						<tr bgcolor="#EAF2F8">
							<%color2=1
							color1=0
					end if
				end if%>

					<td style="font-family: Verdana; size:7"><%=rs("cliente")%></td>
					<td style="font-family: Verdana;"><%=rs("remitente")%></td>
					<td style="font-family: Verdana;"><%=rs("destinatario")%></td>
					<td style="font-family: Verdana;"><center><%=rs("horallegada")%></center></td>
					<td style="font-family: Verdana;"><center><%=rs("horasalida")%></center></td>
					<%if rs("cargaorigen")=0 then%>
						<td style="font-family: Verdana;"><center><%=rs("quellevaorigen")%></center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs("cargaorigen") & " " & trim(rs("quellevaorigen"))%></center></td>
					<%end if%>
					<%if rs("carga")=0 then%>
						<td style="font-family: Verdana;"><center>NO</center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs("carga") & " " & trim(rs("quellevacarga"))%></center></td>
					<%end if%>
					<%if rs("descarga")=0 then%>
						<td style="font-family: Verdana;"><center>NO</center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs("descarga") & " " & trim(rs("quellevadescarga"))%></center></td>
					<%end if%>
					<%if rs("cargadestino")=0 then%>
						<td style="font-family: Verdana;"><center><%=rs("quellevadestino")%></center></td>
					<%else%>
						<td style="font-family: Verdana;"><center><%=rs("cargadestino") & " " & trim(rs("quellevadestino"))%></center></td>
					<%end if%>
				</tr>
				<%rs.movenext
			loop%>
			</tr>
		</table>
		</center>
	
	<%end if%>
	<hr>
  	<p><font color="#993300" size="7" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:close()"><strong>Cerrar 
    	Pantalla</strong></a></font></p>
	</center>
	</body>

</html>