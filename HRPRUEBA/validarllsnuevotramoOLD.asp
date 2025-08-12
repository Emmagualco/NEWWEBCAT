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
    padding: 5px;
    font-family: arial;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 50;
    font-size: 15px;
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
    font-weight: 100;
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
font-size: 13px;
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
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)%>
<body background="fondodepantalla.jpg">
<center>
<font face="verdana" size="6">Bienvenido <%=request.querystring("usuario")%></font>
<hr>
<font face="verdana" size="6">Origen: <%=request.querystring("sitio")%></font>
<hr>
<font face="verdana" size="6">Conductor: <%=request.form("conductor")%></font>
<%conductor=request.form("conductor")%>
<hr>
<%cad="select h.numviaje,h.fechasalida,h.horasalida,h.fechallegada,h.horallegada,h.fechasalidarasa,h.horasalidarasa,h.fechallegadarasa,h.fechasalidarasa from conductores c, viajes v, Hojaruta h where h.numviaje=v.numviaje and v.idconductor=c.idconductor and h.destinatario ='" & request.querystring("sitio") & "'  and h.FechaSalidaRASA is null and h.HoraSalidaRASA is null and h.fechallegadarasa is null and h.horallegadarasa is null order by h.Numviaje desc"
set rs=cn.execute(cad)
if rs.eof=false then%>
<font face="verdana">
<form name="form1" action="validacionlls.asp?sitio=<%=request.querystring("sitio")%>&usuario=<%=request.querystring("usuario")%>&conductor=<%=conductor%>&clave=<%=request.querystring("clave")%>" method="POST">
<table style="text-align: left; width: 147px; height: 208px;" border="0">
	<tbody>	
	<font face="verdana">
	<tr style="background-color: rgb(255, 128, 51);">	
		<td colspan="4" VALIGN="MIDDLE" ALIGN="CENTER"><b>VALIDACIÓN</b></td>
	</tr>
	<tr style="background-color: rgb(51, 138, 255);">	
		<td colspan="4" VALIGN="MIDDLE" ALIGN="CENTER"><b>Registrado por el Conductor</b></td>
	</tr>
	<tr style="background-color: rgb(51, 138, 255);">
		<TD colspan="2" VALIGN="MIDDLE" ALIGN="CENTER"><b>LLEGADA</b></td>
		<TD colspan="2" VALIGN="MIDDLE" ALIGN="CENTER"><b>SALIDA</b></td>
	</tr>
	<tr style="background-color: rgb(51, 138, 255);">
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Fecha:</b></td>
		<TD VALIGN="MIDDLE" ALIGN="CENTER"><%=rs("fechallegada")%></td>
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Fecha:</b></td>
		<TD VALIGN="MIDDLE" ALIGN="CENTER"><%=rs("fechasalida")%></td>
	</tr>
	<tr style="background-color: rgb(51, 138, 255);">
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Hora:</b></td>
		<TD VALIGN="MIDDLE" ALIGN="CENTER"><%=rs("horallegada")%></td>
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Hora:</b></td>
		<TD VALIGN="MIDDLE" ALIGN="CENTER"><%=rs("horasalida")%></td>
	</tr>
	<tr style="background-color: rgb(47, 214, 73);">	
		<td colspan="4" VALIGN="MIDDLE" ALIGN="CENTER"><b>Validar por RASA/NISSAN</b></td>
	</tr>
	<tr style="background-color: rgb(47, 214, 73);">
		<TD colspan="2" VALIGN="MIDDLE" ALIGN="CENTER"><b>LLEGADA</b></td>
		<TD colspan="2" VALIGN="MIDDLE" ALIGN="CENTER"><b>SALIDA</b></td>
	</tr>
	<tr style="background-color: rgb(47, 214, 73);">
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Fecha:</b></td>
				<%mes=month(rs("fechallegada"))%>
				<%if mes<10 then
					mes="0" & month(rs("fechallegada"))
				end if%>
				<%dia=day(rs("fechallegada"))%>
				<%if dia <10 then
					dia="0" & day (rs("fechallegada"))
				end if%>					
				<%anio=year(rs("fechallegada"))%>
				<%cadenahoy=anio & "-" & mes & "-" & dia%>
				<%if mes<10 then
					mes="0"& mes
				end if%>
				<%if dia<10 then
					dia="0"& dia
				end if%>
				<%cade=anio & "-" & mes & "-" & dia%>
		<TD VALIGN="MIDDLE" ALIGN="CENTER"><input name="fechallegada" step="1" type="date" required min="<%=cadenahoy%>" max="<%=cadenahoy%>" value="<%=cadenahoy%>"></td>
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Fecha:</b></td>
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
				<%if mes<10 then
					mes="0"& mes
				end if%>
				<%if dia<10 then
					dia="0"& dia
				end if%>
				<%cade=anio & "-" & mes & "-" & dia%>

		<TD VALIGN="MIDDLE" ALIGN="CENTER"><input name="fechasalida" step="1" type="date" required min="<%=cadenahoy%>" max="<%=cadenahoy%>" value="<%=cadenahoy%>"></td>
	</tr>
	<tr style="background-color: rgb(47, 214, 73);">
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Hora:</b></td>
				<%horaserver=rs("horallegada")%>
				<%hora=hour(horaserver)%>
				<%if hora<10 then
					hora="0" & hour(horaserver)
				end if%>
				<%minu=minute(horaserver)%>
				<%if minu<10 then
					minu="0" & minute(horaserver)
				end if%>
				<%cadhora=hora & ":" & minu%>
				<%ampm=(right(horaserver, 2))%>	
				<%if hour(horaserver)>=12 then
					dia1=day(now()+2)
					if dia1<10 then
						dia1="0"& dia1
					end if

					mes1=month(now()+2)
					if mes1<10 then
						mes1="0"& mes1
					end if
					anio1=year(now()+2)
					caden=anio1 & "-" & mes1 & "-" & dia1
				else
					dia1=day(now()+1)
					if dia1<10 then
						dia1="0"& dia1
					end if
					mes1=month(now()+1)
					if mes1<10 then
						mes1="0"& mes1
					end if
					anio1=year(now()+1)
					caden=anio1 & "-" & mes1 & "-" & dia1
				end if%>

				<%horaserver=time()%>
				<%hora=hour(horaserver)%>
				<%if hora<10 then
					hora="0" & hour(horaserver)
				end if%>
				<%minu=minute(horaserver)%>
				<%if minu<10 then
					minu="0" & minute(horaserver)
				end if%>
				<%cadhorahoy=hora & ":" & minu%>
				<%ampm=(right(horaserver, 2))%>	
				<%if hour(horaserver)>=12 then
					dia1=day(now()+2)
					if dia1<10 then
						dia1="0"& dia1
					end if

					mes1=month(now()+2)
					if mes1<10 then
						mes1="0"& mes1
					end if
					anio1=year(now()+2)
					caden=anio1 & "-" & mes1 & "-" & dia1
				else
					dia1=day(now()+1)
					if dia1<10 then
						dia1="0"& dia1
					end if
					mes1=month(now()+1)
					if mes1<10 then
						mes1="0"& mes1
					end if
					anio1=year(now()+1)
					caden=anio1 & "-" & mes1 & "-" & dia1
				end if%>



		<TD VALIGN="MIDDLE" ALIGN="LEFT"><input name="hora1" type="time" max="<%=cadhorahoy%>"  value="<%=cadhora%>"/></td>
		<TD VALIGN="MIDDLE" ALIGN="RIGHT"><b>Hora:</b></td>
				<%horaserver=time()%>
				<%hora=hour(horaserver)%>
				<%if hora<10 then
					hora="0" & hour(horaserver)
				end if%>
				<%minu=minute(horaserver)%>
				<%if minu<10 then
					minu="0" & minute(horaserver)
				end if%>
				<%cadhora=hora & ":" & minu%>
				<%ampm=(right(horaserver, 2))%>	
				<%if hour(horaserver)>=12 then
					dia1=day(now()+2)
					if dia1<10 then
						dia1="0"& dia1
					end if

					mes1=month(now()+2)
					if mes1<10 then
						mes1="0"& mes1
					end if
					anio1=year(now()+2)
					caden=anio1 & "-" & mes1 & "-" & dia1
				else
					dia1=day(now()+1)
					if dia1<10 then
						dia1="0"& dia1
					end if
					mes1=month(now()+1)
					if mes1<10 then
						mes1="0"& mes1
					end if
					anio1=year(now()+1)
					caden=anio1 & "-" & mes1 & "-" & dia1
				end if%>


				<%horaserver=rs("horallegada")%>
				<%hora=hour(horaserver)%>
				<%if hora<10 then
					hora="0" & hour(horaserver)
				end if%>
				<%minu=minute(horaserver)%>
				<%if minu<10 then
					minu="0" & minute(horaserver)
				end if%>
				<%cadhorall=hora & ":" & minu%>


		<TD VALIGN="MIDDLE" ALIGN="LEFT"><input name="hora2" type="time"  min="<%=cadhorall%>" max="<%=cadhora%>"  value="<%=cadhora%>"/></td>
	</tr>
	<tr>
		<td colspan="4"><b><u>Observaciones</u></b></td>
	<tr>
	<tr>
		<td colspan="4"><font face="verdana" size="1"><input type="text" name="obser" style="width: 400px;" maxlength="500" required></font></td>
	</tr>
	
		<tr>
			<td colspan="2">
				
					<center><input type="submit" class="boton_2" value="Validar TODO"></center>
		</form>
			</td>
			<form name="form7" action="conductoresnuevotramo1.asp?sitio=<%=request.querystring("sitio")%>&usuario=<%=request.querystring("usuario")%>&conductor=<%=conductor%>&clave=<%=request.querystring("clave")%>" method="POST">
			<td colspan="2">
					<center><input type="submit" class="boton_2" value="REGRESAR"></center>
				</form>
			</td>
		</tr>
</tbody>
</table>
</font>
<%end if%>







