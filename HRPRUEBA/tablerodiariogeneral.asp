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
movimt=request("movimt")
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
  	<title>Consulta de Tablero Diario GENERAL</title>

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
<%if request("movimt")="MILK RUN" then%>
<body background="fondo_ondulado.jpg">
	<center>
	<div class="box-blue">
		<div id="primero"><img src="logoCATNUEVO.gif" width="80" heigth="80"></div>
		<div id="segundo"><big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span
		style="font-family: Helvetica,Arial,sans-serif;">Tablero Diario</span></big></big><br></big></div>
		<center><font face="verdana" size="4"><b>Fecha:&nbsp;<%=date()%>&nbsp;Hora:&nbsp;<%=time()%></b></font></center>
	</div>
	<center>
	<font face="verdana" size="14" color="white"><b>MILK RUN</b></font>
	<table width="70%">
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(195, 201, 249);"><h3>Conductor</h3></td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(195, 201, 249);"><h3><center>Patente</center></h3></td>		
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(195, 201, 249);"><h3><center>N� H.R.</center></h3></td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(195, 201, 249);"><h3><center>Fecha H.R.</center></h3></td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(195, 201, 249);"><h3><center>Estado Actual</center></h3></td>	
		</tr>
		<%a=0
		g=0
		viajes=0
		cade="select m.patente, v.etapa, c.nombre as nombrecond, p.RSProveedor as nombreprov, v.NumViaje, v.FechaSalida from viajes v, Conductores c, proveedores p, medios m where v.idmedio=m.idmedio and v.IDConductor =c.IDConductor and v.CodProveedor=p.IDProveedor and (v.etapa='" & "A" & "' or v.etapa='" & "G" & "') and v.tipomov='" & "M" & "' order by c.Nombre,v.numviaje"
		set rs4=cn.execute(cade)
		do while rs4.eof=false
    			If RS4("etapa") = "A" Then
        			a = a + 1
        			demora = 0
				tramo=0
			        Mensaje = ""
			        cad="select * from hojaruta where numviaje=" & RS4("numviaje") & " order by idhojaruta"
				set rs1=cn.execute(cad)
			        Do While RS1.EOF = False
			        	tramo = tramo + 1
            				If tramo = 1 Then
				                If Not IsNull(RS1("horallegada")) And Not IsNull(RS1("horasalida")) Then
					        	Mensaje = "Sali� de " & RS1("destinatario") &  " a las " & rs1("horasalida")
                    					vsalida = 1  
							ahora=hour(now)& ":" & minute(now)
							salida=left(rs1("horasalida"),5)
							demora = DateDiff("n", salida, ahora) / 60

                				End If
                				If Not IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) Then
					        	Mensaje = "Lleg� a " & RS1("destinatario") &  " a las " & rs1("horallegada")
                    					vllegada = 1
							ahora=hour(now)& ":" & minute(now)
							llegada=left(rs1("horallegada"),5)
							demora = DateDiff("n", llegada, ahora) / 60
                				End If
                				If IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) Then
                    					Mensaje = "En camino a " & RS1("destinatario")
                    					nada = 1
                				End If
            				End If
            				If tramo = 2 Then
                				If Not IsNull(RS1("horallegada")) And Not IsNull(RS1("horasalida")) Then
                    					Mensaje = "Sali� de " & RS1("destinatario") &  " a las " & rs1("horasalida")
                    					vsalida = 1
							ahora=hour(now)& ":" & minute(now)
							salida=left(rs1("horasalida"),5)
							demora = DateDiff("n", salida, ahora) / 60
							caden="select * from hojarutamr where numviaje=" & rs4("numviaje")
							set rs5=cn.execute(caden)
							if rs5.eof=false then
								do while rs5.eof=false
									if not isnull(rs5("horallegada")) and not isnull(rs5("horasalida")) then
										mensaje="Sali� " & rs5("destinatario") &  " a las " & rs5("horallegada")
									end if
									if not isnull(rs5("horallegada")) and isnull(rs5("horasalida")) then
										mensaje="Lleg� a " & rs5("destinatario") &  " a las " & rs5("horallegada")
										ahora=hour(now)& ":" & minute(now)
										llegada=left(rs5("horallegada"),5)
										demora = DateDiff("n", llegada, ahora) / 60
									end if
									rs5.movenext
								loop
							end if
                				End If
                				If Not IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) Then
                    					Mensaje = "Lleg� a " & RS1("destinatario") &  " a las " & rs1("horallegada")
                    					vsalida = 0
                    					vllegada = 1
							ahora=hour(now)& ":" & minute(now)
							llegada=left(rs1("horallegada"),5)
							demora = DateDiff("n", llegada, ahora) / 60

                				End If
                				If IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) And vsalida = 1 Then
                    					Mensaje = "En camino a " & RS1("destinatario")
                    					vsalida = 0
                				Else
                    					nada = 1
                				End If
            				End If
            				If tramo = 3 Then
                				If Not IsNull(RS1("horallegada")) And Not IsNull(RS1("horasalida")) Then
                    					Mensaje = "Sali� de " & RS1("destinatario") &  " a las " & rs1("horasalida")
                    					vsalida = 1
							ahora=hour(now)& ":" & minute(now)
							salida=left(rs1("horasalida"),5)
							demora = DateDiff("n", salida, ahora) / 60
							caden="select * from hojarutamr where numviaje=" & rs4("numviaje")
							set rs5=cn.execute(caden)
							if rs5.eof=false then
								do while rs5.eof=false
									if not isnull(rs5("horallegada")) and not isnull(rs5("horasalida")) then
										mensaje="Sali� " & rs5("destinatario") &  " a las " & rs5("horallegada")
									end if
									if not isnull(rs5("horallegada")) and isnull(rs5("horasalida")) then
										mensaje="Lleg� a " & rs5("destinatario") &  " a las " & rs5("horallegada")
										ahora=hour(now)& ":" & minute(now)
										llegada=left(rs5("horallegada"),5)
										demora = DateDiff("n", llegada, ahora) / 60
									end if
									rs5.movenext
								loop
							end if
                				End If
                				If Not IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) Then
                    					Mensaje = "Lleg� a " & RS1("destinatario") &  " a las " & rs1("horallegada")
                    					vllegada = 1
							ahora=hour(now)& ":" & minute(now)
							llegada=left(rs1("horallegada"),5)
							demora = DateDiff("n", llegada, ahora) / 60
                				End If
                				If IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) And vsalida = 1 Then
                    					Mensaje = "En camino a " & RS1("destinatario")
							vsalida=0
                				Else
                    					nada = 1
                				End If
            				End If
            				If tramo = 4 Then
                				If Not IsNull(RS1("horallegada")) And Not IsNull(RS1("horasalida")) Then
                    					Mensaje = "Sali� de " & RS1("destinatario") &  " a las " & rs1("horasalida")
                    					vsalida = 1
							ahora=hour(now)& ":" & minute(now)
							salida=left(rs1("horasalida"),5)
							demora = DateDiff("n", salida, ahora) / 60
                				End If
                				If Not IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) Then
                    					Mensaje = "Lleg� a " & RS1("destinatario") &  " a las " & rs1("horallegada")
                    					vllegada = 1
							ahora=hour(now)& ":" & minute(now)
							llegada=left(rs1("horallegada"),5)
							demora = DateDiff("n", llegada, ahora) / 60
                				End If
                				If IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) And vsalida = 1 Then
                    					Mensaje = "En camino a " & RS1("destinatario")
							vsalida=0
                				Else
                    					nada = 1
                				End If
            				End If
            				If tramo = 5 Then
                				If Not IsNull(RS1("horallegada")) And Not IsNull(RS1("horasalida")) Then
                    					Mensaje = "Sali� de " & RS1("destinatario") &  " a las " & rs1("horasalida")
                    					vsalida = 1
							ahora=hour(now)& ":" & minute(now)
							salida=left(rs1("horasalida"),5)
							demora = DateDiff("n", salida, ahora) / 60
                				End If
                				If Not IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) Then
                   	 				Mensaje = "Lleg� a " & RS1("destinatario") &  " a las " & rs1("horallegada")
                    					vllegada = 1
							ahora=hour(now)& ":" & minute(now)
							llegada=left(rs1("horallegada"),5)
							demora = DateDiff("n", llegada, ahora) / 60
                				End If
                				If IsNull(RS1("horallegada")) And IsNull(RS1("horasalida")) And vsalida = 1 Then
                    					Mensaje = "En camino a " & RS1("destinatario")
                				Else
                    					nada = 1
                				End If
            				End If
					rs1.movenext
				loop
				rs1.close
				if demora>1 then%>
					<tr bgcolor="#F72E4F">
						<td style="font-family: Verdana;"><font color="white"><h4><%=rs4("nombrecond")%></h4></font></td>
						<td style="font-family: Verdana;"><center><font color="white"><h4><%=rs4("patente")%></h4></center></td>
						<td style="font-family: Verdana;"><font color="white"><h4><center><a href="javascript:Abrir_iqnet('consultalchrMR.asp?numviaje=<%=rs4("numviaje")%>')"><%=rs4("numviaje")%></a></center></h4></font></td>
						<td style="font-family: Verdana;"><font color="white"><h4><center><%=rs4("fechasalida")%></center></h4></font></td>
						<td style="font-family: Verdana;"><font color="white"><h4><center><%=mensaje%></center></h4></font></td>
					</tr>	
				<%else%>
					<tr bgcolor="#49F66B">
						<td style="font-family: Verdana;"><h4><%=rs4("nombrecond")%></h4></td>
						<td style="font-family: Verdana;"><center><h4><%=rs4("patente")%></h4></center></td>
						<td style="font-family: Verdana;"><h4><center><a href="javascript:Abrir_iqnet('consultalchrMR.asp?numviaje=<%=rs4("numviaje")%>')"><%=rs4("numviaje")%></a></center></h4></td>
						<td style="font-family: Verdana;"><h4><center><%=rs4("fechasalida")%></center></h4></td>
						<td style="font-family: Verdana;"><h4><center><%=mensaje%></center></h4></td>
					</tr>	
				<%end if%>
				<%viajes=viajes + cuenta%>	
			<%end if%>

			<%if rs4("etapa")="G" then
				g=g+1%>
				<tr bgcolor="#FFFFFF">
					<td style="font-family: Verdana;"><h4><%=rs4("nombrecond")%></h4></td>
					<td style="font-family: Verdana;"><center><h4><%=rs4("patente")%></h4></center></td>
					<td style="font-family: Verdana;"><h4><center><a href="javascript:Abrir_iqnet('consultalchrMR.asp?numviaje=<%=rs4("numviaje")%>')"><%=rs4("numviaje")%></a></center></h4></td>
					<td style="font-family: Verdana;"><h4><center><%=rs4("fechasalida")%></center></h4></td>
					<td style="font-family: Verdana;"><h4><center>SIN MOVIMIENTIOS</center></h4></td>
				</tr>			
			<%end if
			rs4.movenext
		loop%>
	</table>
	</center>
	<center>&nbsp;&nbsp;&nbsp;&nbsp;<font face="verdana" size="8" color="#49F66B">ABIERTAS:&nbsp;<%=a%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="verdana" size="8" color="#FFFFFF">GENERADAS:&nbsp;<%=g%></font>
	<br>
	<hr>
	<form name="form2" action="listadeconductores1.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
		<center><input type="submit" class="boton_2" value="  REGRESAR  "></center>
	</form>
	</center>
	<hr>
</body>
</html>
<%movimt=request("movimt")%>
<%end if%>


<%if request.form("movimt")="MOVIMIENTOS INTERNOS" then%>
<body background="fondo_ondulado.jpg">
	<center>
	<div class="box-blue">
		<div id="primero"><img src="logoCATNUEVO.gif" width="80" heigth="80"></div>
		<div id="segundo"><big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span
		style="font-family: Helvetica,Arial,sans-serif;">Tablero Diario</span></big></big><br></big></div>
		<center><font face="verdana" size="4"><b>Fecha:&nbsp;<%=date()%>&nbsp;Hora:&nbsp;<%=time()%></b></font></center>
	</div>

	<center>
	<font face="verdana" size="14" color="white"><b>MOVIMIENTOS INTERNOS</b></font>

	<center>
	<table width="100%">
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3>Conductor</h3></td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3>Proveedor</h3></td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3><center>Cant. Viajes</center></h3></td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3><center>N� H.R.</center></h3></td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(250, 247, 246);"><h3><center>Fecha H.R.</center></h3></td>	
		</tr>
		<%a=0
		g=0
		viajes=0%>

		<%cade="select v.etapa, c.nombre as nombrecond, p.RSProveedor as nombreprov, v.NumViaje, v.FechaSalida from viajes v, Conductores c, proveedores p where v.IDConductor =c.IDConductor and v.CodProveedor=p.IDProveedor and (v.etapa='" & "A" & "' or v.etapa='" & "G" & "') and v.tipomov='" & "I" & "' order by c.Nombre"
		set rs4=cn.execute(cade)
		do while rs4.eof=false
			cad="select count (numviaje) as cuenta from hojaruta where numviaje=" & rs4("numviaje")
			set rs1=cn.execute(cad)
			cuenta=rs1("cuenta")%>
			<%if rs4("etapa")="A" then
				a=a+1%>
				<tr bgcolor="#49F66B">
					<td style="font-family: Verdana;"><h4><%=rs4("nombrecond")%></h4></td>
					<td style="font-family: Verdana;"><h4><%=rs4("nombreprov")%></h4></td>
					<td style="font-family: Verdana;"><h4><center><%=cuenta%></center></h4></td>
					<td style="font-family: Verdana;"><h4><center><a href="javascript:Abrir_iqnet('consultalchr.asp?numviaje=<%=rs4("numviaje")%>')"><%=rs4("numviaje")%></a></center></h4></td>
					<td style="font-family: Verdana;"><h4><center><%=rs4("fechasalida")%></center></h4></td>
				</tr>	
				<%viajes=viajes + cuenta%>	
			<%end if%>

			<%if rs4("etapa")="G" then
				g=g+1%>
				<tr bgcolor="#F67049">
					<td style="font-family: Verdana;"><h4><%=rs4("nombrecond")%></h4></td>
					<td style="font-family: Verdana;"><h4><%=rs4("nombreprov")%></h4></td>
					<td style="font-family: Verdana;"><h4><center>0</center></h4></td>
					<td style="font-family: Verdana;"><h4><center><a href="javascript:Abrir_iqnet('consultalchr.asp?numviaje=<%=rs4("numviaje")%>')"><%=rs4("numviaje")%></a></center></h4></td>
					<td style="font-family: Verdana;"><h4><center><%=rs4("fechasalida")%></center></h4></td>
				</tr>			
			<%end if%>

			<%rs4.movenext
		loop%>
	</table>
	</center>
	<center><font face="verdana" size="8" color="#000000">VIAJES:&nbsp;<%=viajes%>&nbsp;&nbsp;&nbsp;&nbsp;<font face="verdana" size="8" color="#49F66B">ABIERTAS:&nbsp;<%=a%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="verdana" size="8" color="#F67049">GENERADAS:&nbsp;<%=g%></font>
	<br>
	<hr>
	<form name="form2" action="listadeconductores1.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
		<center><input type="submit" class="boton_2" value="  REGRESAR  "></center>
	</form>
	</center>
	<hr>
</body>
</html>
<%movimt=movimt%>
<%end if%>