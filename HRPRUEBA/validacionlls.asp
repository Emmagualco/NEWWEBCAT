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

input[type=radio] { border: 0px; width: 3%; height: 3em; }


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
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
cade="select * from hojaruta where idhojaruta=" & request.querystring("hr")
set rs=cn.execute(cade)
if rs.eof=false then
	idhr=rs("idhojaruta")
	destino=trim(rs("destinatario"))
	rs1.CursorType = adOpenKeyset
	rs1.LockType = adLockOptimistic
	rs1.open "select observacionesrasa,validadopor,fechavalidacion,horavalidacion,fechallegadarasa,horallegadarasa,fechasalidarasa,horasalidarasa from hojaruta where idhojaruta=" & idhr,cn
	rs1("fechallegadarasa")=request.form("fechallegada")
	rs1("fechasalidarasa")=request.form("fechasalida")
	rs1("horallegadarasa")=request.form("hora1")
	rs1("horasalidarasa")=request.form("hora2")
	rs1("validadopor")=trim(request.querystring("usuario"))
	rs1("fechavalidacion")=date()
	rs1("horavalidacion")=time()
	rs1("observacionesrasa")=ucase(trim(request.form("obser")))
	rs1.update
	rs1.close
end if
rs.close%>
<body background="fondo_ondulado.jpg">
<center>
<font face="verdana" size="6" color="white">Bienvenido <%=request.querystring("usuario")%></font>
<hr>
<font face="verdana" size="6" color="white">Origen: <%=destino%></font>
<hr>
<font face="verdana" size="6" color="white">Conductor: <%=request.querystring("conductor")%></font>
<hr>

<%cad="select fechallegada,horallegada,fechasalida,horasalida,fechallegadarasa,horallegadarasa,fechasalidarasa,horasalidarasa from hojaruta where idhojaruta=" & idhr
set rs=cn.execute(cad)
if rs.eof=false then%>
<font face="verdana">
<table  border="1">
	<tbody>
		<tr>
			<td colspan="4" style="vertical-align: top; text-align: center;"><b>Se Validó</b></td>
		</tr>
		<tr>
			<td colspan="2" align="center">LLEGADA</td>
			<td colspan="2" align="center">SALIDA</td>
		</tr>
		<tr>
			<td align="right">Fecha:</td>
			<td align="left"><%=rs("fechallegadarasa")%></td>
			<td align="right">Fecha:</td>
			<td align="left"><%=rs("fechasalidarasa")%></td>
		</tr>
		<tr>
			<td align="right">Hora:</td>
			<td align="left"><%=rs("horallegadarasa")%></td>
			<td align="right">Hora:</td>
			<td align="left"><%=rs("horasalidarasa")%></td>
		</tr>

	</tbody>
</table>
</font>
<form name="form6" action="generarnuevotramo.asp?usuario=<%=request.querystring("usuario")%>&conductor=<%=request.querystring("conductor")%>&idhr=<%=idhr%>&clave=<%=request.querystring("clave")%>" method="POST">
<table>
		<tr>
			<td align="right"><font face="verdana" size="6">Próximo Destino:</font></td>
			<td>
				<%cad="select rstercero from terceros where estado='" & "A" & "' and tipomov='" & "I" & "'  order by rstercero"
				set rs=cn.execute(cad)%>
				<select name="destino" required>
					<%do while rs.eof=false
						if trim(rs("rstercero"))<>destino then%>
							<option><%=rs("rstercero")%></option>
						<%end if
						rs.movenext
					loop
					rs.close%>
				</select>
			</td>
		</tr>


		<tr>
			<td colspan="2" align="center">
				
					<input type="submit" class="boton_2" value="Generar Nuevo Tramo">
				
			</td>
</tr>
<tr><td colspan="2">&nbsp;</td></tr>
<tr>
<td colspan="2"><font face="verdana" size="4"><b>Si no existen más tramos para generar, se debe informar al conductor que cierre la Hoja de Ruta</b></font></td> 
		</tr> 
</table>
</form>
<br>
<br>
<center>
<form name="form5" action="listadeconductores1.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
	<center><input type="submit" class="boton_2" value="REGRESAR"></center>
</form>
</center>
<%end if%>







