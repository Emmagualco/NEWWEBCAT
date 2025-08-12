<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
usuario=request.form("usuario")
clave=request.form("clave")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
cade="select * from usuariosrasa where nombre='" & trim(request.form("usuario")) & "' and clave='" & trim(request.form("clave")) & "'"
set rs=cn.execute(cade)
if rs.eof=false then%>
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
	height: 150px;
  	width: 950px;
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
	background: url(cielo.jpg) "CCC no-repeat;
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

  .boton_3{
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
  .boton_3:hover{
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

/* ESTILOS PARA EL CHECKBOX */

input[type=checkbox] {
	visibility: hidden;
}

.checkbox-JASoft {
	display: inline-block;
	position: relative;
    width: 70px;
	height: 30px;
	background:#630c46;
	border-radius: 15px;
	box-shadow: inset 0px 1px 1px rgba(0,0,0,0.6), 0px 1px 0px rgba(255,255,255,0.3);   
}

.checkbox-JASoft label {

    /* aspecto */
    display: block;
    width: 34px;
	height: 20px;
	border-radius: 17px;
	box-shadow: 0px 2px 5px 0px rgba(0,0,0,0.35);
	background: blue;
	background: linear-gradient(to top, #e69e32 0%, #ebd5e4 40%, #59ebf0 100%);
    cursor: pointer;
    
    /* Posicionamiento */
    position: absolute;
    top: 5px;
	left: 5px;
    z-index: 1;
    
	/* Comportamiento */
    transition: all .4s ease;
    
    /* ocultar el posible texto que tenga */
    overflow: hidden;
    text-indent: 35px;  
    transition: text-indent 0s;
}

/* estado activado */
.checkbox-JASoft input[type=checkbox]:checked + label {
	left: auto;
    right: 5px;
}

.checkbox-JASoft:after {
	content: 'NO';
	font: 12px/30px Arial, sans-serif;
	color: #AAA;
	position: absolute;
	right: 10px;
    z-index: 0;
	font-weight: bold;
	text-shadow: 1px 1px 0px rgba(255,255,255,.20);
}

.checkbox-JASoft:before {
	content: 'SÍ';
	font: 12px/30px Arial, sans-serif;
	color: lime;
	position: absolute;
	left: 10px;
	z-index: 0;
	font-weight: bold;
}


</style>  	
</head><body background="fondo_ondulado.jpg">

<table width="100%">
	<tr>
		<td width="100%" align="center"><font face="verdana" size="6" color="white">Bienvenido <%=request("usuario")%></font></td>
	</tr>
</table>
<hr>

<center>
<div class="box-blue">
	<div id="primero"><img src="logoCATNUEVO.gif" width="80" heigth="80"></div>
	<div id="segundo"><big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span
style="font-family: Helvetica,Arial,sans-serif;">Consulta de
Recorridos por Conductor</span></big></big><br>
</big>
</div>
</div>
<div class="box-gray">

<table width="800px" border="0">
<tr bgcolor="#0288d1">
<td align="center">
<br>
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

<form name="form2" action="consultainformes.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="post">
	<font color="white" face="verdana">Seleccione Conductor:</font> 
	<select name="cmbconductor" required>
		<%cade="select distinct nombre from conductores where estado='" & "A" & "' order by nombre"
		set rs=cn.execute(cade)
		do while rs.eof=false%>
			<option><%=trim(rs("nombre"))%></option>
			<%rs.movenext
		loop%>
	</select>
	<br>
	<br>
	<font color="white" face="verdana">Seleccionar TODOS los Conductores</font>	
        <div class="checkbox-JASoft">
            <input type="checkbox" name="todos" id="checkAvanzado" value="Valor"/>
            <label for="checkAvanzado">TEXTO QUE NO DEBERÍA VERSE</label>
        </div>
	<br>
	<br>	
	<font color="white" face="verdana">Desde: <input name="txtdesde" step="1" type="date" value="<%=cadenahoy%>" required>   Hasta: <input name="txthasta" step="1" type="date" value="<%=cadenahoy%>" required></font>
	<input value="   Consultar   " type="submit" class="boton_2"><br>
</form>
</td>
</tr>
</table> 
</center>
</div>
<center>
<form name="form3" action="nuevotramo.html">
	<button class="boton_2" onclick="location.href='nuevotramo.html'">REGRESAR</button>
</form>
</center>
</body>
</html>
<%else%>
	<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>  
		<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
		<title>Error de Ingreso</title>
			<link rel="stylesheet" href="estilohr.css">
		</head>
		<body bgcolor="#9091A0">
			<br>
			<br>
			<br>
			<br>
			<div class="box-green">
				<center><font face="verdana" size="10">No existe combinación N° Usuario / Password. Por favor, verifique los datos ingresados.</font></center>
			</div>
			<br>
			<br>
			<div>
				<form name="form2" action="nuevotramo.html">
					<center><a href="nuevotramo.html"><img style="width: 100px; height: 100px;" src="flechaatras.jpg"></a></center>
				</form>
			</div> 
		</body>
	</html>	

<%end if%>	