<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
usuario=request.querystring("usuario")
clave=request.querystring("clave")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
cn.execute("delete hrvalidadas")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>  
	<meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
  	<title>Consulta de Hojas de Ruta Validadas</title>

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
	background: url(cielo.jpg) "CCC no-repeat";
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
	content: 'S�';
	font: 12px/30px Arial, sans-serif;
	color: lime;
	position: absolute;
	left: 10px;
	z-index: 0;
	font-weight: bold;
}


</style>  	
</head>
<body background="fondo_ondulado.jpg">
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
			Hojas de Ruta Validadas</span></big></big><br></big>
				<%mes=month(request.form("txtdesde"))%>
				<%if mes<10 then
					mes="0" & month(request.form("txtdesde"))
				end if%>
				<%dia=day(request.form("txtdesde"))%>
				<%if dia <10 then
					dia="0" & day (request.form("txtdesde"))
				end if%>					
				<%anio=year(request.form("txtdesde"))%>
				<%cadenadesde=dia & "/" & mes & "/" & anio%>
				<%desde=mes & "/" & dia & "/" & anio%>
				<%mes=month(request.form("txthasta"))%>
				<%if mes<10 then
					mes="0" & month(request.form("txthasta"))
				end if%>
				<%dia=day(request.form("txthasta"))%>
				<%if dia <10 then
					dia="0" & day (request.form("txthasta"))
				end if%>					
				<%anio=year(request.form("txthasta"))%>
				<%cadenahasta=dia & "/" & mes & "/" & anio%>
				<%hasta=mes & "/" & dia & "/" & anio%>
				<td style="text-align: center;"><big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span
				style="font-family: Helvetica,Arial,sans-serif;">Desde:&nbsp; <%=cadenadesde%>&nbsp;&nbsp;&nbsp;Hasta:&nbsp;<%=cadenahasta%></span></big></big></big></td>
		</div>
	</div>
	<center>
	<table border="1">
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha H.R.</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">N� Hoja de Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Patente</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Conductor</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Inicio Validaci�n</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Fin Validaci�n</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Horas</td>
		</tr>
		<%cade="select * from viajes where fechasalida between '" & desde & "' and '" & hasta & "' order by numviaje"
		set rs=cn.execute(cade)
		totalhoras=0
		Do While RS.EOF = False
			if rs("etapa")<>"A" or rs("etapa")<>"G" then
    			If Not IsNull(RS("horainiciovalidada")) And Not IsNull(RS("horafinvalidada")) Then%>
				<tr bgcolor="#4DF603 ">
					<td style="font-family: Verdana;"><h6><center><%=RS("fechasalida")%></center></h6></td>
      					<td style="font-family: Verdana;"><h6><center><%=RS("numviaje")%></center></h6></td>
        				<%caden="select patente from medios where idmedio=" & RS("idmedio")
					set rs2=cn.execute(caden)
        				patente = Trim(RS2("patente"))
        				RS2.Close%>
        				<td style="font-family: Verdana;"><h6><%=patente%></h6></td>
        				<%caden="select nombre from conductores where idconductor=" & RS("idconductor")
					set rs2=cn.execute(caden)
        				conductor = Trim(RS2("nombre"))
        				RS2.Close%>
					<td style="font-family: Verdana;"><h6><%=conductor%></h6></td>
        				<td style="font-family: Verdana;"><center><h6><%=RS("fechainiciovalidada")%>-<%=Trim(RS("horainiciovalidada"))%></h6></center></td>
        				<td style="font-family: Verdana;"><center><h6><%=RS("fechafinvalidada")%>-<%=Trim(RS("horafinvalidada"))%></h6></center></td>
        				<%If Trim(RS("horafinvalidada")) > Trim(RS("horainiciovalidada")) Then
						cadena = Trim(RS("horafinvalidada"))
 						hora = Mid(cadena, 1, 2)
						minuto = Mid(cadena, 4, 2)
            					If minuto < 15 Then
                					hora = hora
					                minuto = "00"
					        End If
            					If minuto >= 15 And minuto <= 44 Then
                					hora = hora
                					minuto = "30"
            					End If
            					If minuto > 44 And minuto <= 59 Then
                					hora = hora + 1
                					If hora = 1 Then
                    						hora = "01"
                					End If
                					If hora = 2 Then
                    						hora = "02"
                					End If
                					If hora = 3 Then
                    						hora = "03"
               			 			End If
                					If hora = 4 Then
                    						hora = "04"
                					End If
                					If hora = 5 Then
                    						hora = "05"
                					End If
                					If hora = 6 Then
                    						hora = "06"
                					End If
    							If hora = 7 Then
                   						hora = "07"
                					End If
                					If hora = 8 Then
                    						hora = "08"
                					End If
                					If hora = 9 Then
                    						hora = "09"
                					End If
                					minuto = "00"
						End If
            					finvalida = hora & ":" & minuto
						cadena = Trim(RS("horainiciovalidada"))
 						hora = Mid(cadena, 1, 2)
						minuto = Mid(cadena, 4, 2)
            					If minuto < 15 Then
					                hora = hora
					                minuto = "00"
					        End If
					        If minuto >= 15 And minuto <= 44 Then
					                hora = hora
					                minuto = "30"
					        End If
					        If minuto > 44 And minuto <= 59 Then
					                hora = hora + 1
                					If hora = 1 Then
                    						hora = "01"
                					End If
                					If hora = 2 Then
                    						hora = "02"
                					End If
                					If hora = 3 Then
                    						hora = "03"
               			 			End If
                					If hora = 4 Then
                    						hora = "04"
                					End If
                					If hora = 5 Then
                    						hora = "05"
                					End If
                					If hora = 6 Then
                    						hora = "06"
                					End If
    							If hora = 7 Then
                   						hora = "07"
                					End If
                					If hora = 8 Then
                    						hora = "08"
                					End If
                					If hora = 9 Then
                    						hora = "09"
                					End If
							minuto = "00"
					     	End If
            					iniciovalida = hora & ":" & minuto
            					If IsNull(RS("horafinvalidada")) Or IsNull(RS("horainiciovalidada")) Then
					                totalhs = 0
					        Else
							if finvalida="24:00" then
								finvalida="00:00"
								totalhs = DateDiff("n", iniciovalida, finvalida) / 60 + 24
							else
					                	totalhs = DateDiff("n", iniciovalida, finvalida) / 60
							end if
					        End If
        				Else
            					If Not IsNull(RS("horainiciovalidada")) And Not IsNull(RS("horafinvalidada")) Then
							cadena = Trim(RS("horafinvalidada"))
	 						hora = Mid(cadena, 1, 2)
							minuto = Mid(cadena, 4, 2)
            						If minuto < 15 Then
                						hora = hora
						                minuto = "00"
						        End If
            						If minuto >= 15 And minuto <= 44 Then
                						hora = hora
	                					minuto = "30"
        	    					End If
            						If minuto > 44 And minuto <= 59 Then
                						hora = hora + 1
                						If hora = 1 Then
                    							hora = "01"
                						End If
                						If hora = 2 Then
                    							hora = "02"
	                					End If
        	        					If hora = 3 Then
                	    						hora = "03"
               				 			End If
                						If hora = 4 Then
                    							hora = "04"
                						End If
                						If hora = 5 Then
                    							hora = "05"
	                					End If
        	        					If hora = 6 Then
                	    						hora = "06"
                						End If
    								If hora = 7 Then
                   							hora = "07"
                						End If
                						If hora = 8 Then
                    							hora = "08"
	                					End If
        	        					If hora = 9 Then
                	    						hora = "09"
                						End If
                						minuto = "00"
							End If
            						finvalida = hora & ":" & minuto
                                    
							cadena = Trim(RS("horainiciovalidada"))
 							hora = Mid(cadena, 1, 2)
							minuto = Mid(cadena, 4, 2)
            						If minuto < 15 Then
					                	hora = hora
						                minuto = "00"
						        End If
						        If minuto >= 15 And minuto <= 44 Then
						                hora = hora
						                minuto = "30"
						        End If
					        	If minuto > 44 And minuto <= 59 Then
						                hora = hora + 1
        	        					If hora = 1 Then
                	    						hora = "01"
                						End If
                						If hora = 2 Then
                    							hora = "02"
                						End If
                						If hora = 3 Then
                    							hora = "03"
	               			 			End If
        	        					If hora = 4 Then
                	    						hora = "04"
                						End If
                						If hora = 5 Then
                    							hora = "05"
                						End If
                						If hora = 6 Then
                    							hora = "06"
	                					End If
    								If hora = 7 Then
                	   						hora = "07"
                						End If
                						If hora = 8 Then
                    							hora = "08"
                						End If
                						If hora = 9 Then
                    							hora = "09"
	                					End If
								minuto = "00"
						     	End If
            						iniciovalida = hora & ":" & minuto

	                				If IsNull(RS("horafinvalidada")) Or IsNull(RS("horainiciovalidada")) Then
					                	totalhs = 0
					                Else
                   						totalhs = DateDiff("n", iniciovalida, finvalida) / 60 + 24
                					End If
            					Else	
			                		totalhs = 0
            					End If
        				End If%>
        				<td style="font-family: Verdana;"><h6><%=totalhs%></h6></td>
				</tr>
				<%rs4.CursorType = adOpenKeyset
				rs4.LockType = adLockOptimistic
				rs4.open"hrvalidadas",cn
				rs4.addnew
				rs4("fecha")=rs("fechasalida")
				rs4("numviaje")=rs("numviaje")
				rs4("fechainicio")=RS("fechainiciovalidada")
				rs4("horainicio")=Trim(RS("horainiciovalidada"))
				rs4("fechafin")=RS("fechafinvalidada")
				rs4("horafin")=Trim(RS("horafinvalidada"))
				rs4("horas")=totalhs
				rs4("patente")=patente
				rs4("conductor")=conductor
				rs4("validoinicio")=RS("usuariovalidainicio")
				rs4("validofin")=RS("usuariovalidafin")
				rs4.update
				rs4.close
				totalhoras=totalhoras + totalhs
			end if
			end if
			rs.movenext
		loop%>
	</table>
	</center>
	<center>
	<br>
	<hr>
	<font face="verdana" color="white" size="6">Total de Horas Validadas:&nbsp;<%=totalhoras%></font>
	<hr>
	</center>
	<center>
	<table>
		<tr>
			<td>
				<form name="form5" action="excelhrvalidadas.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>&desde=<%=cadenadesde%>&hasta=<%=cadenahasta%>" method="POST">
					<center><input type="submit" class="boton_2" value="  EXPORTAR A EXCEL  "></center>
				</form>
			</td>
			<td>
				<form name="form3" action="listadeconductores1.asp?usuario=<%=request.querystring("usuario")%>&clave=<%=request.querystring("clave")%>" method="POST">
					<center><input type="submit" class="boton_2" value="  REGRESAR  "></center>
				</form>	
			</td>
		</tr>
	</table>
	</center>
</body>
</html>
