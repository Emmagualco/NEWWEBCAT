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
dim total
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
    font-family: arial black;
    text-transform: uppercase;
    padding-left: 10px;
    padding-right: 10px;
    font-weight: 100;
    font-size: 20px;
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

	<body background="fondo_ondulado.jpg">
	<center>
	<div class="box-blue">
		<div id="primero"><img src="logoCATNUEVO.gif" width="80" heigth="80"></div>
		<div id="segundo"><big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span
		style="font-family: Helvetica,Arial,sans-serif;">Consulta de Recorridos por Conductor</span></big></big><br></big></div>
	</div>
	<div class="box-gray">
		<table width="800px" border="0">
			<tr bgcolor="#0288d1">
				<%cade="select * from conductores where documento='" & request.querystring("dni") & "' and claveweb='" & request.querystring("clave") & "'"
				set rs=cn.execute(cade)
				if rs.eof=false then
					idconductor=rs("idconductor")
					nombrecond=trim(rs("nombre"))
					foto=rs("foto")
				end if%>
				<td colspan="1" rowspan="2" style="vertical-align: middle;"><img src="<%=rs("foto")%>" width="100" height="100"></td>
				<td style="text-align: center;"><big style="font-weight: bold;"><big style="color: rgb(255, 255, 255);"><big><span
				style="font-family: Helvetica,Arial,sans-serif;">Conductor:&nbsp;<%=nombrecond%></span></big></big></big></td>
			</tr>
			<tr align="center">
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
				<td style="text-align: center;"><big style="font-weight: bold;"><big style="color: rgb(255, 255, 255);"><big><span
				style="font-family: Helvetica,Arial,sans-serif;">Desde:&nbsp; <%=cadenadesde%>&nbsp;&nbsp;&nbsp;Hasta:&nbsp;<%=cadenahasta%></span></big></big></big></td>
			</tr>
		</table>
	</div>
	<center>
	<table>
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">N� Hoja de Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Origen</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Destino</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Tipo Mov.</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Horas Empleadas</td>
		</tr>
		<%total=0%>
		<%totalhs=0%>
		<%cade="select * from conductores where documento='" & request.querystring("dni") & "' and claveweb='" & request.querystring("clave") & "'"
		set rs4=cn.execute(cade)
		do while rs4.eof=false
			idconductor=rs4("idconductor")
			cad="select * from viajes where idconductor=" & idconductor & " and fechasalida between '" & request.form("txtdesde") & "' and '" & request.form("txthasta") & "' AND web='" & "S" & "' order by fechasalida"
			set rs1=cn.execute(cad)
			do while rs1.eof=false
				cade="select * from hojaruta where numviaje=" & rs1("numviaje") & " order by idhojaruta"
				set rs2=cn.execute(cade)
				if rs2.eof=false then
					color1=0
					color2=0
					inicio = Trim(RS2("horallegada"))
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
		
						<td style="font-family: Verdana;"><h6><%=rs1("fechasalida")%></h6></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("hojaderuta")%></h6></center></td>
						<td style="font-family: Verdana;"><h6><%=rs2("remitente")%></h6></td>
						<td style="font-family: Verdana;"><h6><%=rs2("destinatario")%></h6></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("fechallegada")& "-" & trim(rs2("horallegada"))%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("fechasalida") & "-" & trim(rs2("horasalida"))%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("tipo")%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("quellevaorigen")%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("quellevadestino")%></h6></center></td>
						<%if not isnull(rs2("horallegada")) and not isnull(rs2("horasalida")) then
        						If Trim(RS2("horasalida")) > inicio Then
						        	cadena = Trim(RS2("horasalida"))
 							        hora = Mid(cadena, 1, 2)
							        minuto = Mid(cadena, 4, 2)
    
            							If minuto < "15" Then
                							hora = hora
                							minuto = "00"
            							End If
            							If minuto >= "15" And minuto <= "44" Then
                							hora = hora
                							minuto = "30"
            							End If
            							If minuto > "44" And minuto <= "59" Then
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
            							horadesalida = hora & ":" & minuto
            							cadena = Trim(inicio)
    
    		        					hora = Mid(cadena, 1, 2)
            							minuto = Mid(cadena, 4, 2)
            							If minuto < "15" Then
                							hora = hora
                							minuto = "00"
            							End If
            							If minuto >= "15" And minuto <= "44" Then
                							hora = hora
                							minuto = "30"
            							End If
            							If minuto > "44" And minuto <= "59" Then
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
                						minuto = "0"
            						End If
    
            						horadellegada = hora & ":" & minuto

			            			If IsNull(RS2("horasalida")) Or IsNull(RS2("horallegada")) Then
                    						horast = 0
            						Else
                						If horadesalida = "24:00" Then
                    							horadesalida = "00:00"
                    							horast = DateDiff("n", horadellegada, horadesalida) / 60 + 24
                						Else
                    							horast = DateDiff("n", horadellegada, horadesalida) / 60
                						End If
                					End If

            						If horast > 0 Then%>
                						<td style="font-family: Verdana;"><center><h6><%=FormatNumber(horast,2)%></h6></center></td>
            						<%Else%>
                						<td style="font-family: Verdana;"><center><h6>0</h6></center></td>
            						<%End If
            						inicio = Trim(horadesalida)
            						fechainicio = RS2("fechasalida")
            						totalhs = totalhs + horast
            						horast = 0
        					Else
            						If RS2("fechasalida") = fechainicio Then
                						horast = 0
                						totalhs = totalhs + horast
                						horadesalida = inicio
                						horadellegada = inicio
            						Else
                						cadena = Trim(RS2("horasalida"))
                						hora = Mid(cadena, 1, 2)
                						minuto = Mid(cadena, 4, 2)
                						If minuto < "15" Then
                        						hora = hora
                        						minuto = "00"
                						End If
                						If minuto >= "15" And minuto <= "44" Then
                        						hora = hora
                        						minuto = "30"
                						End If
               		 					If minuto > "44" And minuto <= "59" Then
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
                					horadesalida = hora & ":" & minuto
                					cadena = Trim(inicio)
                					hora = Mid(cadena, 1, 2)
                					minuto = Mid(cadena, 4, 2)
                					If minuto < 15 Then
                    						hora = hora
                    						minuto = "00"
                					End If
                					If minuto >= "15" And minuto <= "44" Then
                    						hora = hora
                    						minuto = "30"
                					End If
                					If minuto > "44" And minuto <= "59" Then
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
                				horadellegada = hora & ":" & minuto
            				End If
            				If IsNull(RS2("horasalida")) Or IsNull(RS2("horallegada")) Then
                    				horast = 0
            				Else
                				If RS2("fechasalida") <> fechainicio Then
                    					horast = DateDiff("n", horadellegada, horadesalida) / 60 + 24
                				Else
                    					horast = DateDiff("n", horadellegada, horadesalida) / 60
                				End If
            				End If
            				If horast > 0 Then%>
                				<td style="font-family: Verdana;"><center><h6><%=FormatNumber(horast,2)%></h6></center></td>
            				<%Else%>
                				<td style="font-family: Verdana;"><center><h6>0</h6></center></td>
            				<%End If
            				inicio = Trim(horadesalida)
            				fechainicio = RS2("fechasalida")
            				totalhs = totalhs + horast
            				horast = 0
       				End If
    			Else%>
        			<td style="font-family: Verdana;"><center><h6>0</h6></center></td>
    			<%End If%>

			</tr>
			<%rs2.movenext
			loop%>

<%if request.querystring("tipo")<>"MILK RUN" then%>
			<%caden="select * from viajes where numviaje=" & rs1("numviaje")
			set rs5=cn.execute(caden)
			if not isnull(rs5("horainiciovalidada")) and not isnull(rs5("horafinvalidada")) then
				cadena = Trim(RS5("horafinvalidada"))
 				hora = Mid(cadena, 1, 2)
				minuto = Mid(cadena, 4, 2)
    
            			If minuto < "15" Then
                			hora = hora
                			minuto = "00"
            			End If
            			If minuto >= "15" And minuto <= "44" Then
                			hora = hora
                			minuto = "30"
            			End If
            			If minuto > "44" And minuto <= "59" Then
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
            			horafinvalidada = hora & ":" & minuto
            			cadena = Trim(rs5("horainiciovalidada"))
    
    		        	hora = Mid(cadena, 1, 2)
            			minuto = Mid(cadena, 4, 2)
            			If minuto < "15" Then
                			hora = hora
                			minuto = "00"
            			End If
            			If minuto >= "15" And minuto <= "44" Then
                			hora = hora
                			minuto = "30"
            			End If
            			If minuto > "44" And minuto <= "59" Then
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
                			minuto = "0"
            			End If
    
            			horainiciovalidada = hora & ":" & minuto

                		If horafinvalidada = "24:00" Then
                    			horafinvalidada = "00:00"
                    			hsRASA=DateDiff("n", horainiciovalidada, horafinvalidada) / 60 + 24
                		Else
                    			hsRASA=DateDiff("n", horainiciovalidada, horafinvalidada) / 60
                		End If
				if rs5("fechainiciovalidada")<>rs5("fechafinvalidada") then
					hsRASA=DateDiff("n", horainiciovalidada, horafinvalidada) / 60 + 24
				end if
			end if
end if%>
					<tr>
						<td colspan="9" align="right" style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Total Hs. Validadas por Renault:&nbsp; <%=hsRASA%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total Hs. HR&nbsp; <%=rs1("numviaje")%></td>
						<td align="center" style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);"><%=totalhs%></td>
					</tr>
					<%total=total+totalhs%>
					<%totalhs=0
					rsRASA=0%>

			<%end if
			rs2.close
				rs1.movenext
			loop
			rs1.close
			rs4.movenext
		loop%>
	</table>
	<br>
	<hr>
	<%if total>0 then%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de Horas:&nbsp;<%=FormatNumber(total,2)%></span></big></big>
	<%else%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de Horas:&nbsp;0</span></big></big>
	<%end if%>
	<hr>
	<form name="form2" action="listarhr1nva.asp?idhr=<%=request.querystring("idhr")%>&tipo=<%=trim(request.querystring("tipo"))%>&dni=<%=request.querystring("dni")%>&clave=<%=request.querystring("clave")%>" method="POST">
		<center><input type="submit" class="boton_2" value="  REGRESAR  "></center>
	</form>
	</body>
</html>