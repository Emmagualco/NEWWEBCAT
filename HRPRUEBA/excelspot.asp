<%@Language="VBScript"%>
<%Response.Buffer = False%>
<%set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
set rs4=server.createobject("ADODB.Recordset")
set rs5=server.createobject("ADODB.Recordset")
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename=excelspot.xls"
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)%>
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
<big style="color: rgb(0, 0, 0);"><big><span style="font-family: Helvetica,Arial,sans-serif;"><u>Consulta de Recorridos Desde:&nbsp; <%=cadenadesde%>&nbsp;&nbsp;&nbsp;Hasta:&nbsp;<%=cadenahasta%></u></span></big></big>
<%if request.querystring("unomi")=1 then%>
<br>
<big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span style="font-family: Helvetica,Arial,sans-serif;">Consulta de Recorridos por Conductor:&nbsp;<%=trim(request.querystring("conductor"))%></span></big></big><br></big>
	<table>
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">N° Hoja de Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Origen</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Destino</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Tipo Mov.</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Horas Empleadas</td>
		</tr>
		<%totalhs=0%>
		<%cade="select idconductor from conductores where nombre='" & trim(request.querystring("conductor")) & "'"
		set rs4=cn.execute(cade)
		do while rs4.eof=false
			idconductor=rs4("idconductor")
			cad="select * from viajes where idconductor=" & idconductor & " and fechasalida between '" & request.querystring("desde") & "' and '" & request.querystring("hasta") & "' AND web='" & "S" & "' and tipomov='" & "S" & "' order by fechasalida"
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
									if horadellegada<>"00:00" then
                    								horast = DateDiff("n", horadellegada, horadesalida) / 60 + 24
									else
										horast=0
									end if
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
			loop
			end if
			rs2.close
			rs1.movenext
		loop
		rs1.close
			rs4.movenext
		loop%>
	</table>
	<br>
	<hr>
	<%if totalhs>0 then%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de Horas:&nbsp;<%=FormatNumber(totalhs,2)%></span></big></big>
	<%else%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de Horas:&nbsp;0</span></big></big>
	<%end if%>
	<hr>
	</center>
<%else%>
	<center>
		<big style="font-weight: bold;"><big style="color: rgb(0, 0, 153);"><big><span style="font-family: Helvetica,Arial,sans-serif;">Consulta de Recorridos TODOS los Conductores</span></big></big><br></big>
		<table width="800px" border="0">
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
	<br>
	<center>
	<table>
		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">N° Hoja de Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Conductor</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Origen</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Destino</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Tipo Mov.</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Horas Empleadas</td>
		</tr>
		<%totalhs=0%>
		<%cad="select * from viajes where fechasalida between '" & request.querystring("desde") & "' and '" & request.querystring("hasta") & "' AND web='" & "S" & "' and tipomov='" & "S" & "' order by fechasalida, idconductor"
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
					<%cad6="select nombre from conductores where idconductor=" & rs1("idconductor")
					set rs5=cn.execute(cad6)
					conductor=rs5("nombre")%>
					<td style="font-family: Verdana;"><center><h6><%=conductor%></h6></center></td>
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
									if horadellegada<>"00:00" then
                    								horast = DateDiff("n", horadellegada, horadesalida) / 60 + 24
									else
										horast=0
									end if
                						Else
                    							horast = DateDiff("n", horadellegada, horadesalida) / 60
                						End If
							end if
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
            			End If%>
            			<%If IsNull(RS2("horasalida")) Or IsNull(RS2("horallegada")) Then
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
				loop
			end if
			rs2.close
			rs1.movenext
		loop%>
	</table>
	<br>
	<hr>
	<%if totalhs>0 then%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de Horas:&nbsp;<%=FormatNumber(totalhs,2)%></span></big></big>
	<%else%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de Horas:&nbsp;0</span></big></big>
	<%end if%>
	<hr>
<%end if%>