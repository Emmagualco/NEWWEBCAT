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
Response.AddHeader "content-disposition", "inline; filename=excelmilkrun.xls"
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)%>
<%if request.querystring("unomr")=1 then%>
	<%mes=month(request.querystring("desde"))%>
	<%if mes<10 then
		mes="0" & month(request.querystring("desde"))
	end if%>
	<%dia=day(request.querystring("desde"))%>
	<%if dia <10 then
		dia="0" & day (request.querystring("desde"))
	end if%>					
	<%anio=year(request.querystring("desde"))%>
	<%cadenadesde=dia & "/" & mes & "/" & anio%>
	<%mes=month(request.querystring("hasta"))%>
	<%if mes<10 then
		mes="0" & month(request.querystring("hasta"))
	end if%>
	<%dia=day(request.querystring("hasta"))%>
	<%if dia <10 then
		dia="0" & day (request.querystring("hasta"))
	end if%>					
	<%anio=year(request.querystring("hasta"))%>
	<%cadenahasta=dia & "/" & mes & "/" & anio%>
	<br>
	<br>
	<center>
	<table>
		<tr>
			<td colspan="10" align="center"><font face="verdana" size="6"><u>Consulta de Recorridos Desde:&nbsp; <%=cadenadesde%>&nbsp;&nbsp;&nbsp;Hasta:&nbsp;<%=cadenahasta%></u></font>
		</td>
		</tr>

		<tr>
			<td colspan="10" align="center"><font face="verdana" size="6" color="blue">MILK RUN</font></td>
		</tr>
		<tr>
			<td colspan="10" align="center"><font face="verdana" size="6" color="blue">Conductor:&nbsp;<%=request.querystring("conductor")%></font></td>
		</tr>

		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">N° Hoja de Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Origen</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Destino</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Patente</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Prevista Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Prevista Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Horas</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">KMS</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Llegada Real</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Salida Real</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Observaciones Conductor</td>
		</tr>
		<%totalhs=0%>

		<%cade="select idconductor from conductores where nombre='" & trim(request.querystring("conductor")) & "'"
		set rs4=cn.execute(cade)
		kms=0
		do while rs4.eof=false
			idconductor=rs4("idconductor")
			cad="select * from viajes where idconductor=" & idconductor & " and fechasalida between '" & request.querystring("desde") & "' and '" & request.querystring("hasta") & "' AND web='" & "S" & "' and tipomov='" & "M" & "' and shipment is null order by fechasalida"
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
						<%cade="select m.patente from viajes v, medios m where v.idmedio=m.idmedio and v.numviaje=" & rs1("numviaje")
						set rs7=cn.execute(cade)
						patente=trim(rs7("patente"))
						rs7.close%>
						<td style="font-family: Verdana;"><h6><%=PATENTE%></h6></td>
						<td style="font-family: Verdana;" bgcolor="#c39bd3"><center><h6><%=rs2("fechallegada")& " " & trim(rs2("horallegada"))%></h6></center></td>
						<td style="font-family: Verdana;" bgcolor="#c39bd3"><center><h6><%=rs2("fechallegadarasa")& " " & trim(rs2("horallegadarasa"))%></h6></center></td>
						<td style="font-family: Verdana;" bgcolor="#ccd1d1"><center><h6><%=rs2("fechasalida") & " " & trim(rs2("horasalida"))%></h6></center></td>
						<td style="font-family: Verdana;" bgcolor="#ccd1d1"><center><h6><%=rs2("fechasalidarasa") & " " & trim(rs2("horasalidarasa"))%></h6></center></td>
            					<%horast=0%>
						<%If not IsNull(RS2("horasalida")) and not IsNull(RS2("horallegada")) Then
                    					horast = DateDiff("n", RS2("horallegada"), RS2("horasalida")) 
							If horast >= 60 Then
    								horan = Int(horast / 60)
							Else
    								horan = 0
							End If
							minutos = horast - (horan * 60)
							If minutos < 15 Then
    								horast = horan
							End If
							If minutos >= "15" And minutos <= "44" Then
    								horast = horan + 0.5
							End If
							If minutos > "44" And minutos <= "59" Then
    								horast = horan + 1
							End If
						end if%>
						<td style="font-family: Verdana;"><center><h6><%=horast%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("quellevaorigen")%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("quellevadestino")%></h6></center></td>
                				<td style="font-family: Verdana;"><center><h6><%=rs2("distancia")%></h6></center></td>						
						<td style="font-family: Verdana;"><center><h6><%=rs2("horallegadareal")%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("horasalidareal")%></h6></center></td>
						<td style="font-family: Verdana;"><center><h6><%=rs2("observacionesconductor")%></h6></center></td>
						</tr>
						<%kms=kms+rs2("distancia")%>
						<%rs2.movenext
					loop
					cad="select * from hojarutamr where numviaje=" & rs1("numviaje") & " order by idhojaruta"
					set rs5=cn.execute(cad)
					if rs5.eof=false then%>
						<tr>
							<td colspan="16" align="center" style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">
							Destinos Internos de Descarga</td>
						</tr>
						<%do while rs5.eof=false%>
							<tr bgcolor="#29D1F6">
								<td style="font-family: Verdana;"><h6><%=rs5("fechasalida")%></h6></td>
								<td style="font-family: Verdana;"><center><h6><%=rs5("hojaderuta")%></h6></center></td>
								<td style="font-family: Verdana;"><h6><%=rs5("remitente")%></h6></td>
								<td style="font-family: Verdana;"><h6><%=rs5("destinatario")%></h6></td>
								<%cade="select m.patente from viajes v, medios m where v.idmedio=m.idmedio and v.numviaje=" & rs1("numviaje")
								set rs7=cn.execute(cade)
								patente=trim(rs7("patente"))
								rs7.close%>
								<td style="font-family: Verdana;"><h6><%=PATENTE%></h6></td>
								<td style="font-family: Verdana;" bgcolor="#c39bd3"><center><h6><%=rs5("fechallegada")& " " & trim(rs5("horallegada"))%></h6></center></td>
								<td style="font-family: Verdana;"><center><h6>&nbsp;</h6></center></td>
								<td style="font-family: Verdana;" bgcolor="#ccd1d1"><center><h6><%=rs5("fechasalida") & " " & trim(rs5("horasalida"))%></h6></center></td>
            							<%horast=0%>
								<%If not IsNull(RS5("horasalida")) and not IsNull(RS5("horallegada")) Then
                    							horast = DateDiff("n", RS5("horallegada"), RS5("horasalida")) 
									If horast >= 60 Then
    										horan = Int(horast / 60)
									Else
    										horan = 0
									End If
									minutos = horast - (horan * 60)
									If minutos < 15 Then
    										horast = horan
									End If
									If minutos >= "15" And minutos <= "44" Then
    										horast = horan + 0.5
									End If
									If minutos > "44" And minutos <= "59" Then
    										horast = horan + 1
									End If
								end if%>
								<td style="font-family: Verdana;"><center><h6>&nbsp;</h6></center></td>
								<td style="font-family: Verdana;"><center><h6><%=horast%></h6></center></td>
								<td style="font-family: Verdana;"><center><h6><%=rs5("quellevaorigen")%></h6></center></td>
								<td style="font-family: Verdana;"><center><h6><%=rs5("quellevadestino")%></h6></center></td>
	                					<td style="font-family: Verdana;"><center><h6>0</h6></center></td>
							</tr>
							<%rs5.movenext
						loop
					end if
					rs5.close
				end if
				rs2.close
				rs1.movenext%>
				<tr><td colspan="16"><hr></td></tr>
			<%loop
			rs1.close
			rs4.movenext
		loop%>
	</table>
	<br>
	<hr>
	<%if kms>0 then%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de KMS:&nbsp;<%=kms%></span></big></big>
	<%else%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de KMS:&nbsp;0</span></big></big>
	<%end if%>
	<hr>
<%else%>

	<%mes=month(request.querystring("desde"))%>
	<%if mes<10 then
		mes="0" & month(request.querystring("desde"))
	end if%>
	<%dia=day(request.querystring("desde"))%>
	<%if dia <10 then
		dia="0" & day (request.querystring("desde"))
	end if%>					
	<%anio=year(request.querystring("desde"))%>
	<%cadenadesde=dia & "/" & mes & "/" & anio%>
	<%mes=month(request.querystring("hasta"))%>
	<%if mes<10 then
		mes="0" & month(request.querystring("hasta"))
	end if%>
	<%dia=day(request.querystring("hasta"))%>
	<%if dia <10 then
		dia="0" & day (request.querystring("hasta"))
	end if%>					
	<%anio=year(request.querystring("hasta"))%>
	<%cadenahasta=dia & "/" & mes & "/" & anio%>
	<br>
	<br>
	<center>
	<table>
		<tr>
			<td colspan="10" align="center"><font face="verdana" size="6"><u>Consulta de Recorridos Desde:&nbsp; <%=cadenadesde%>&nbsp;&nbsp;&nbsp;Hasta:&nbsp;<%=cadenahasta%></u></font>
		</td>
		</tr>

		<tr>
			<td colspan="10" align="center"><font face="verdana" size="6" color="blue">MILK RUN</font></td>
		</tr>
		<tr>
			<td colspan="10" align="center"><font face="verdana" size="6" color="blue">Todos los Conductores</font></td>
		</tr>

		<tr>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">N° Hoja de Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Conductor</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Patente</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Origen</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Destino</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Prevista Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Fecha y Hora Prevista Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Horas</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Llegada</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Contenido Salida</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">KMS</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Llegada Real</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Hora Salida Real</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">Observaciones Conductor</td>
		</tr>
		<%totalhs=0%>
		<%cad="select * from viajes where fechasalida between '" & request.querystring("desde") & "' and '" & request.querystring("hasta") & "' AND web='" & "S" & "' and tipomov='" & "M" & "' and shipment is null order by fechasalida, idconductor"
		set rs1=cn.execute(cad)
		kms=0
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
					<%cade="select m.patente from viajes v, medios m where v.idmedio=m.idmedio and v.numviaje=" & rs1("numviaje")
					set rs7=cn.execute(cade)
					patente=trim(rs7("patente"))
					rs7.close%>
					<td style="font-family: Verdana;"><h6><%=PATENTE%></h6></td>
					<td style="font-family: Verdana;"><h6><%=rs2("remitente")%></h6></td>
					<td style="font-family: Verdana;"><h6><%=rs2("destinatario")%></h6></td>
					<td style="font-family: Verdana;" bgcolor="#c39bd3"><center><h6><%=rs2("fechallegada")& " " & trim(rs2("horallegada"))%></h6></center></td>
					<td style="font-family: Verdana;" bgcolor="#c39bd3"><center><h6><%=rs2("fechallegadarasa")& " " & trim(rs2("horallegadarasa"))%></h6></center></td>
					<td style="font-family: Verdana;" bgcolor="#ccd1d1"><center><h6><%=rs2("fechasalida") & " " & trim(rs2("horasalida"))%></h6></center></td>
					<td style="font-family: Verdana;" bgcolor="#ccd1d1"><center><h6><%=rs2("fechasalidarasa") & " " & trim(rs2("horasalidarasa"))%></h6></center></td> 
          				<%horast=0%>
					<%If not IsNull(RS2("horasalida")) and not IsNull(RS2("horallegada")) Then
                    				horast = DateDiff("n", RS2("horallegada"), RS2("horasalida")) 
						If horast >= 60 Then
    							horan = Int(horast / 60)
						Else
    							horan = 0
						End If
						minutos = horast - (horan * 60)
						If minutos < 15 Then
    							horast = horan
						End If
						If minutos >= "15" And minutos <= "44" Then
    							horast = horan + 0.5
						End If
						If minutos > "44" And minutos <= "59" Then
    							horast = horan + 1
						End If
					end if%>
					<td style="font-family: Verdana;"><center><h6><%=horast%></h6></center></td>
					<td style="font-family: Verdana;"><center><h6><%=rs2("quellevaorigen")%></h6></center></td>
					<td style="font-family: Verdana;"><center><h6><%=rs2("quellevadestino")%></h6></center></td>
                			<td style="font-family: Verdana;"><center><h6><%=rs2("distancia")%></h6></center></td>
					<td style="font-family: Verdana;"><center><h6><%=rs2("horallegadareal")%></h6></center></td>
					<td style="font-family: Verdana;"><center><h6><%=rs2("horasalidareal")%></h6></center></td>
					<td style="font-family: Verdana;"><center><h6><%=rs2("observacionesconductor")%></h6></center></td>
					</tr>
					<%kms=kms+rs2("distancia")
					rs2.movenext
				loop
				cad="select * from hojarutamr where numviaje=" & rs1("numviaje") & " order by idhojaruta"
				set rs5=cn.execute(cad)
				if rs5.eof=false then%>
					<tr>
						<td colspan="16" align="center" style="vertical-align: top; font-family: Verdana; font-weight: bold; color: red; background-color: rgb(255, 204, 153);">
						Destinos Internos de Descarga</td>
					</tr>
					<%do while rs5.eof=false%>
						<tr bgcolor="#29D1F6">
							<td style="font-family: Verdana;"><h6><%=rs5("fechasalida")%></h6></td>
							<td style="font-family: Verdana;"><center><h6><%=rs5("hojaderuta")%></h6></center></td>
							<%cad6="select nombre from conductores where idconductor=" & rs1("idconductor")
							set rs6=cn.execute(cad6)
							conductor=rs6("nombre")%>
							<td style="font-family: Verdana;"><center><h6><%=conductor%></h6></center></td>
							<%cade="select m.patente from viajes v, medios m where v.idmedio=m.idmedio and v.numviaje=" & rs1("numviaje")
							set rs7=cn.execute(cade)
							patente=trim(rs7("patente"))
							rs7.close%>
							<td style="font-family: Verdana;"><h6><%=PATENTE%></h6></td>
							<td style="font-family: Verdana;"><h6><%=rs5("remitente")%></h6></td>
							<td style="font-family: Verdana;"><h6><%=rs5("destinatario")%></h6></td>
							<td style="font-family: Verdana;" bgcolor="#c39bd3"><center><h6><%=rs5("fechallegada")& " " & trim(rs5("horallegada"))%></h6></center></td>
							<td style="font-family: Verdana;"><center><h6>&nbsp;</h6></center></td>							
							<td style="font-family: Verdana;" bgcolor="#ccd1d1"><center><h6><%=rs5("fechasalida") & " " & trim(rs5("horasalida"))%></h6></center></td>
           						<%horast=0%>
							<%If not IsNull(RS5("horasalida")) and not IsNull(RS5("horallegada")) Then
                    						horast = DateDiff("n", RS5("horallegada"), RS5("horasalida")) 
								If horast >= 60 Then
    									horan = Int(horast / 60)
								Else
    									horan = 0
								End If
								minutos = horast - (horan * 60)
								If minutos < 15 Then
    									horast = horan
								End If
								If minutos >= "15" And minutos <= "44" Then
    									horast = horan + 0.5
								End If
								If minutos > "44" And minutos <= "59" Then
    									horast = horan + 1
								End If
							end if%>
							<td style="font-family: Verdana;"><center><h6>&nbsp;</h6></center></td>
							<td style="font-family: Verdana;"><center><h6><%=horast%></h6></center></td>
							<td style="font-family: Verdana;"><center><h6><%=rs5("quellevaorigen")%></h6></center></td>
							<td style="font-family: Verdana;"><center><h6><%=rs5("quellevadestino")%></h6></center></td>
                					<td style="font-family: Verdana;"><center><h6>0</h6></center></td>
						</tr>
						<%rs5.movenext
					loop
				end if
				rs5.close
			end if
			rs2.close
			rs1.movenext%>
			<tr><td colspan="16"><hr></td></tr>
		<%loop%>
	</table>
	<br>
	<hr>
	<%if kms>0 then%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de KMS:&nbsp;<%=kms%></span></big></big>
	<%else%>
		<center><big><big><span style="font-family: Verdana; color: rgb(51, 51, 255);">Total
		de KMS:&nbsp;0</span></big></big>
	<%end if%>
	<hr>
<%end if%>