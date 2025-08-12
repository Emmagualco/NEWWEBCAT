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
Response.AddHeader "content-disposition", "inline; filename=RGT.xls"
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
				<big style="color: rgb(0, 0, 0);"><big><span style="font-family: Helvetica,Arial,sans-serif;"><u>Consulta RGT Desde:&nbsp; <%=cadenadesde%>&nbsp;&nbsp;&nbsp;Hasta:&nbsp;<%=cadenahasta%></u></span></big></big>
	<center>
	<table>
		<tr>
			<td colspan="29">&nbsp;</td>
		</tr>
	</table>
	<table border="1">
		<tr align="center">
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Fecha</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Dominio</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Tipo Camión</td>	
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Km Ruta</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Km MR</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Unidad</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Origen</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(248, 252, 158);">ML Tot</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">ML R</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Capac. R</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Carga R</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(75, 251, 248);">TLLR</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">ML N</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Capac. N</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(125, 221, 219);">Carga N</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(75, 251, 248);">TLLN</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(75, 251, 248);">Circuito</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Proveedor</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Tipo Flujo</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Cargo</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Observac.</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Incidente RASA</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: black; background-color: rgb(239, 250, 4);">Incidente NARG</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Otros RASA</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Otros NARG</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Rep. RASA</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">Rep. UTE</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">ILN</td>
			<td style="vertical-align: top; font-family: Verdana; font-weight: bold; color: white; background-color: rgb(162, 56, 56);">No Prod.</td>

		</tr>
		<%totalhs=0%>

		<%cade="select * from rgtdef where fecha between '" & request.querystring("desde")  & "' and '" & request.querystring("hasta") & "' order by numviaje"
		set rs=cn.execute(cade)
		do while rs.eof=false%>
			<tr bgcolor="white" align="center">
				<td><font face="verdana" size="2"><%=rs("fecha")%></font></td>
				<td><font face="verdana" size="2"><%=rs("dominio")%></font></td>
				<td><font face="verdana" size="2"><%=rs("tipocamion")%></font></td>
				<td><font face="verdana" size="2"><%=rs("kmruta")%></font></td>
				<td><font face="verdana" size="2"><%=rs("kmmr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("unidad")%></font></td>
				<td><font face="verdana" size="2"><%=rs("origen")%></font></td>
				<td><font face="verdana" size="2"><%=rs("mltot")%></font></td>
				<td><font face="verdana" size="2"><%=rs("mlr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("capacr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("cargar")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("tllr")%></font></td>
				<td><font face="verdana" size="2"><%=rs("mln")%></font></td>					
				<td><font face="verdana" size="2"><%=rs("capacn")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("cargan")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("tlln")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("circuito")%></font></td>
				<td><font face="verdana" size="2"><%=rs("proveedor")%></font></td>
				<td><font face="verdana" size="2"><%=rs("tipoflujo")%></font></td>
				<td><font face="verdana" size="2"><%=rs("cargo")%></font></td>
				<td><font face="verdana" size="2"><%=rs("observac")%></font></td>
				<td><font face="verdana" size="2"><%=rs("incidenterasa")%></font></td>
				<td><font face="verdana" size="2"><%=rs("incidentenarg")%></font></td>
				<td><font face="verdana" size="2"><%=rs("otrosrasa")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("otrosnarg")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("reprasa")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("repute")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("iln")%></font></td>				
				<td><font face="verdana" size="2"><%=rs("noprod")%></font></td>
			</tr>
			<%rs.movenext
		loop%>
	</table>
