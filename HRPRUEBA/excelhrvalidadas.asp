<%@ Language=VBScript %> 
<%dim cn
dim rs
dim rs1
dim cad
dim cad1
dim canti
set cn=server.createobject("adodb.connection")
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")%>
<%
Session.LCID = 11274
%> 
<%cad="Driver={SQL Server};Persist Security Info=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"							
cn.open(cad)
cad1="select * from hrvalidadas order by numviaje"
set rs=cn.execute(cad1)
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename=Hojas de Ruta Validadas.xls"%> 

<TABLE width="100" BORDER="0"> 

	<TR>
		<td colspan="10"><div align="center"><font color="#0066CC" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong><em>Consulta de Hojas de Ruta Validadas</em></strong></font></div></td>
	</tr>
	<TR>
		<td colspan="10"><div align="center"><font color="#0066CC" size="5" face="Verdana, Arial, Helvetica, sans-serif"><strong><em>Desde:<%=request.querystring("desde")%>&nbsp;&nbsp;&nbsp;Hasta:<%=request.querystring("hasta")%></em></strong></font></div></td>
	</tr>
</table>
<br>
<br>
<TABLE width="100" BORDER="1"> 
	<tr>	
		<%For i = 0 to rs.Fields.Count - 1 %>
				<%if rs(i).Name="Fecha" then
					titulo="Fecha H.R."
				end if
				if rs(i).Name="Numviaje" then
					titulo="N° Hoja de Ruta"
				end if
				if rs(i).Name="FechaInicio" then
					titulo="Fecha Inicio"
				end if
				if rs(i).Name="HoraInicio" then
					titulo="Hora Inicio"
				end if
				if rs(i).Name="FechaFin" then
					titulo="Fecha Fin"
				end if
				if rs(i).Name="HoraFin" then
					titulo="Hora Fin"
				end if
				if rs(i).Name="Horas" then
					titulo="Horas"
				end if
				if rs(i).Name="Patente" then
					titulo="Patente"
				end if
				if rs(i).Name="Conductor" then
					titulo="Conductor"
				end if
				if rs(i).Name="ValidoInicio" then
					titulo="Usuario Validó Inicio"
				end if
				if rs(i).Name="ValidoFin" then
					titulo="Usuario Validó Fin"
				end if%>

				<TD style="vertical-align: top; background-color: rgb(51, 102, 255); color: white; font-family: Verdana; text-align: center;"><B><% = titulo %></B></TD> 
		<%Next %> 
	</TR> 
	<%canti=0%>
	<%horas=0%>
	<%color1="#ECF7C0"%>
	<%color2="#E5DDD0"%>
	<%color=color1%>
	<%Do While Not rs.EOF %> 
		<%if color=color1 then
			color=color2
		else
			color=color1
		end if%>
		<tr bgcolor="<%=color%>"> 
			<% For i = 0 to rs.Fields.Count - 1 %>
				<TD style="vertical-align: top; font-family: Verdana; color: black; font-size:80%;"><%=rs(i)%></TD>	
				<%if i=6 then
					horas=horas+rs(i)
				end if%>  
			<%Next %> 
		</TR> 
		<%canti=canti+1
		rs.MoveNext 
	Loop 
	rs.Close%> 
</TABLE>
<br>
<hr>
<center><font color="#0066CC" size="5" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total de Hojas de Ruta Validadas:&nbsp;<%=canti%>&nbsp;&nbsp;&nbsp;&nbsp;Cantidad de Horas:&nbsp;<%=horas%></font></center>
<hr>