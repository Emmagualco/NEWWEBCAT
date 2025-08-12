<%usuario=request.querystring("usuario")
clave=request.querystring("clave")
set cn=server.createobject("ADODB.Connection")
set rs=server.createobject("ADODB.Recordset")
set rs1=server.createobject("ADODB.Recordset")
set rs2=server.createobject("ADODB.Recordset")
set rs3=server.createobject("ADODB.Recordset")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
cad="Driver={SQL Server};Persist SecurityInfo=True;Description=sqldemo;SERVER=172.24.206.1;UID=sa;PWD=bfg2007;DATABASE=LCARG"
cn.open(cad)
rs3.CursorType = adOpenKeyset
rs3.LockType = adLockOptimistic
rs3.open"select * from viajes where numviaje=" & request.querystring("hr"),cn
rs3("fechainiciovalidada")=request.form("fechallegada")
rs3("fechafinvalidada")=request.form("fechasalida")
rs3("horainiciovalidada")=request.form("hora1")
rs3("horafinvalidada")=request.form("hora2")
rs3("fechallegada")=request.form("fechasalida")
rs3("fechasalida")=request.form("fechallegada")
rs3("horallegada")=request.form("hora2")
rs3("horasalida")=request.form("hora1")
rs3("fechainiciovalidacion")=date()
rs3("fechafinvalidacion")=date()
rs3("horainiciovalidacion")=time()
rs3("horafinvalidacion")=time()
rs3("usuariovalidainicio")=trim(request.querystring("usuario"))
rs3("usuariovalidafin")=trim(request.querystring("usuario"))
rs3("observacionesvalidacion")=ucase(trim(request.form("obser")))
rs3.update
rs3.close
response.redirect("listadeconductores1.asp?usuario="&usuario & "&clave=" & clave)%>





