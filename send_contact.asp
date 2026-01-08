<%
Response.ContentType = "application/json"
Response.Charset = "utf-8"

' Leer campos
Dim nombre, apellido, email, telefono, mensaje
nombre = Trim(Request.Form("nombre"))
apellido = Trim(Request.Form("apellido"))
email = Trim(Request.Form("email"))
telefono = Trim(Request.Form("telefono"))
mensaje = Trim(Request.Form("mensaje"))

' Validar
If nombre = "" Or email = "" Or mensaje = "" Then
  Response.Status = "400 Bad Request"
  Response.Write "{""success"":false,""error"":""Faltan campos requeridos""}"
  Response.End
End If

' Intentar enviar email via CDO
On Error Resume Next
Dim objMsg
Set objMsg = Server.CreateObject("CDO.Message")

' Ajustar según el servidor SMTP disponible. Por defecto intenta localhost:25
objMsg.From = email
objMsg.To = "info@groupecat.com.ar"
objMsg.Subject = "Contacto web: " & nombre & " " & apellido

Dim bodyText
bodyText = "Nombre: " & nombre & vbCrLf & _
           "Apellido: " & apellido & vbCrLf & _
           "Email: " & email & vbCrLf & _
           "Teléfono: " & telefono & vbCrLf & _
           "Mensaje: " & vbCrLf & mensaje

objMsg.TextBody = bodyText

' Configuración CDO para SMTP local
Dim cfg
Set cfg = objMsg.Configuration
cfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
cfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
cfg.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
cfg.Fields.Update

objMsg.Send

If Err.Number <> 0 Then
  Dim errMsg
  errMsg = Replace(Err.Description, """", "'")
  Response.Status = "500 Internal Server Error"
  Response.Write "{""success"":false,""error"":""Error al enviar el email: " & errMsg & ""}" 
Else
  Response.Write "{""success"":true}"
End If

Set objMsg = Nothing
Set cfg = Nothing
%>
