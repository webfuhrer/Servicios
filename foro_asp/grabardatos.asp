

<%
Dim nombre, comentario
nombre = Session("nombre")
comentario = Request.Form("comentario")
Dim fs, file, csvData
Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set file = fs.OpenTextFile(Server.MapPath("data.csv"), 8, True)

csvData = nombre & "," & comentario & vbCrLf 
file.Write csvData

file.Close
Set file = Nothing
Set fs = Nothing
Response.Redirect "recuperadatos.asp"
%>