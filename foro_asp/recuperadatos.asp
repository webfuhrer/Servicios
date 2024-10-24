<%
Dim nombre
If Request.Form("nombre") <> "" Then
    nombre = Request.Form("nombre")
    Session("nombre") = nombre
End If

If nombre <> "" Then
    Response.Write("Hola, " & nombre & "!<br>")

End If


Function ReadCSVFile(filePath)
    Dim fso, file, content
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set file = fso.OpenTextFile(Server.MapPath(filePath), 1)

    
    content = file.ReadAll
    
    file.Close
    Set file = Nothing
    Set fso = Nothing
    ReadCSVFile = content
End Function

Dim csvFilePath
csvFilePath = "data.csv"

Dim csvContent
csvContent = ReadCSVFile(csvFilePath)

Dim csvLines
csvLines = Split(csvContent, vbCrLf)

For Each line In csvLines
    Dim fields
    fields = Split(line, ",")

    If UBound(fields) >= 0 Then
        Response.Write(fields(0) & " dijo:" & "<br>" & fields(1) & "<br>")
    End If
    
    
Next
%>
<!DOCTYPE html>
<html>
<head>
    <title>My Web Page</title>
</head>
<body>
    <form action="grabardatos.asp" method="post">
        <textarea name="comentario"></textarea>
        <br>
        <input type="submit" value="Submit">
    </form>
</body>
</html>