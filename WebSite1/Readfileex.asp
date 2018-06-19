<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title></title>
</head>
<body>
    <% 

Filename = "D:\VSREPO\bennytest\readme.txt"    ' file to read
 'ForReading = 1, ForWriting = 2, ForAppending = 3
 'TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        
' Create a filesystem object
 '       Const ForReading = 1, ForWriting = 2, ForAppending = 3
'Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

Dim lineData
Set fso = Server.CreateObject("Scripting.FileSystemObject") 
set fs = fso.OpenTextFile(Filename , 1, true) 
Do Until fs.AtEndOfStream 
    lineData = fs.ReadLine
    'do some parsing on lineData to get image data
    'output parsed data to screen
    'Response.Write lineData
        Response.write "<pre>" &lineData & "</pre><hr>"
Loop 

fs.close: set fs = nothing 
%>
    <%
        dim filesys, filetxt
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set filesys = CreateObject("Scripting.FileSystemObject")
Set filetxt = filesys.OpenTextFile("D:\VSREPO\bennytest\somefile.txt", ForAppending, True)
filetxt.WriteLine("Your text goes here.")
filetxt.Close 

        %>
</body>
</html>