<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title></title>
</head>
<body>
    <%
        
Dim FOLDERX
FOLDERX = year(date) & right(month(now())+100,2)
        
Dim mySmartUpload
Dim file
Dim intCount
intCount=0
FilePath="d:\temp"
Set fds = Server.CreateObject("Scripting.FileSystemObject")
if Not fds.FolderExists(FilePath) then fds.CreateFolder(FilePath)
set fds = nothing

Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

mySmartUpload.AllowedFilesList = "jpg,gif,tif,pdf,xls,doc"
mySmartUpload.DeniedFilesList = "com,exe,bat,asp"

mySmartUpload.MaxFileSize = 10485760
mySmartUpload.TotalMaxFileSize = 104857600

mySmartUpload.Upload
        %>
</body>
</html>