<!DOCTYPE html>
<html>
<head>
    
    <META content="text/html; charset=big5" http-equiv=Content-Type>
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
FilePath="d:\temp\"
Set fds = Server.CreateObject("Scripting.FileSystemObject")
if Not fds.FolderExists(FilePath) then fds.CreateFolder(FilePath)
set fds = nothing

Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

mySmartUpload.AllowedFilesList = "jpg,gif,tif,pdf,xls,doc"
mySmartUpload.DeniedFilesList = "com,exe,bat,asp"

mySmartUpload.MaxFileSize = 10485760
mySmartUpload.TotalMaxFileSize = 104857600

mySmartUpload.Upload
        For each file In mySmartUpload.Files
	Set fds = Server.CreateObject("Scripting.FileSystemObject")
	
	Dim DateX, MonthX, DayX, HourX, MinuteX, SecondX
	DateX = ""
	'DateX = right(year(now()),2)
	'MonthX = right(month(now())+100,2)
	'DayX = right(day(now())+100,2)
	HourX = right(hour(now())+100,2)
	MinuteX = right(minute(now())+100,2)
	SecondX = right(second(now())+100,2)
	
	'DateX = DateX & MonthX & DayX & HourX & MinuteX & SecondX & MS
	DateX = HourX & MinuteX & SecondX & MS
	
	FileNameX = replace(file.FileName,"." & file.FileExt,"") & "_" & DateX & "." & file.FileExt
	set fds = nothing
	
	If not file.IsMissing Then
		file.SaveAs(FilePath & FileNameX)
        intCount = intCount + 1
	End If
	
	If not file.IsMissing Then
	
		
		'rs.open "Select UNID from UPLOAD_FILE Order by UNID DESC", ConnWKF
		'if not rs.eof then
		'	UNID = rs("UNID")
		'end if
		'rs.close
		
		if Str_FName = "" then
			Str_FName =  FilePath & FileNameX
		else
			Str_FName =  Str_FName & "、" & FilePath & FileNameX
		end if
	end if
Next

        %>
</body>
</html>