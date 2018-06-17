<!DOCTYPE html>
<html>
  <head>
<title>UT FREIGHT SERVICE LTD. - Back Office System Login Page </title>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" title="Style" href="include/Main.css">
</head>
<body>
    <%
'Set rs = Server.CreateObject("ADODB.Recordset")
'Set rs1 = Server.CreateObject("ADODB.Recordset")

DC_PATH = Server.MapPath("test.xls") 
        response.Write(DC_PATH)
        Set objConn = Server.CreateObject("ADODB.Connection")
	'objConn.Open "Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; DBQ=" & DC_PATH & request("filename") & ";"
        objConn.Open "Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; DBQ=" & DC_PATH & ";"
        'rs.open "SELECT * FROM [Sheet1$A1:AD10]", objConn
        rs.open "SELECT * FROM [Sheet1$A1:AD10]", objConn
        while not rs.eof
		if DO_NOX = "" then
			DO_NOX 	= rs.Fields(29).Value
		else
			DO_NOX 	= DO_NOX & ", " & rs.Fields(29).Value
		end if
		rs.movenext
	wend
        %>
</body>
</html>