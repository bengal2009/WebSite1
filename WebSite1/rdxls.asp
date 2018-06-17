<!DOCTYPE html>
<html>
  <head>
<title>UT FREIGHT SERVICE LTD. - Back Office System Login Page </title>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" title="Style" href="include/Main.css">
</head>
<body>
    <%
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")

DC_PATH = Server.MapPath("test.xls") 
        response.Write(DC_PATH+"<br>")
        Set objConn = Server.CreateObject("ADODB.Connection")
	'objConn.Open "Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; DBQ=" & DC_PATH & request("filename") & ";"
        Constr="Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; DBQ=" & DC_PATH & ";"
        response.Write (Constr+"<br>")
        objConn.Open Constr
        'rs.open "SELECT * FROM [Sheet1$A1:AD10]", objConn
        rs.open "SELECT * FROM [Sheet1$]", objConn
        i=1
        while not rs.eof
		response.Write("<p>"+cstr(i)+":"+rs(0)+"</p>")
        i=i+1
        rs.movenext
        
	wend
        set objConn=nothing
        %>
</body>
</html>