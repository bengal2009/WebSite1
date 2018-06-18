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
        'response.Write(DC_PATH+"<br>")
        %>
    <form action="rdxls.asp" method="post" name="form1">
    <table align=center border=1 cellPadding=1 cellSpacing=0 width="700" bordercolor=#d3d3d3 bordercolordark=#dcdcdc bordercolorlight=#ffffff>
<tr class="title1" height="28" align="center">
	<td WIDTH="5%">No.</td>
	<td WIDTH="45%">P/O File Name</td>
	<td WIDTH="28%">P/O Create Date</td>
	<td WIDTH="12%">File Size</td>
	<td WIDTH="10%">Import</td>
</tr>
    <%
        Set objConn = Server.CreateObject("ADODB.Connection")
	'objConn.Open "Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; DBQ=" & DC_PATH & request("filename") & ";"
        Constr="Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; DBQ=" & DC_PATH & ";"
        response.Write (Constr+"<br>")
        objConn.Open Constr
        'rs.open "SELECT * FROM [Sheet1$A1:AD10]", objConn
        rs.open "SELECT * FROM [Sheet1$]", objConn
        i=1
        while not rs.eof
        %>
        <tr align="center" class="title1txt<%if i mod 2 then Response.Write "1" else Response.Write "2" end if%>" height="28">
	<td><%=i%>.</td>
	<td><%=rs(0)%></td>
	<td><%=rs(1) %></td>
	<td><%=rs(2)%>KB</td>
	<td>
	<%=rs(3)%>
	</td>
	</tr>
        <%
		'response.Write("<p>"+cstr(i)+":"+rs(0)+"</p>")
        i=i+1
        rs.movenext
        
	wend
            rs.Close
              rs.open "SELECT count(1) FROM [Sheet1$]", objConn
            
        set objConn=nothing
        %>
        </table>
    <%=rs(0) %>
        <input type="button" name="PREALERT" value="Upload Pre-Alert" class="resetb" style="width:150px;height:30px" onclick="OpenWin_Upload()">
        </form>
</body>
</html>
<SCRIPT LANGUAGE=javascript>
    function OpenWin_Upload() {
        /*if (form1.MBL_NOX.value=='' && form1.HBL_NOX.value=='') {
            alert('Please choose the HAWB Number!');
            return false;
        }
        varnewwin=window.open("UploadImg1.asp?MAWBX="+form1.MBL_NOX.value+"&HAWBX="+form1.HBL_NOX.value,"UploadImg","status=yes,width=600,height=400,top=10,left=10,scrollbars=yes")
        varnewwin.focus();
    }*/
        if (confirm('Are you sure to finish this process?')) {
            return true;
        }
        return false;
    }
    </script>