<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title></title>
</head>
<body>
    <%
        Set objExcel = CreateObject("Excel.Application")
	        objExcel.Visible = True
            objExcel.DisplayAlerts = False
            with objExcel
                .workbooks.open("D:\VSREPO\test.xls")
                .ActiveSheet.Cells(2,1).Value = "1"
                .ActiveSheet.Cells(2,2).Value = "2"
                .ActiveSheet.Cells(2,3).Value = "3"
                .ActiveSheet.Cells(2,4).Value = "4"
                '.Save("C:\Inetpub\wwwroot\testExcel\Temp1.xls")
            end with
            'Set objWorkbook = objExcel.Workbooks.Open("D:\VSREPO\test.xls")
        'objExcel.workbook.Save("D:\VSREPO\test3.xls")
'objExcel.Workbooks.Close
objExcel.Save("D:\VSREPO\test3.xls")
objExcel.Quit

	'objExcel.Application.Quit
	Set objExcel = Nothing
	
        %>
</body>
</html>