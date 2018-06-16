<!DOCTYPE html>
<html>
  <head>
<title>UT FREIGHT SERVICE LTD. - Back Office System Login Page </title>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" title="Style" href="include/Main.css">
</head>
<body>
    <%
        
	' Start Generate Excel #################################################################
	Set objExcel=CreateObject("Excel.Application")
	objExcel.DisplayAlerts=False	' 不顯示警告
	ObjExcel.Visible=False			' 不顯示介面
	
	Set objWB = objExcel.Workbooks.Add
	Set objWS = objWB.Worksheets(1)
	Set colSheets = objWB.Sheets
	
	objExcel.Workbooks(1).WorkSheets(1).Name = "運況暨庫存資料"
	
	With objExcel.Range("A1:V1")
		.ColumnWidth = 15.0
		.Font.Size = 10
		.HorizontalAlignment = -4108
		.Interior.ColorIndex = 44
	End With
	
	objExcel.cells(1,1).value = "No."
	objExcel.cells(1,2).value = "會簽單號"
	objExcel.cells(1,3).value = "出貨單號"
	objExcel.cells(1,4).value = "出貨日期"
	objExcel.cells(1,5).value = "提單號碼"
	objExcel.cells(1,6).value = "運輸" & chr(10) & "方式"
	objExcel.cells(1,7).value = "項目"
	objExcel.cells(1,8).value = "母件物料號碼"
	objExcel.cells(1,9).value = "母件批號"
	objExcel.cells(1,10).value = "子件物料號碼"
	objExcel.cells(1,11).value = "子件批號"
	objExcel.cells(1,12).value = "產品名稱"
	objExcel.cells(1,13).value = "狀態"
	objExcel.cells(1,14).value = "儲位"
	objExcel.cells(1,15).value = "出口日期"
	objExcel.cells(1,16).value = "出口" & chr(10) & "數量"
	objExcel.cells(1,17).value = "入庫日期"
	objExcel.cells(1,18).value = "入庫" & chr(10) & "數量"
	objExcel.cells(1,19).value = "出貨日期"
	objExcel.cells(1,20).value = "出貨" & chr(10) & "數量"
	objExcel.cells(1,21).value = "發票號碼"
        objExcel.Range("A" & i & ":A" & i ).ColumnWidth = 4
		objExcel.Range("B" & i & ":B" & i ).ColumnWidth = 11
		objExcel.Range("C" & i & ":C" & i ).ColumnWidth = 9
		objExcel.Range("D" & i & ":D" & i ).ColumnWidth = 10
		objExcel.Range("E" & i & ":E" & i ).ColumnWidth = 15
		objExcel.Range("F" & i & ":F" & i ).ColumnWidth = 5
		objExcel.Range("G" & i & ":G" & i ).ColumnWidth = 5
		objExcel.Range("H" & i & ":H" & i ).ColumnWidth = 18
		objExcel.Range("I" & i & ":I" & i ).ColumnWidth = 11
		objExcel.Range("J" & i & ":J" & i ).ColumnWidth = 15
		objExcel.Range("K" & i & ":K" & i ).ColumnWidth = 11
		objExcel.Range("L" & i & ":L" & i ).ColumnWidth = 20
		objExcel.Range("M" & i & ":M" & i ).ColumnWidth = 7
		objExcel.Range("N" & i & ":N" & i ).ColumnWidth = 8
		objExcel.Range("O" & i & ":O" & i ).ColumnWidth = 20
		objExcel.Range("P" & i & ":P" & i ).ColumnWidth = 6
		objExcel.Range("Q" & i & ":Q" & i ).ColumnWidth = 20
		objExcel.Range("R" & i & ":R" & i ).ColumnWidth = 6
		objExcel.Range("S" & i & ":S" & i ).ColumnWidth = 20
		objExcel.Range("T" & i & ":T" & i ).ColumnWidth = 6
		objExcel.Range("U" & i & ":U" & i ).ColumnWidth = 15
	'With objExcel.Range("A1:U" & i)
	'	.Font.Size = 10
	'	.BORDERS.Weight = 1
	'End With
        
	'FilePath = BL_AES_PATH & "庫存動態_" & DateX & ".xls"
        FilePath ="d:\utweb\test.xls"
	objWB.SaveAs(FilePath)
	objWB.Save
	
	objExcel.Application.Quit
	Set objExcel = Nothing
        response.Write("<p>Done!</p>")
        response.redirect (FilePath)
        %>
</body>
</html>