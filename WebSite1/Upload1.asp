<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title></title>
</head>
<body>
    <form method="post" action="UploadImg2.asp" enctype="multipart/form-data" name=form1>
        <table border="0" width="100%" bgcolor="#EEEEEE" cellpadding=0 cellspacing=0>
		<tr class="title1">
			<td align=center height=25>File 1: <input type="file" name="File1" size="60" class="inputx" accept="image/gif, image/jpeg"></td>
		</tr>
		<tr class="title1">
			<td align=center height=25>File 2: <input type="file" name="File2" size="60" class="inputx"></td>
		</tr>
		<tr class="title1">
			<td align=center height=25>File 3: <input type="file" name="File3" size="60" class="inputx"></td>
		</tr>
		<tr>
			<td align=center>
				<input type="submit" value="File Upload" class="greenb" name="Submit" style="width:120px;height:22px" onclick="return DataCheck();">
			</td>
		</tr>
		</table>
        </form>
</body>
</html>
<script language=javascript>
<!--
Status = "<%=request("Status")%>";
if (Status=='correct') {	
	alert('The file(s) was/were been uploaded to the PMS server.\nPlease update what the type of this Pre-Alert.');
}
else if (Status=='large') {
	alert('Sorry, your picture file size is too large(under 200 kb).');
}
else if (Status=='over') {
	alert('Sorry, one component limits 5 maximum picture files.');
}

function DataCheck() {
	if (form1.File1.value=='' && form1.File2.value=='' && form1.File3.value=='') {
		alert('Sorry，please select at least one file to upload!');
		return false;
	}
	return true;
}

function SendCheck() {
	if(confirm('Are you sure to finish and send the notice?')) {
		//alert(form2.RMKX.value);
		strOrg = form2.RMKX.value;
		strOrg = strOrg.replace(/\r\n/g,"@!@");
		window.location = "UPloadImg1.asp?HAWBX=<%=request("HAWBX")%>&SendNotice=true&RMKX="+strOrg;
	}
	return false;
}
//-->
</script>