


<cfparam name="url.action" default="">


<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
	<title>Word Extractor</title>



		
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	
	<link rel="start"  	href="index.cfm" />
	<link rel="made" 	href="mailto:james@webworldinc.com" />
</head>	


<body>




<h1>Uploaded New Files</h1>



<form action="index.cfm?action=extract" method="post" enctype="multipart/form-data">


<input type="file" name="wordDocx" />

<br />
<button type="submit" style="height : 50px; width : 100px; font-size : 20pt;" >Upload</button>
</form>  


<cfif url.action EQ "extract">

	<h4>Results</h4>

	<cfscript>
	objExtract = new wordextractor();
	
	strPath = GetDirectoryFromPath(GetBaseTemplatePath() ) & "temp/";
	if (!DirectoryExists(strPath)) {
		DirectoryCreate(strPath);
		}

	fileInfo = FileUpload(strPath, "wordDocx", "application/*", "overwrite");

	writedump(fileInfo);

	result = objExtract.extractDocx("#strPath##fileinfo.serverfile#");
	</cfscript>
	
	
	<cfoutput>
		<p>Raw length: #len(objExtract.xmlString)#</p>
		<p>Extracted length: #len(result)#</p>
		<hr />
		#result#
		<hr />
		<pre>#encodeForXML(result)#</pre>
	</cfoutput>


</cfif>






</body>
</html>

