


<cfparam name="url.action" default="">


<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
	<title>Word Extractor</title>



		
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	
	<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" rel="stylesheet" integrity="sha256-7s5uDGW3AHqw6xtJmNNtr+OBRJUlgkNJEo78P4b0yRw= sha512-nNo+yCHEyn0smMxSswnf/OnX6/KwJuZTlNZBjauKhTK0c+zT+q5JOCx0UFhXQ6rJR9jg6Es8gPuD2uZcYDLqSw==" crossorigin="anonymous">
	
	<link rel="start"  	href="index.cfm" />
	<link rel="made" 	href="mailto:james@webworldinc.com" />
</head>	


<body>


<div class="container">

<div class="jumbotron">
	<h1>Word Extractor</h1>
	<p>Get HTML out of a Word Document</p>
</div>



<form action="index.cfm?action=extract" method="post" enctype="multipart/form-data" class="form-inline">

	<div class="form-group">
		<input type="file" name="wordDocx" />
	</div>

	<button type="submit" class="btn btn-primary btn-large">Upload</button>
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
		<p>Raw length: #objExtract.xmlString.len()#</p>
		<p>Extracted length: #result.len()#</p>
		<hr />
		#result#
		<hr />
		<pre>#result.encodeForXML()#</pre>
	</cfoutput>


</cfif>



</div>

</body>
</html>



