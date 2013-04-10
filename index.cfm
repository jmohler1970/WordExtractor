<!---
   Copyright 2011 James Mohler

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
--->   

<cfparam name="url.action" default="">


<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
	<title>Word ExtractorFoundation</title>

	<style type="text/css">
		@import url(main.css);
	</style>

		
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	
	<link rel="start"  	href="index.cfm" />
	<link rel="made" 	href="mailto:james@webworldinc.com" />
</head>	


<body>




<h4>Uploaded New Files</h4>



<form action="index.cfm?action=extract" method="post" enctype="multipart/form-data">


<input type="file" name="wordDocx" />

<br />
<button type="submit" style="height : 50px; width : 100px; font-size : 20pt;" >Upload</button>
</form>  


<cfif url.action EQ "extract">

	<h4>Results</h4>

	<cfset objExtract = createobject("component", "wordextractor")>
	
	<cfset strPath = GetDirectoryFromPath(GetBaseTemplatePath() ) & "temp/">

	<cffile action="upload" fileField="wordDocx" destination="#strPath#" nameConflict="Overwrite">


	<cfset result = objExtract.extractDocx("#strPath##cffile.serverfile#")>
	
	<cfoutput>
		<p>Raw length: #len(objExtract.xmlString)#</p>
		<p>Extracted length: #len(result)#</p>
		<hr />
		#result#
		<hr />
		<pre>#htmleditformat(result)#</pre>
	</cfoutput>


</cfif>






</body>
</html>

