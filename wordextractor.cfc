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


<cfcomponent> 

<cfscript>
this.xmlPara = ""; // parsed into XML nodes
this.xmlString = ""; // raw text
this.proofErr = "spellEnd";
variables.listcounter = 1; // used to set order list numbers
variables.CRLF = Chr(13) & Chr(10);


package string function ReadNode (required xml Node) {
	
	var result = "";
	var wpPr = "";
	var wrPr = ""; // Does bold, italic
	var wpStyle = "";
	var wnumPr = ""; // ordered or unordered in html
	var wnumID = "";
	var startElementName = "";
	
	if (StructIsEmpty(arguments.Node))	{
		return "";
		}
		
		
    
    for (Element in arguments.node.xmlChildren)	{
    	
    	startwVal = "";
    	// Start Tags
    	switch (Element.xmlName)	{
    	
    		
    		case "w:p" :
    			wVal = ""; // default paragraph style
    			wnumid = ""; // This actually the type of list
    		   			
    			if (ArrayLen(Element.XMLChildren) != 0)	{
    			
	    			if (Element.XMLChildren[1].xmlName == "w:pPr")	{
	    				wpPr = Element.XMLChildren[1];
	    				    					    				
	    				if (wpPr.XMLChildren[1].xmlName == "w:pStyle")	{
	    					wpStyle = wpPr.XMLChildren[1];
	    						    				    					
	    					wVal = wpStyle.XMLAttributes["w:val"];
	    					}
	    				
	    				if (ArrayLen(wpPr.XMLChildren) >= 2) {
	    					
	    					if (wpPr.XMLChildren[2].xmlName == "w:numPr")	{
	    						wnumPr = wpPr.XMLChildren[2];
	    						if (wnumPr.XMLChildren[2].xmlName == "w:numID")	{
	    							wnumid = wnumPr.XMLChildren[2].xmlAttributes["w:val"];
	    							}
	    						}
	    					}			
	    				}
		    		}	
	    		
	    		switch (wVal)	{
	    			case "Heading1" :
	    				variables.listcounter = 1;
	    				result &= '<h1>#this.ReadNode(Element)#</h1>#variables.crlf#';
	    				break;
	    			
	    			case "Heading2" :
	    				variables.listcounter = 1;
	    				result &= '<h2>#this.ReadNode(Element)#</h2>#variables.crlf#';
	    				break;
	    				
	    			case "Heading3" :
	    				variables.listcounter = 1;
	    				result &= '<h3>#this.ReadNode(Element)#</h3>#variables.crlf#';
	    				break;
	    			
	    			case "Heading4" :
	    				variables.listcounter = 1;
	    				result &= '<h4>#this.ReadNode(Element)#</h4>#variables.crlf#';
	    				break;
					
					case "ListParagraph" :
						
						if (wnumid == 2)	{
							result &= '<ol start="#listcounter#"><li>#this.ReadNode(Element)#</li></ol>#variables.crlf#';
							}
						else	{
	    					result &= '<li>#this.ReadNode(Element)#</li>#variables.crlf#';
	    					}
	    					
	    				variables.listcounter++;
	
	    				break;


	    			   		
	    			default :
	    				variables.listcounter = 1;
	    				result &= '<p>#this.ReadNode(Element)#</p>#variables.crlf#';
	    			}
	    			
	    				
	    		break; // end of w:p
	   	
	    			
	    	case "w:r"	: // This handles bolds and italics
				wrPr = "";
				if (Element.XMLChildren[1].XMLName == "w:rPr")	{
					wrPr = Element.XMLChildren[1].XMLChildren[1].XMLName;
					}
					
				switch (wrPr)	{
					case "w:b" :
						result &= "<b>#this.ReadNode(Element)#</b>";
						break;
					
					case "w:i" :
						result &= "<i>#this.ReadNode(Element)#</i>";
						break;
					default :
						// Unknown
						result &= this.ReadNode(Element);
					}	
				break;	
				
	    			
	    	    		
    		case "w:t" :
				result &= Element.xmlText;
				break;
			
			
			case "w:ProofErr" :
				// Word divides this into separate areas
				// skip this.proofErr = Element.XMLAttributes["w:type"];
				break;
			
			case "w:pStyle" :
				// skip variables.currentTag = Element.XMLAttributes["w:val"];
				break;
			
			case "w:instrText" :
				// skip
				break;
			
			default :
				result &= Element.xmlText;
			 	
			}	
    		
   		// Inner text
       	result &= this.readNode(Element);
       	
        	
		}     // End for
		
	return result;
	}	// End function
</cfscript>


<cffunction name="extractDocx" returnType="string">
	<cfargument name="pathToDocx" required="true" type="string">

	<cfset var xmlPara = "">

	<cfzip action="read" file="#arguments.pathToDocx#" entrypath="word\document.xml" variable="this.xmlString">

	<cfset this.xmlPara = xmlparse(this.xmlString).document.body>

	<cfreturn this.ReadNode(this.xmlPara)>
</cffunction>


</cfcomponent>




