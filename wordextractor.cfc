/*
The MIT License (MIT)

Copyright (c) 2011, 2018 James Mohler

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/ 


component output="false"	{


this.xmlPara = ""; // parsed into XML nodes
this.xmlString = ""; // raw text
this.proofErr = "spellEnd";
variables.listcounter = 1; // used to set order list numbers
variables.CRLF = Chr(13) & Chr(10);


private string function ReadNode (required xml Node) output="false" {

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



	for (var Element in arguments.node.xmlChildren)	{

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
					result &= '<h1>#ReadNode(Element)#</h1>#variables.crlf#';
					break;

				case "Heading2" :
					variables.listcounter = 1;
					result &= '<h2>#ReadNode(Element)#</h2>#variables.crlf#';
					break;

				case "Heading3" :
					variables.listcounter = 1;
					result &= '<h3>#ReadNode(Element)#</h3>#variables.crlf#';
					break;

				case "Heading4" :
					variables.listcounter = 1;
					result &= '<h4>#ReadNode(Element)#</h4>#variables.crlf#';
					break;

					case "ListParagraph" :

						if (wnumid == 2)	{
							result &= '<ol start="#listcounter#"><li>#ReadNode(Element)#</li></ol>#variables.crlf#';
							}
						else	{
						result &= '<li>#ReadNode(Element)#</li>#variables.crlf#';
						}
						
					variables.listcounter++;

					break;



				default :
					variables.listcounter = 1;
					result &= '<p>#ReadNode(Element)#</p>#variables.crlf#';
				}


			break; // end of w:p


		case "w:r"	: // This handles bolds and italics
				wrPr = "";
				if (Element.XMLChildren[1].XMLName == "w:rPr")	{
					wrPr = Element.XMLChildren[1].XMLChildren[1].XMLName;
					}
	
				switch (wrPr)	{
					case "w:b" :
						result &= "<b>#ReadNode(Element)#</b>";
						break;

					case "w:i" :
						result &= "<i>#ReadNode(Element)#</i>";
						break;
					default :
						// Unknown
						result &= ReadNode(Element);
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
		result &= readNode(Element);


		}     // End for

	return result;
	}	// End function



string function extractDocx(required string pathToDocX) output="false"	{

	cfzip(action="read", file=arguments.pathToDocx, entrypath="word\document.xml", variable="this.xmlString");
	
	this.xmlPara = xmlparse(this.xmlString).document.body;
	
	return ReadNode(this.xmlPara);
	}


} // end component	




