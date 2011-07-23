<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="html"/>
<xsl:template match="/">
    <html>
    	<head>
    		<link REL="StyleSheet" HREF="Report.css" TYPE="text/css"/>
    		<script language="javascript" type="text/javascript">
<![CDATA[
                function toggleMenu(id,b)
    			{
    				if (document.getElementById)
    				{
    					var e = document.getElementById(id);
    					var b = document.getElementById(b);

    					if (e)
    					{
    						if (e.style.display != "block")
    						{
    							e.style.display = "block";
    							b.src='_images/check-min.jpg';
    						}
    						else
    						{
    							e.style.display = "none";
    							b.src='_images/check-plus.jpg';
    						}
    					}
    				}
    			}
    			function expandall()
    			{
    				var e = document.all.tags("div");
    				var b = document.all.tags("img");

    				for (var i = 0; i < e.length; i++)
    				{
    					if (e[i].style.display == "none")
    					{
    						e[i].style.display = "block";
    					}
    				}
    				for (var i = 0; i < b.length; i++)
    				{
    					if (b[i].id != "m1" && b[i].id != "m2")
    					{
    						if (b[i].src.substring(b[i].src.lastIndexOf("_"), b[i].src.length) == "_images/check-plus.jpg")
    						{
    							b[i].src = "_images/check-min.jpg";
    						}
    					}
    				}
    			}
    			function collapseall()
    			{
    				var e = document.all.tags("div");
    				var b = document.all.tags("img");

    				for (var i = 0; i < e.length; i++)
    				{
    					if(e[i].id != "Step" && e[i].id != "Summary" && e[i].id != "Application")
    					{
    						if (e[i].style.display == "block")
    						{
    							e[i].style.display = "none";
    						}
    					}
    				}
    				for (var i = 0; i < b.length; i++)
    				{
    					if (b[i].id != "m1" && b[i].id != "m2")
    					{
    						if (b[i].src.substring(b[i].src.lastIndexOf("_"), b[i].src.length) == "_images/check-min.jpg")
    							{
    								b[i].src = "_images/check-plus.jpg";
    							}
    					}
    				}
    			}]]>
    		</script>
    	</head>
        <body>
            <table ID="TableTitle">
      			<tr Class ="header">
    				<td>
                        <xsl:value-of select="Report/@Header"/>
                    </td>
                </tr>
                <tr>
                    <td>
                        Automation Test Report
                    </td>
                </tr>
            </table>
            <br/>
            <div id="Application" Style="position: relative; display:block; text-align:center">
                <table ID="TableApplication">
          			<tr>
        				<td>
                            Application Name:
                        </td>
                        <td>
    					    <span><xsl:value-of select="Report/@ApplicationName"/></span>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            Release:
                        </td>
                        <td>
    					    <span><xsl:value-of select="Report/@Release"/></span>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            Build:
                        </td>
                        <td>
    					    <span><xsl:value-of select="Report/@Build"/></span>
                        </td>
                    </tr>
                </table>
            </div>
            <table>
    			<tr>
                    <td>
                        <strong>Start Time :</strong>
                    </td>
                    <td>
                        <span><xsl:value-of select="Report/@StartTime"/></span>
                    </td>
                </tr>
                <tr>
                    <td>
                        <strong>End Time :</strong>
                    </td>
                    <td>
                        <span><xsl:value-of select="Report/@EndTime"/></span>
                    </td>
                </tr>
                <tr>
                    <td>
                        <strong>Execute Time :</strong>
                    </td>
                    <td>
                        <span><xsl:value-of select="Report/@ExecuteHourTime"/><xsl:text> Hr(s) </xsl:text></span>
                        <span><xsl:value-of select="Report/@ExecuteMinuteTime"/><xsl:text> Min(s)</xsl:text></span>
                    </td>
                </tr>
    		</table>
            <br/>
            <div id="Summary" Style="position:relative; display:block; text-align:center">
                <table id="TableSummary" cellSpacing="0" cellPadding="0">
                    <tr class = "Head">
                        <td>
                            Summary of TestSuites
                        </td>
                        <td>
                            Passed
                        </td>
                        <td>
                            Failed
                        </td>
                    </tr>
                    <xsl:for-each select="Report/TestSuite">
                    <tr class = "Count">
                        <td><xsl:value-of select="@Desc"/></td>
                        <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) = 0])"/></td>
                        <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) &gt; 0])" /></td>
                    </tr>
                    </xsl:for-each>
                    <tr class = "Count">
                        <td>Total</td>
                        <td><xsl:value-of select="count(Report/TestSuite/TestCase[count(Step[@Status='2']) = 0]) "/></td>
                        <td><xsl:value-of select="count(Report/TestSuite/TestCase[count(Step[@Status='2']) &gt; 0]) "/></td>
                    </tr>
                </table>
            </div>
            <br/>
    		<img src="_images/check-plus.jpg" onClick="expandall()" id="m1"/><xsl:text> </xsl:text><a href="#" onClick="expandall()"><span>Expand All</span></a>
    		<xsl:text> </xsl:text>
            <img src="_images/check-min.jpg" onClick="collapseall()" id="m2"/><xsl:text> </xsl:text><a href="#" onClick="collapseall()"><span>Collapse All</span></a>
            <br/>
    	    <xsl:apply-templates select="Report/TestSuite">
    		</xsl:apply-templates>
            <br/>
            <xsl:variable name="lower">abcdefghijklmnopqrstuvwxyz</xsl:variable>
            <xsl:variable name="upper">ABCDEFGHIJKLMNOPQRSTUVWXYZ</xsl:variable>
            <xsl:for-each select="Report/TestSuite">
                <div id="SummaryOfTestCase" Style="position:relative; display:none; text-align:center">
                    <table id="SummaryOfTestCaseTable" cellSpacing="0" cellPadding="0">
        				<tr class = "SummaryOfTestCaseTable_Head">
                            <td>
                                <xsl:value-of select="@Desc"/><br/>Test Cases
                            </td>
                            <td>
                                Critical
                            </td>
                            <td>
                                 High
                            </td>
                            <td>
                                Medium
                            </td>
                            <td>
                                 Low
                            </td>
                        </tr>
                        <tr class = "TotalPassed">
        					<td>Total Passed</td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) = 0 and contains(translate(@ID,$lower,$upper), '_C_') ])"/></td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) = 0 and contains(translate(@ID,$lower,$upper), '_H_') ])"/></td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) = 0 and contains(translate(@ID,$lower,$upper), '_M_') ])"/></td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) = 0 and contains(translate(@ID,$lower,$upper), '_L_') ])"/></td>
                        </tr>

                        <tr class = "TotalFailed">
                            <td>Total Failed</td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) &gt; 0 and contains(translate(@ID,$lower,$upper), '_C_') ])"/></td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) &gt; 0 and contains(translate(@ID,$lower,$upper), '_H_') ])"/></td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) &gt; 0 and contains(translate(@ID,$lower,$upper), '_M_') ])"/></td>
                            <td><xsl:value-of select="count(TestCase[count(Step[@Status='2']) &gt; 0 and contains(translate(@ID,$lower,$upper), '_L_') ])"/></td>
                        </tr>
        			</table>
        		</div>
            <br/>
    	    </xsl:for-each>
        </body>
    </html>
    </xsl:template>
	
	<xsl:template match="Report/TestSuite">
		<div Style="position:relative; display:none">
			<table id="TestSuite" cellSpacing="1" cellPadding="1"  onClick="toggleMenu('div{position()}.0', 'm{position()}.0')">
				<tr>
					<td class="message">
						<img id="m{position()}.0" src="_images/check-plus.jpg" border="0"/>
						<xsl:text> </xsl:text><xsl:value-of select="@Desc"/>
                    </td>
                    <td class="Passed">Passed -
                        <xsl:value-of select ="count(TestCase[count(Step[@Status='2']) = 0])"/>
                    </td>
					<td class="Failed">Failed -
						<xsl:value-of select ="count(TestCase[count(Step[@Status='2']) > 0])"/>
                    </td>
                </tr>
			</table>
		</div>
		<div id="div{position()}.0" Style="display:none">
			<xsl:apply-templates select="TestCase">
			<xsl:with-param name="TestCasePosition" select="position()"/>
			</xsl:apply-templates>
		</div>
	</xsl:template>
	
	<xsl:template match="TestCase">
			<xsl:param name="TestCasePosition"/>
			<div id="TestCase" Style="left:10px; position: relative; display:block">
    			<table cellSpacing="1" cellPadding="1" align="center" Class="TestCase" onClick="toggleMenu('div{concat($TestCasePosition,'.',position())}', 'm{concat($TestCasePosition,'.',position())}')">
    				<tr>
    					<td class="message">
    						<img id="m{concat($TestCasePosition,'.',position())}" src="_images\check-plus.jpg" border="0"/>
                            <xsl:text> </xsl:text><xsl:value-of select="@ID"/>
                            <xsl:text> </xsl:text><xsl:value-of select="@Desc"/>
                        </td>
    					 <xsl:element name="td">
                              <xsl:attribute name="class">
                                <xsl:if test="count(Step[@Status='2']) = 0">
                                Passed
                                </xsl:if>
                                <xsl:if test="count(Step[@Status='2']) &gt; 0">
                                Failed
                                </xsl:if>
                			</xsl:attribute>
            			    <xsl:if test="count(Step[@Status='2']) = 0">
                            	<xsl:text>Passed </xsl:text>
                            </xsl:if>
                            <xsl:if test="count(Step[@Status='2']) &gt; 0">
                            	<xsl:text>Failed </xsl:text>
                            </xsl:if>
    			          </xsl:element>
                    </tr>
    			</table>
			</div>
			<div id="div{concat($TestCasePosition,'.',position())}" Style="display:none">
			<xsl:apply-templates select="Step">
			</xsl:apply-templates>
			</div>
	</xsl:template>

	<xsl:template match="Step">
    	<div id="Step" Style="left:20px; position:relative; display:block">
    		<table cellSpacing="1" cellPadding="1" align="center" Class="Step">
    			<xsl:element name="tr">
                    <xsl:if test="position()mod 2">
                        <xsl:attribute name="class">
                        M1
                        </xsl:attribute>
                        </xsl:if>
                        <xsl:if test="not(position()mod 2)">
                        <xsl:attribute name="class">
                        M2
                        </xsl:attribute>
                    </xsl:if>

                    <xsl:element name="td">
                        <xsl:attribute name="class">
                        <xsl:if test="@Status='1'">
                        Passed
                        </xsl:if>
                        <xsl:if test="@Status='2'">
                        Failed
                        </xsl:if>
                        </xsl:attribute>
                        <xsl:text>- </xsl:text>
                        <xsl:value-of select="."/>
                    </xsl:element>
                    
                    <xsl:if test="@Detail">
                        <xsl:element name="td">
                            <xsl:attribute name="class">
                            <xsl:if test="@Status='1'">
                            DetailPassInfo
                            </xsl:if>
                            <xsl:if test="@Status='2'">
                            DetailFailInfo
                            </xsl:if>
                            </xsl:attribute>
                            <xsl:text> </xsl:text><xsl:value-of select="@Detail"/>
                        </xsl:element>
                    </xsl:if>
                    
                    <xsl:if test="@ExpectedResult">
                        <xsl:element name="td">
                            <xsl:attribute name="class">
                            <xsl:if test="@Status='1'">
                            Passed
                            </xsl:if>
                            <xsl:if test="@Status='2'">
                            Failed
                            </xsl:if>
                            </xsl:attribute>
                            <xsl:text> </xsl:text><xsl:value-of select="@ExpectedResult"/>
                        </xsl:element>
                    </xsl:if>
        		    
                    <xsl:if test="@ActualResult">
                        <xsl:element name="td">
                            <xsl:attribute name="class">
                            <xsl:if test="@Status='1'">
                            Passed
                            </xsl:if>
                            <xsl:if test="@Status='2'">
                            Failed
                            </xsl:if>
                            </xsl:attribute>
                            <xsl:text> </xsl:text><xsl:value-of select="@ActualResult"/>
                        </xsl:element>
                    </xsl:if>
        		    
                    <xsl:if test="@ScreenShotPath">
                        <xsl:element name="td">
                            <xsl:attribute name="class">
                            LinkToFile
                            </xsl:attribute>
                            <a href="{@ScreenShotPath}" target="_new">Screenshot</a>
                        </xsl:element>
                    </xsl:if>

                    <xsl:if test="@Filepath">
                        <xsl:element name="td">
                            <xsl:attribute name="class">
                            LinkToFile
                            </xsl:attribute>
                            <xsl:call-template name="GetFileName">
                                <xsl:with-param name="filepath" select="@Filepath"/>
                            </xsl:call-template>
                        </xsl:element>
                    </xsl:if>
    			</xsl:element>
    		</table>
    	</div>
	</xsl:template>

    <xsl:template name="GetFileName">
    <xsl:param name="filepath"/>
        <xsl:variable name="filename" select="substring-after($filepath,'\')"/>
        <xsl:choose>
        <xsl:when test="contains($filename,'\')">
          <xsl:call-template name="GetFileName">
              <xsl:with-param name="filepath" select="$filename"/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
            <xsl:text> </xsl:text>
            <a href="{@Filepath}" target="_new"><xsl:value-of select="$filename"/></a>
        </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

</xsl:stylesheet>





















