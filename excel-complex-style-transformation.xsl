<?xml version="1.0"?>

<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" 
                xmlns:x="urn:schemas-microsoft-com:office:excel"
                xmlns:html="http://www.w3.org/TR/REC-html40"
                exclude-result-prefixes="ss x html"
                version="1.0">
<xsl:output method="html"/>
<xsl:key name="style" match="ss:Style" use="@ss:ID" />
<xsl:key name="cell" match="ss:Cell" use="@ss:StyleID" />
<xsl:template match="/">
<xsl:element name="style">
    <xsl:for-each select="ss:Workbook/ss:Styles/ss:Style">
        <xsl:choose>
            <xsl:when test="@ss:ID='Default'">
    .default</xsl:when>
            <xsl:otherwise>
    .<xsl:value-of select="@ss:ID"/>
                <xsl:if  test="@ss:Name">, .<xsl:value-of select="translate(@ss:Name,translate(@ss:Name,'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',''),'')" /></xsl:if>
            </xsl:otherwise>
        </xsl:choose>
	{<!-- if there is a font-element, but a specific sub-element (other than font-related attributes) are missing, there is NO inheritance from parent style -->
		<xsl:if test="@ss:Parent">
			<xsl:if test="not(ss:Font/@ss:Italic) and key('style', @ss:Parent)/ss:Font/@ss:Italic=1">
		font-style:normal;</xsl:if>
			<xsl:if test="not(ss:Font/@ss:Bold) and key('style', @ss:Parent)/ss:Font/@ss:Bold=1">
		font-weight:normal;</xsl:if>
			<xsl:if test="not(ss:Font/@ss:Color) and key('style', @ss:Parent)/ss:Font/@ss:Color">
		color:<xsl:value-of select="key('style', 'Default')/ss:Font/@ss:Color"/>;</xsl:if>
			<xsl:if test="(not(ss:Font/@ss:StrikeThrough) and key('style', @ss:Parent)/ss:Font/@ss:StrikeThrough=1) and ((not(ss:Font/@ss:Underline) and key('style', @ss:Parent)/ss:Font/@ss:Underline))">
		text-decoration:initial;</xsl:if>
			<xsl:if test="not(ss:Font/@ss:VerticalAlign) and key('style', @ss:Parent)/ss:Font/@ss:VerticalAlign">
		vertical-align:baseline;</xsl:if>
		</xsl:if>
				<xsl:if test="ss:Font/@ss:Italic=1">
        font-style:italic;</xsl:if>
                <xsl:if test="ss:Font/@ss:Bold=1">
        font-weight: bold;</xsl:if>                
				<xsl:if test="ss:Font/@ss:Color"><!-- strictly color:unset;, i.e. NO inheritance from parent style or default style; but the desired result can be better achieved by inheritance from the default-style. Note that this is only relevant here. Color is the only attribute with an unpredictable value that can be unset via the front-end -->
        color:<xsl:value-of select="ss:Font/@ss:Color"/>;</xsl:if>
                <xsl:if test="ss:Font/@ss:StrikeThrough=1 or ss:Font/@ss:Underline">
					<xsl:choose>
						<xsl:when test="ss:Font/@ss:StrikeThrough=1 and ss:Font/@ss:Underline='Single'">
		text-decoration:line-through underline;
						</xsl:when>
						<xsl:when test="ss:Font/@ss:StrikeThrough=1 and ss:Font/@ss:Underline='Double'">
		text-decoration:line-through underline double;<!-- closest match in css, strike-through effect is double line as well -->
						</xsl:when>
						<xsl:when test="ss:Font/@ss:StrikeThrough=1">
        text-decoration:line-through;</xsl:when>
                <xsl:when test="ss:Font/@ss:Underline='Single'">
        text-decoration:underline;</xsl:when>
                <xsl:when test="ss:Font/@ss:Underline='Double'">
        text-decoration:underline;
        text-decoration-style: double;</xsl:when>
					</xsl:choose>
				</xsl:if>
                <xsl:if test="(ss:Font/@ss:FontName) and ((@ss:ID='Default') or (not(ss:Font/@ss:FontName = key('style', 'Default')/ss:Font/@ss:FontName)))">
        font-family:<xsl:value-of select="ss:Font/@ss:FontName"/>
                    <xsl:choose>
                        <xsl:when test="ss:Font/@x:Family='Swiss'">, sans-serif</xsl:when>
                        <xsl:when test="ss:Font/@x:Family='Roman'">, serif</xsl:when>
                        <xsl:when test="ss:Font/@x:Family='Modern'">, monospace</xsl:when>
                        <xsl:when test="ss:Font/@x:Family='Script'">, cursive</xsl:when>
                        <xsl:when test="ss:Font/@x:Family='Decorative'">, fantasy</xsl:when>
                    </xsl:choose>;</xsl:if>
                <xsl:choose>
                    <xsl:when test="(ss:Font/@ss:Size) and ((@ss:ID='Default') or (not(ss:Font/@ss:Size = key('style', 'Default')/ss:Font/@ss:Size)))">
                        <xsl:choose>
                            <xsl:when test="ss:Font/@ss:VerticalAlign">
                                <xsl:if test="ss:Font/@ss:VerticalAlign='Superscript'">
        vertical-align:super;
        font-size:<xsl:value-of select="round(ss:Font/@ss:Size*8.3) div 10"/>pt;</xsl:if>
                                <xsl:if test="ss:Font/@ss:VerticalAlign='Subscript'">
        vertical-align:sub;
        font-size:<xsl:value-of select="round(ss:Font/@ss:Size*8.3) div 10"/>pt;</xsl:if>
                            </xsl:when>
                            <xsl:otherwise>
        font-size:<xsl:value-of select="ss:Font/@ss:Size"/>pt;</xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:if test="ss:Font/@ss:VerticalAlign='Superscript'">vertical-align:super;
        font-size:0.83em;</xsl:if>
                        <xsl:if test="ss:Font/@ss:VerticalAlign='Subscript'">vertical-align:sub;
        font-size:0.83em;</xsl:if>
                    </xsl:otherwise>
        </xsl:choose>
        <!-- cell formats: NO synchronisation with parent element in XSLT. CSS inheritance applies: some 'unset' attributes might be inherited from style template despite being switched off in cell--><xsl:if test="ss:Alignment/@ss:Horizontal='Left'">
        text-align:left;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Horizontal='Center'">
        text-align:center;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Horizontal='Right'">
        text-align:right;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Horizontal='Left'">
        text-align:left;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Horizontal='Fill'">
        text-align:justify;</xsl:if><!-- change to justify-all once browser's support it -->
        <xsl:if test="ss:Alignment/@ss:Horizontal='Justify'">
        text-align:justify;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Horizontal='CenterAcrossSelection'">
        text-align:center;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Horizontal='Distributed'">
        text-align:justify;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Vertical='Top'">
        vertical-align: top;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Vertical='Center'">
        vertical-align: middle;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Vertical='Bottom'">
        vertical-align: bottom;</xsl:if>
        <xsl:if test="ss:Interior/@ss:Color">
        background-color:<xsl:value-of select="ss:Interior/@ss:Color"/>;</xsl:if><!-- no support for patterns other than solid -->
        <xsl:if test="ss:Alignment/@ss:VerticalText=1">
        writing-mode: vertical-lr;
        text-orientation: upright;</xsl:if>
        <xsl:if test="ss:Alignment/@ss:Rotate">
        transform: rotate(<xsl:value-of select="-1*ss:Alignment/@ss:Rotate"/>deg);</xsl:if><!-- rotation will not work correctly out of the box -->
        <xsl:if test="ss:Alignment/@ss:ReadingOrder='RightToLeft'">
        direction: rtl;</xsl:if>
        <xsl:for-each select="ss:Borders/ss:Border">
        border-<xsl:choose>
                <xsl:when test="@ss:Position='Bottom'">bottom</xsl:when>
                <xsl:when test="@ss:Position='Left'">left</xsl:when>
                <xsl:when test="@ss:Position='Right'">right</xsl:when>
                <xsl:when test="@ss:Position='Top'">top</xsl:when>
            </xsl:choose>:<xsl:value-of select="@ss:Weight"/><xsl:if test="not(@ss:Weight)">1</xsl:if>px <xsl:choose>
                <xsl:when test="@ss:LineStyle='Double'">double</xsl:when>
                <xsl:when test="@ss:LineStyle='Dot'">dotted</xsl:when>
                <xsl:when test="@ss:LineStyle='DashDotDot'">dotted</xsl:when>
                <xsl:when test="@ss:LineStyle='DashDot'">dotted</xsl:when>
                <xsl:when test="@ss:LineStyle='SlantDashDot'">dashed</xsl:when>
                <xsl:when test="@ss:LineStyle='Dash'">dashed</xsl:when>
                <xsl:otherwise>solid</xsl:otherwise>
            </xsl:choose><xsl:if test="@ss:Color"><xsl:text> </xsl:text><xsl:value-of select="@ss:Color"/></xsl:if>;</xsl:for-each>
    }
</xsl:for-each>
</xsl:element>
        <xsl:for-each select="ss:Workbook/ss:Worksheet">
        <div id="{translate(./@ss:Name,translate(./@ss:Name,'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',''),'')}">
        <xsl:for-each select="ss:Table"><table class="default">
        <xsl:for-each select="ss:Row"><tr>
        <xsl:for-each select="ss:Cell">
            <xsl:element name="td">
                        <xsl:if test="@ss:StyleID">
                            <xsl:attribute name="class">
								<xsl:value-of select="@ss:StyleID"/>
								<xsl:if test="string(key('style', @ss:StyleID)/@ss:Name)">
										<xsl:text> </xsl:text>
										<xsl:value-of select="translate(key('style', @ss:StyleID)/@ss:Name,translate(key('style', @ss:StyleID)/@ss:Name,'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',''),'')"/>
								</xsl:if>
								<xsl:if test="key('style', @ss:StyleID)/@ss:Parent">
										<xsl:text> </xsl:text>
										<xsl:value-of select="key('style', @ss:StyleID)/@ss:Parent"/>
								</xsl:if>
								<xsl:if test="string(key('style', key('style', @ss:StyleID)/@ss:Parent)/@ss:Name)">
										<xsl:text> </xsl:text>
										<xsl:value-of select="translate(key('style', key('style', @ss:StyleID)/@ss:Parent)/@ss:Name,translate(key('style', key('style', @ss:StyleID)/@ss:Parent)/@ss:Name,'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',''),'')"/>
								</xsl:if>
							</xsl:attribute>
                            <xsl:if test="((key('style', @ss:StyleID)/ss:Font/@ss:Bold=1) and (.//html:B)) or ((key('style', @ss:StyleID)/ss:Font/@ss:Italic=1) and (.//html:I)) or (.//html:Font/@html:Color) or ((key('style', @ss:StyleID)/ss:Font/@ss:StrikeThrough=1) and (.//html:S)) or ((key('style', @ss:StyleID)/ss:Font/@ss:Underline) and (.//html:U)) or (((key('style', @ss:StyleID)/ss:Font/@ss:VerticalAlign='Superscript') and (.//html:Sup)) or ((key('style', @ss:StyleID)/ss:Font/@ss:VerticalAlign='Superscript') and (.//html:Sup)))">
                            <xsl:attribute name="style">
                                <xsl:if test="(key('style', @ss:StyleID)/ss:Font/@ss:Bold=1) and (.//html:B)">font-weight:normal;</xsl:if>
                                <xsl:if test="(key('style', @ss:StyleID)/ss:Font/@ss:Italic=1) and (.//html:I)">font-style:normal;</xsl:if>
                                <xsl:if test=".//html:Font/@html:Color">color:<xsl:value-of select="key('style', 'Default')/ss:Font/@ss:Color"/>;</xsl:if>
                                <xsl:if test="(key('style', @ss:StyleID)/ss:Font/@ss:StrikeThrough=1) and (.//html:S)">text-decoration:initial;</xsl:if>
                                <xsl:if test="(key('style', @ss:StyleID)/ss:Font/@ss:Underline) and (.//html:U)">text-decoration:initial;</xsl:if>
                                <xsl:if test="((key('style', @ss:StyleID)/ss:Font/@ss:VerticalAlign='Superscript') and (.//html:Sup)) or ((key('style', @ss:StyleID)/ss:Font/@ss:VerticalAlign='Superscript') and (.//html:Sup))">vertical-align:baseline;<xsl:if test="key('style', @ss:StyleID)/ss:Font/@ss:Size">font-size:<xsl:value-of select="key('style', @ss:StyleID)/ss:Font/@ss:Size"/>pt;</xsl:if></xsl:if>
                            </xsl:attribute>
                            </xsl:if>
                        </xsl:if>
    	                <xsl:apply-templates/>
             </xsl:element>
        </xsl:for-each>
        </tr></xsl:for-each>
        </table></xsl:for-each>
        </div>
        </xsl:for-each>
</xsl:template>

<xsl:template match="node()">
    <xsl:choose>
        <xsl:when test="name(.)='Font'">
            <xsl:choose>
                <xsl:when test="not(@*)">
                    <xsl:apply-templates/>
                </xsl:when>
                <xsl:when test="(not(@html:Face) and not(@html:Size) and (@html:Color) and (@html:Color = key('style', 'Default')/ss:Font/@ss:Color))">
                    <xsl:apply-templates/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:element name="span">
                        <xsl:attribute name="style">
                            <xsl:if test="((@html:Color) and (not(@html:Color = key('style', 'Default')/ss:Font/@ss:Color)))">color:<xsl:value-of select="@html:Color"/>;</xsl:if>
                            <xsl:if test="@html:Size">font-size:<xsl:value-of select="@html:Size"/>pt;</xsl:if>
                            <xsl:if test="@html:Face">font-family:<xsl:value-of select="@html:Face"/>
                            <xsl:choose>
                                <xsl:when test="@x:Family='Swiss'">,sans-serif</xsl:when>
                                <xsl:when test="@x:Family='Roman'">,serif</xsl:when>
                                <xsl:when test="@x:Family='Modern'">,monospace</xsl:when>
                                <xsl:when test="@x:Family='Script'">,cursive</xsl:when>
                                <xsl:when test="@x:Family='Decorative'">,fantasy</xsl:when>
                            </xsl:choose>;</xsl:if>
                        </xsl:attribute>
                        <xsl:apply-templates/>
                    </xsl:element>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:when>
        <xsl:when test="name(.)='U'">
            <xsl:choose>
                <xsl:when test="@html:Style='text-underline:double'">
                    <u style="text-decoration-style: double"><xsl:apply-templates /></u>
                </xsl:when>
                <xsl:otherwise>
                    <u><xsl:apply-templates /></u>
                </xsl:otherwise>
            </xsl:choose> 
        </xsl:when>
        <xsl:when test="name(.)='B'">
            <b><xsl:apply-templates /></b>
        </xsl:when>
        <xsl:when test="name(.)='I'">
            <i><xsl:apply-templates /></i>
        </xsl:when>
        <xsl:when test="name(.)='Sup'">
            <sup><xsl:apply-templates /></sup>
        </xsl:when>
        <xsl:when test="name(.)='Sub'">
            <sub><xsl:apply-templates /></sub>
        </xsl:when>
        <xsl:when test="name(.)='S'">
            <del><xsl:apply-templates /></del>
        </xsl:when>
        <xsl:otherwise>
            <xsl:apply-templates />
        </xsl:otherwise>
    </xsl:choose> 
</xsl:template>

<xsl:template name="break">
  <xsl:param name="text" select="string(.)"/>
  <xsl:choose>
    <xsl:when test="contains($text, '&#10;')">
      <xsl:value-of select="substring-before($text, '&#10;')"/>
      <br/>
      <xsl:call-template name="break">
        <xsl:with-param 
          name="text" 
          select="substring-after($text, '&#10;')"
        />
      </xsl:call-template>
    </xsl:when>
    <xsl:otherwise>
      <xsl:value-of select="$text"/>
    </xsl:otherwise>
  </xsl:choose>
</xsl:template>

<!--<xsl:template match="text()"><xsl:value-of select="." /></xsl:template>-->
<xsl:template match="text()">
  <xsl:call-template name="break" />
</xsl:template>

</xsl:stylesheet>