<?xml version='1.0'?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
xmlns:xs="http://www.w3.org/2001/XMLSchema" 
xmlns:fn="http://www.w3.org/2005/xpath-functions"  
xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" 
xmlns:bif="http://www.benefitfocus.com/schemas/bifedi" 
xmlns="urn:schemas-microsoft-com:office:spreadsheet"
xmlns:html="http://www.w3.org/TR/REC-html40"
exclude-result-prefixes="fn xs bif">
<!--<xsl:output method="xml" indent="yes"/>-->
<xsl:output method="xml" indent="yes" omit-xml-declaration="no" />
<xsl:variable name="dateTime" select="current-date()"/>
<xsl:variable name="root" select="/"/>

<xsl:attribute-set name="table_set">
   <xsl:attribute name="ss:ExpandedColumnCount">
   	<xsl:value-of select="number(6)"/>
	</xsl:attribute>
	   <xsl:attribute name="ss:ExpandedRowCount">
   	<xsl:value-of select="number(bif:totalSegments()+3)"/>
	</xsl:attribute>
	   <xsl:attribute name="x:FullColumns">
   	<xsl:value-of select="number(1)"/>
	</xsl:attribute>
	   <xsl:attribute name="x:FullRows">
   	<xsl:value-of select="number(1)"/>
	</xsl:attribute>
	   <xsl:attribute name="ss:DefaultColumnWidth">
   	<xsl:value-of select="number(53)"/>
	</xsl:attribute>
	<xsl:attribute name="ss:DefaultRowHeight">
   	<xsl:value-of select="number(14)"/>
	</xsl:attribute>
</xsl:attribute-set>

<xsl:template match="/" >
<xsl:variable name="lmarkets" select="distinct-values(//market[@type != '']/@type)"/>
<Workbook> <!--xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">-->
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Benefitfocus</Author>
  <LastAuthor>Benefitfocus</LastAuthor>
  <Created> <xsl:value-of select="$dateTime"/></Created>
  <LastSaved><xsl:value-of select="$dateTime"/></LastSaved>
  <Company>Benefitfocus</Company>
  <Version>14.0</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>7260</WindowHeight>
  <WindowWidth>22020</WindowWidth>
  <WindowTopX>240</WindowTopX>
  <WindowTopY>80</WindowTopY>
  <TabRatio>600</TabRatio>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="m410364160">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="m410364180">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="m410364200">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="m410364220">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="m410339368">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="m410339388">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="m410339408">
   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s73">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s74">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Interior ss:Color="#FCF305" ss:Pattern="Solid"/>
  </Style>
 </Styles>
  <Worksheet ss:Name="Sheet1">
  <Table xsl:use-attribute-sets="table_set">
   <Column ss:AutoFitWidth="0" ss:Width="105"/>
   <Column ss:AutoFitWidth="0" ss:Width="67"/>
   <Column ss:Width="33"/>
   <Column ss:Width="39"/>
   <Column ss:Width="51"/>
   <Row ss:AutoFitHeight="0">
    <Cell ss:MergeDown="1" ss:StyleID="m410339368"><Data ss:Type="String">Market Segment</Data></Cell>
    <Cell ss:MergeDown="1" ss:StyleID="m410339388"><Data ss:Type="String">Region</Data></Cell>
    <Cell ss:MergeAcross="2" ss:StyleID="m410339408"><Data ss:Type="String">Group Totals</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="3" ss:StyleID="s73"><Data ss:Type="String">Adds</Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="String">Terms</Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="String">Changes</Data></Cell>
    <Cell ss:StyleID="s74"><Data ss:Type="String">Totals</Data></Cell>
   </Row>
 	<xsl:for-each select="$lmarkets">
		<xsl:call-template name="createCells">
			<xsl:with-param name="mark" select="."/>
		</xsl:call-template>
	</xsl:for-each>
	<Row ss:AutoFitHeight="0">
    <Cell ss:StyleID="s74"><Data ss:Type="String">Totals</Data></Cell>
    <Cell ss:StyleID="s74"/>
    <Cell ss:StyleID="s74">
	<xsl:attribute name="ss:Formula">
   		<xsl:value-of select="concat(string('=SUM(R[-'),string(bif:totalSegments()),string(']C:R[-1]C)'))"/>
	</xsl:attribute>
	<Data ss:Type="Number"/></Cell>
    <Cell ss:StyleID="s74">
	<xsl:attribute name="ss:Formula">
   		<xsl:value-of select="concat(string('=SUM(R[-'),string(bif:totalSegments()),string(']C:R[-1]C)'))"/>
	</xsl:attribute>
	<Data ss:Type="Number"/></Cell>
    <Cell ss:StyleID="s74" >
	<xsl:attribute name="ss:Formula">
   		<xsl:value-of select="concat(string('=SUM(R[-'),string(bif:totalSegments()),string(']C:R[-1]C)'))"/>
	</xsl:attribute>
	<Data ss:Type="Number"/></Cell>
    <Cell ss:StyleID="s74"/>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Print>
    <ValidPrinterInfo/>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <PageLayoutZoom>0</PageLayoutZoom>
   <Selected/>
   <Panes />
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
</xsl:template>

<xsl:template name="createCells">
	<xsl:param name="mark" as="xs:string"/>
<!--	<xsl:variable name="marketreg">
		<xsl:choose>
			<xsl:when test="bif:total-sum-regions($mark) = bif:total-sum-nonregions($mark)"> 
			<xsl:sequence select="$root//market[@type = $mark]/region[@type != '']/@type"/>
			</xsl:when>
			<xsl:otherwise> 
			<xsl:sequence select="$root//market[@type = $mark]/region[@type != '' or @type = '']/@type"/> 
			</xsl:otherwise>
		</xsl:choose>
	 </xsl:variable>-->
	<xsl:variable name="cmarketreg">
		<xsl:choose>
			<xsl:when test="bif:total-sum-regions($mark) = bif:total-sum-nonregions($mark)"> 
			<xsl:value-of select="count(distinct-values($root//market[@type = $mark]/region[@type != '']/@type))"/>
			</xsl:when>
			<xsl:otherwise> 
			<xsl:value-of select="count(distinct-values($root//market[@type = $mark]/region[@type != '' or @type = '']/@type))"/> 
			</xsl:otherwise>
		</xsl:choose>
	 </xsl:variable>
	 <xsl:choose>
	 <xsl:when test="bif:total-sum-regions($mark) = bif:total-sum-nonregions($mark)">
	<xsl:for-each select="$root//market[@type = $mark]/region[@type != '']/@type">
	<xsl:choose>
		<xsl:when test="position() = 1">
		<xsl:call-template name="createcountcell1">
			<xsl:with-param name="mark" select="$mark"/>
			<xsl:with-param name="mreg" select="."/>
			<xsl:with-param name="amark" select="$cmarketreg"/>
		</xsl:call-template>
		</xsl:when>
		<xsl:otherwise>
		<xsl:call-template name="createcountcell2">
			<xsl:with-param name="mark" select="$mark"/>
			<xsl:with-param name="mreg" select="."/>
		</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
	</xsl:for-each>
</xsl:when>
<xsl:otherwise>
	<xsl:for-each select="$root//market[@type = $mark]/region/@type">
	<xsl:choose>
		<xsl:when test="position() = 1">
		<xsl:call-template name="createcountcell1">
			<xsl:with-param name="mark" select="$mark"/>
			<xsl:with-param name="mreg" select="."/>
			<xsl:with-param name="amark" select="$cmarketreg"/>
		</xsl:call-template>
		</xsl:when>
		<xsl:otherwise>
		<xsl:call-template name="createcountcell2">
			<xsl:with-param name="mark" select="$mark"/>
			<xsl:with-param name="mreg" select="."/>
			<xsl:with-param name="amark" select="$cmarketreg"/>
		</xsl:call-template>
		</xsl:otherwise>
	</xsl:choose>
	</xsl:for-each>
</xsl:otherwise>
</xsl:choose>
</xsl:template>

<xsl:template name="createcountcell1" >
	<xsl:param name="mark"/>
	<xsl:param name="mreg"/>
	<xsl:param name="amark"/>
	<xsl:variable name="madds"> 
	<xsl:choose>
	<xsl:when test="$mreg != ''">
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = $mreg]/counts/count[@type='ADD']/text()"/>
	</xsl:when>
		<xsl:when test="$mreg = '' and $amark &gt; 1">
	<xsl:value-of select="bif:calc-nonreagion($mark,'ADD')"/>
	</xsl:when>
	<xsl:otherwise>
		<xsl:value-of select="$root//market[@type = $mark]/region[@type = '']/counts/count[@type='ADD']/text()"/>
	</xsl:otherwise>
	</xsl:choose>
	</xsl:variable>
	<xsl:variable name="mchange" >
	<xsl:choose>
		<xsl:when test="$mreg != ''">
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = $mreg]/counts/count[@type='CHANGE']/text()"/>
	</xsl:when>
		<xsl:when test="$mreg = '' and $amark &gt; 1">
	<xsl:value-of select="bif:calc-nonreagion($mark,'CHANGE')"/>
	</xsl:when>
	<xsl:otherwise>
		<xsl:value-of select="$root//market[@type = $mark]/region[@type = '']/counts/count[@type='CHANGE']/text()"/>
	</xsl:otherwise>
	</xsl:choose>
	</xsl:variable>
	<xsl:variable name="mterm">
	<xsl:choose>
   	<xsl:when test="$mreg != ''">
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = $mreg]/counts/count[@type='TERM']/text()"/>
	</xsl:when>
		<xsl:when test="$mreg = '' and $amark &gt; 1">
	<xsl:value-of select="bif:calc-nonreagion($mark,'TERM')"/>
	</xsl:when>
	<xsl:otherwise>
		<xsl:value-of select="$root//market[@type = $mark]/region[@type = '']/counts/count[@type='TERM']/text()"/>
	</xsl:otherwise>
	</xsl:choose>
	</xsl:variable>
	<Row ss:AutoFitHeight="0">
    <Cell>
	<xsl:choose>
		<xsl:when test="$amark &gt; 1">
	 <xsl:attribute name="ss:MergeDown"><xsl:value-of select="number($amark - 1)"/></xsl:attribute>
	<xsl:attribute name="ss:StyleID">m410364160</xsl:attribute>
	</xsl:when>
	<xsl:otherwise>
		<xsl:attribute name="ss:StyleID">m410364160</xsl:attribute>
	</xsl:otherwise>
	</xsl:choose>
	<Data ss:Type="String">
	<xsl:choose>
	<xsl:when test="string-length($mark) &lt; 1">
	<xsl:value-of select="string('N\A')"/>
	</xsl:when>
	<xsl:otherwise>
	<xsl:value-of select="string($mark)"/>
	</xsl:otherwise>
	</xsl:choose>
	</Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="String">
	<xsl:choose>
	<xsl:when test="string-length($mreg) &lt; 1">
	<xsl:value-of select="string('N\A')"/>
	</xsl:when>
	<xsl:otherwise>
	<xsl:value-of select="string($mreg)"/>
	</xsl:otherwise>
	</xsl:choose>
	</Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="Number"><xsl:value-of select="number($madds)"/></Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="Number"><xsl:value-of select="number($mterm)"/></Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="Number"><xsl:value-of select="number($mchange)"/></Data></Cell>
    <Cell ss:StyleID="s74" ss:Formula="=SUM(RC[-3]:RC[-1])"><Data ss:Type="Number"><xsl:value-of select="number($madds) + number($mchange) + number($mterm)"/></Data></Cell>
   </Row>
</xsl:template>

<xsl:template name="createcountcell2">
	<xsl:param name="mark"/>
	<xsl:param name="mreg"/>
	<xsl:param name="amark"/>
	<xsl:variable name="madds"> 
	<xsl:choose>
	<xsl:when test="$mreg != ''">
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = $mreg]/counts/count[@type='ADD']/text()"/>
	</xsl:when>
	<xsl:when test="$mreg = '' and $amark &gt; 1">
	<xsl:value-of select="bif:calc-nonreagion($mark,'ADD')"/>
	</xsl:when>
	<xsl:otherwise>
		<xsl:value-of select="$root//market[@type = $mark]/region[@type = '']/counts/count[@type='ADD']/text()"/>
	</xsl:otherwise>
	</xsl:choose>
	</xsl:variable>
	<xsl:variable name="mchange" >
	<xsl:choose>
		<xsl:when test="$mreg != ''">
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = $mreg]/counts/count[@type='CHANGE']/text()"/>
	</xsl:when>
		<xsl:when test="$mreg = '' and $amark &gt; 1">
	<xsl:value-of select="bif:calc-nonreagion($mark,'CHANGE')"/>
	</xsl:when>
	<xsl:otherwise>
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = '']/counts/count[@type='CHANGE']/text()"/>
	</xsl:otherwise>
	</xsl:choose>
	</xsl:variable>
	<xsl:variable name="mterm">
	<xsl:choose>
   	<xsl:when test="$mreg != ''">
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = $mreg]/counts/count[@type='TERM']/text()"/>
	</xsl:when>
	<xsl:when test="$mreg = '' and $amark &gt; 1">
	<xsl:value-of select="bif:calc-nonreagion($mark,'TERM')"/>
	</xsl:when>
	<xsl:otherwise>
	<xsl:value-of select="$root//market[@type = $mark]/region[@type = '']/counts/count[@type='TERM']/text()"/>
	</xsl:otherwise>
	</xsl:choose>
	</xsl:variable>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="2" ss:StyleID="s73"><Data ss:Type="String"><xsl:value-of select="string($mreg)"/></Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="Number"><xsl:value-of select="number($madds)"/></Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="Number"><xsl:value-of select="number($mterm)"/></Data></Cell>
    <Cell ss:StyleID="s73"><Data ss:Type="Number"><xsl:value-of select="number($mchange)"/></Data></Cell>
    <Cell ss:StyleID="s74" ss:Formula="=SUM(RC[-3]:RC[-1])"><Data ss:Type="Number"><xsl:value-of select="number($madds) + number($mchange) + number($mterm)"/></Data></Cell>
   </Row>
</xsl:template>

<xsl:function name="bif:total-sum-regions" as="xs:double">
	<xsl:param name="markseg" as="xs:string"/>
	<xsl:variable name="markreg" select="$root//market[@type = $markseg]/region[@type != '']/counts/count[text() != '0']"/>
	<xsl:sequence select="sum($markreg)"/>
</xsl:function>

<xsl:function name="bif:total-sum-nonregions" as="xs:double" >
	<xsl:param name="markseg" as="xs:string"/>
	<xsl:variable name="markreg" select="$root//market[@type = $markseg]/region[@type = '']/counts/count[text() != '0']"/>
	<xsl:sequence select="sum($markreg)" />
</xsl:function>

<xsl:function name="bif:calc-nonreagion" as="xs:double">
	<xsl:param name="markseg" as="xs:string"/>
	<xsl:param name="ctype" as="xs:string"/>
	<xsl:variable name="markreg" select="$root//market[@type = $markseg]/region[@type != '']/counts/count[@type = $ctype]"/>
	<xsl:variable name="marknoreg" select="$root//market[@type = $markseg]/region[@type = '']/counts/count[@type = $ctype]"/>
	<xsl:sequence select="abs(sum($markreg) - sum($marknoreg))" />
</xsl:function>

<xsl:function name="bif:totalMarkets" as="xs:double">
<xsl:sequence select="count(distinct-values($root//market[@type != '']/@type))"/>
</xsl:function>

<xsl:function name="bif:totalSegments" as="xs:double">
<xsl:sequence select="count($root//market[@type != '']/region[@type = '' or @type != ''])"/>
</xsl:function>

</xsl:stylesheet>