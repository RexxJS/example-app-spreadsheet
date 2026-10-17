# Writes tests/headless/lib/xlsx_fixtures.ts: XLSX files made by hand (their
# XML written here, zipped by Python's zipfile, deflated), as base64.
import zipfile, io, base64
CT='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'''
RELS='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
WB='''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="0"/><sheets><sheet name="Sales" sheetId="1" r:id="rId1"/><sheet name="Notes &amp; more" sheetId="2" r:id="rId2" state="hidden"/></sheets><definedNames><definedName name="QTY">Sales!$B$2:$B$3</definedName><definedName name="RATE" localSheetId="0">Sales!$D$1</definedName><definedName name="DOUBLE">LAMBDA(x,x*2)</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Sales!$A$1:$D$4</definedName></definedNames></workbook>'''
WBRELS='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="worksheet" Target="/xl/worksheets/sheet2.xml"/><Relationship Id="rId3" Type="sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId4" Type="styles" Target="styles.xml"/></Relationships>'''
SST='''<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5"><si><t>Item</t></si><si><t>Qty</t></si><si><r><t>pe</t></r><r><rPr><b/></rPr><t>ns</t></r></si><si><t>ink &amp; nib</t></si><si><t>007</t></si></sst>'''
STY='''<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="14"/><xf numFmtId="10"/></cellXfs></styleSheet>'''
S1='''<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="20" customWidth="1"/><col min="5" max="5" hidden="1"/></cols><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>Total</t></is></c><c r="D1" s="3"><v>0.2</v></c></row>
<row r="2" ht="30" customHeight="1"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>3</v></c><c r="C2" s="1"><f t="shared" ref="C2:C3" si="0">B2*(1+$D$1)</f><v>3.6</v></c></row>
<row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3"><v>1.5</v></c><c r="C3" s="1"><f t="shared" si="0"/><v>1.8</v></c></row>
<row r="4" hidden="1"><c r="A4" t="b"><v>1</v></c><c r="B4"><f>SUM(QTY)</f><v>4.5</v></c><c r="C4" s="2"><v>45306</v></c><c r="D4" t="s"><v>4</v></c></row>
<row r="200"><c r="A200"><v>9</v></c></row>
<row r="5"><c r="AB5"><v>1</v></c></row>
</sheetData><mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells></worksheet>'''
S2='''<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>Sales!B2*2</f><v>6</v></c><c r="B1" t="str"><f>"x"&amp;"y"</f><v>xy</v></c></row></sheetData></worksheet>'''
def make(files):
    b=io.BytesIO()
    with zipfile.ZipFile(b,'w',zipfile.ZIP_DEFLATED) as z:
        for n,t in files: z.writestr(n,t)
    return base64.b64encode(b.getvalue()).decode()
basic=make([("[Content_Types].xml",CT),("_rels/.rels",RELS),("xl/workbook.xml",WB),("xl/_rels/workbook.xml.rels",WBRELS),("xl/sharedStrings.xml",SST),("xl/styles.xml",STY),("xl/worksheets/sheet1.xml",S1),("xl/worksheets/sheet2.xml",S2)])
# a stored (not deflated) one, 1904 dates, no shared strings or styles
b=io.BytesIO()
with zipfile.ZipFile(b,'w',zipfile.ZIP_STORED) as z:
    z.writestr("[Content_Types].xml",CT); z.writestr("_rels/.rels",RELS)
    z.writestr("xl/workbook.xml",'<workbook xmlns="x" xmlns:r="r"><workbookPr date1904="1"/><sheets><sheet name="Only" sheetId="1" r:id="rId1"/></sheets></workbook>')
    z.writestr("xl/_rels/workbook.xml.rels",'<Relationships><Relationship Id="rId1" Target="worksheets/a.xml"/></Relationships>')
    z.writestr("xl/worksheets/a.xml",'<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>')
stored=base64.b64encode(b.getvalue()).decode()
STY2='''<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><i/><u/><sz val="14"/><color rgb="FFFF0000"/><name val="Calibri"/></font><font><strike/><sz val="11"/><color theme="4"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF000000"/></left><right/><top style="medium"><color theme="5"/></top><bottom style="mediumDashed"><color rgb="FF00FF00"/></bottom><diagonal/></border></borders>
<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0"><alignment horizontal="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="1"/><xf numFmtId="4" fontId="0" fillId="0" borderId="0"><alignment horizontal="right"/></xf></cellXfs></styleSheet>'''
S3='''<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Big</t></is></c><c r="B1" s="2"><v>2</v></c><c r="C1" s="3"><v>1234.5</v></c></row><row r="2"><c r="A2"><v>5</v></c></row></sheetData>
<dataValidations count="4"><dataValidation type="whole" operator="between" errorStyle="warning" allowBlank="1" showInputMessage="1" prompt="one to ten" sqref="A2:A3"><formula1>1</formula1><formula2>10</formula2></dataValidation><dataValidation type="list" sqref="B2"><formula1>"Yes,No"</formula1></dataValidation><dataValidation type="custom" sqref="C2:C3"><formula1>C2&gt;A2</formula1></dataValidation><dataValidation type="list" sqref="D2"><formula1>$F$1:$F$3</formula1></dataValidation></dataValidations><legacyDrawing r:id="rId2"/></worksheet>'''
CM='''<?xml version="1.0" encoding="UTF-8"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>me</author></authors><commentList><comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t>me:</t></r><r><t xml:space="preserve"> the big one</t></r></text></comment></commentList></comments>'''
S3RELS='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="comments" Target="../comments1.xml"/><Relationship Id="rId2" Type="vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>'''
WB2='''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Looks" sheetId="1" r:id="rId1"/></sheets></workbook>'''
WB2RELS='''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Target="styles.xml"/></Relationships>'''
styled=make([("[Content_Types].xml",CT),("_rels/.rels",RELS),("xl/workbook.xml",WB2),("xl/_rels/workbook.xml.rels",WB2RELS),("xl/styles.xml",STY2),("xl/worksheets/sheet1.xml",S3),("xl/worksheets/_rels/sheet1.xml.rels",S3RELS),("xl/comments1.xml",CM),("xl/drawings/vmlDrawing1.vml","<xml/>")])
STY3='''<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="2"><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><font><color theme="4"/></font></dxf></dxfs></styleSheet>'''
S4='''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>pens</t></is></c><c r="B2"><v>3</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>ink</t></is></c><c r="B3"><v>12</v></c></row>
<row r="4"><c r="A4" t="inlineStr"><is><t>Total</t></is></c><c r="B4"><f>SUBTOTAL(109,Stock[Qty])</f><v>15</v></c></row></sheetData>
<conditionalFormatting sqref="B2:B3"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>10</formula></cfRule><cfRule type="expression" dxfId="1" priority="1" stopIfTrue="1"><formula>$A2="ink"</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="A2:A3"><cfRule type="containsText" dxfId="1" priority="3" operator="containsText" text="en"><formula>NOT(ISERROR(SEARCH("en",A2)))</formula></cfRule><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule><cfRule type="dataBar" priority="5"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule><cfRule type="iconSet" priority="6"><iconSet iconSet="3Arrows" reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule><cfRule type="timePeriod" dxfId="0" priority="7" timePeriod="last7Days"><formula>x</formula></cfRule><cfRule type="aboveAverage" dxfId="0" priority="8" aboveAverage="0"/><cfRule type="dataBarX" priority="9"/></conditionalFormatting>
<tableParts count="1"><tablePart r:id="rId1"/></tableParts>
<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext></extLst></worksheet>'''
TB='''<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Stock" ref="A1:B4" totalsRowCount="1"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="Item" totalsRowLabel="Total"/><tableColumn id="2" name="Qty" totalsRowFunction="sum"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'''
S4RELS='''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="../tables/table1.xml"/></Relationships>'''
WB3='''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Rules" sheetId="1" r:id="rId1"/></sheets></workbook>'''
WB3RELS='''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Target="styles.xml"/><Relationship Id="rId5" Target="theme/theme1.xml"/></Relationships>'''
parts=make([("[Content_Types].xml",CT),("_rels/.rels",RELS),("xl/workbook.xml",WB3),("xl/_rels/workbook.xml.rels",WB3RELS),("xl/styles.xml",STY3),("xl/worksheets/sheet1.xml",S4),("xl/worksheets/_rels/sheet1.xml.rels",S4RELS),("xl/tables/table1.xml",TB),
  ("xl/theme/theme1.xml",'<a:theme xmlns:a="a" name="Mine"/>'),("docProps/core.xml",'<cp:coreProperties xmlns:cp="c"><dc:creator xmlns:dc="d">Paul</dc:creator></cp:coreProperties>'),("xl/charts/chart1.xml","<c/>"),("xl/media/image1.png","PNG"),("xl/pivotTables/pivotTable1.xml","<p/>")])
CONTENT='''<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet>
<table:table table:name="Stock"><table:table-column table:number-columns-repeated="3"/>
<table:table-row><table:table-cell office:value-type="string"><text:p>Item</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>Qty</text:p></table:table-cell><table:table-cell table:number-columns-spanned="2" office:value-type="string"><text:p>Wide</text:p></table:table-cell><table:covered-table-cell/></table:table-row>
<table:table-row><table:table-cell office:value-type="string"><text:p>pens<text:s text:c="2"/>red</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell><table:table-cell table:formula="of:=[.B2]*2" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell><table:table-cell office:value-type="percentage" office:value="0.25"><text:p>25%</text:p></table:table-cell></table:table-row>
<table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell><table:table-cell table:number-columns-repeated="2" office:value-type="string"><text:p>007</text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell office:value-type="date" office:date-value="2024-01-15"><text:p>01/15/24</text:p></table:table-cell><table:table-cell office:value-type="time" office:time-value="PT36H00M00S"><text:p>36:00:00</text:p></table:table-cell></table:table-row>
<table:table-row table:number-rows-repeated="1048570"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>
</table:table><table:table table:name="Two"><table:table-row><table:table-cell office:value-type="float" office:value="1.5"><text:p>1.5</text:p></table:table-cell></table:table-row></table:table>
</office:spreadsheet></office:body></office:document-content>'''
ods=make([("mimetype","application/vnd.oasis.opendocument.spreadsheet"),("content.xml",CONTENT)])
big=make([("a.txt","hello "*5000+"".join(chr(32+(i*7)%90) for i in range(3000)))])
open('/Users/paul/scm/spraedsheet/tests/headless/lib/xlsx_fixtures.ts','w').write(f'''// XLSX files made by hand for the tests: their XML written in
// mkxlsx.py beside this file (by hand, not by GridCraft or Excel), zipped by
// Python's zipfile, as base64. BASIC: two sheets (the second hidden),
// shared and rich strings, a shared formula, number formats built in and
// the file's own, a merge, a frozen row, widths and heights, hidden rows
// and columns, defined names, cells past the sheet's edge. STORED: one
// sheet, stored not deflated, 1904 dates. BIG: one text entry, deflated
// with long matches. STYLED: fonts, fills, borders (rgb and theme colours),
// alignment, a comment, four validations (one a list from a range).
// PARTS: conditional formatting (nine rules, one of a kind there is not),
// a table with a totals row, an x14 sparkline group, a theme and document
// properties to carry over, a chart, a picture and a pivot to report.
// ODS: an OpenDocument spreadsheet, two tables: strings, floats, a formula
// with its cached value, a percentage, booleans, a date, a duration,
// repeated rows and columns (the last row a million of them, empty), a span.
export const BASIC = "{basic}";
export const STORED = "{stored}";
export const BIG = "{big}";
export const STYLED = "{styled}";
export const PARTS = "{parts}";
export const ODS = "{ods}";
''')
print(len(basic),len(stored),len(big))
