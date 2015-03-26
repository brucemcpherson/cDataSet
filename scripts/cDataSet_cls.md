# VBA Project: **cDataSet**
## VBA Module: **[cDataSet](/scripts/cDataSet.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (cDataSet) was automatically created on 26/03/2015 10:03:40 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cDataSet

---
VBA Procedure: **self**  
Type: **Get**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Property Get self() As cDataSet*  

**no arguments required for this procedure**


---
VBA Procedure: **activeListObject**  
Type: **Get**  
Returns: **ListObject**  
Scope: **Public**  
Description: ****  

*Public Property Get activeListObject() As ListObject*  

**no arguments required for this procedure**


---
VBA Procedure: **intersectListObject**  
Type: **Function**  
Returns: **ListObject**  
Scope: **Private**  
Description: ****  

*Private Function intersectListObject(r As Range) As ListObject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **makeListObject**  
Type: **Function**  
Returns: **ListObject**  
Scope: **Public**  
Description: ****  

*Public Function makeListObject(Optional sName As String = vbNullString) As ListObject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|True| vbNullString|


---
VBA Procedure: **visibleRowsCount**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get visibleRowsCount() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **recordFilter**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Property Get recordFilter() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **keyColumn**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get keyColumn() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **keepFresh**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Property Get keepFresh() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cDataSets](/libraries/cDataSets_cls.md "cDataSets")**  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cDataSets*  

**no arguments required for this procedure**


---
VBA Procedure: **name**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get name() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **rows**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  
Description: ****  

*Public Property Get rows() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **columns**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  
Description: ****  

*Public Property Get columns() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **headings**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  
Description: ****  

*Public Property Get headings() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **where**  
Type: **Get**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Public Property Get where() As Range*  

**no arguments required for this procedure**


---
VBA Procedure: **headingRow**  
Type: **Get**  
Returns: **[cHeadingRow](/libraries/cHeadingRow_cls.md "cHeadingRow")**  
Scope: **Public**  
Description: ****  

*Public Property Get headingRow() As cHeadingRow*  

**no arguments required for this procedure**


---
VBA Procedure: **headingRow**  
Type: **Set**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Set headingRow(p As cHeadingRow)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|[cHeadingRow](/libraries/cHeadingRow_cls.md "cHeadingRow")|False||


---
VBA Procedure: **cell**  
Type: **Get**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Public**  
Description: ****  

*Public Property Get cell(rowID As Variant, sid As Variant) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|False||
sid|Variant|False||


---
VBA Procedure: **isCellTrue**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Property Get isCellTrue(rowID As Variant, sid As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|False||
sid|Variant|False||


---
VBA Procedure: **value**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get value(rowID As Variant, sid As Variant, Optional complain As Boolean = True) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|False||
sid|Variant|False||
complain|Boolean|True| True|


---
VBA Procedure: **letValue**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function letValue(p As Variant, rowID As Variant, sid As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|False||
rowID|Variant|False||
sid|Variant|False||


---
VBA Procedure: **toString**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get toString(rowID As Variant, sid As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|False||
sid|Variant|False||


---
VBA Procedure: **row**  
Type: **Get**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Scope: **Public**  
Description: ****  

*Public Property Get row(rowID As Variant) As cDataRow*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|False||


---
VBA Procedure: **column**  
Type: **Get**  
Returns: **[cDataColumn](/libraries/cDataColumn_cls.md "cDataColumn")**  
Scope: **Public**  
Description: ****  

*Public Property Get column(sid As Variant) As cDataColumn*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|False||


---
VBA Procedure: **jObject**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Scope: **Public**  
Description: ****  

*Public Property Get jObject(Optional jSonConv As eJsonConv = eJsonConvPropertyNames, Optional datesToIso As Boolean = False, Optional includeParseTypes As Boolean = False, Optional includeDataSetName As Boolean = True, Optional dataSetName As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
jSonConv|eJsonConv|True| eJsonConvPropertyNames|
datesToIso|Boolean|True| False|
includeParseTypes|Boolean|True| False|
includeDataSetName|Boolean|True| True|
dataSetName|String|True| vbNullString|


---
VBA Procedure: **refresh**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function refresh(Optional rowID As Variant, Optional sid As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|True||
sid|Variant|True||


---
VBA Procedure: **Commit**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub Commit(Optional p As Variant, Optional rowID As Variant, Optional sid As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|True||
rowID|Variant|True||
sid|Variant|True||


---
VBA Procedure: **create**  
Type: **Function**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Private**  
Description: ****  

*Private Function create(rp As Range, Optional sn As String = vbNullString, Optional blab As Boolean = False, Optional keepFresh As Boolean = False, Optional stopAtFirstEmptyRow = True, Optional sKey As String = vbNullString, Optional maxDataRows As Long = 0) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rp|Range|False||
sn|String|True| vbNullString|
blab|Boolean|True| False|
keepFresh|Boolean|True| False|
stopAtFirstEmptyRow|Variant|True||
sKey|String|True| vbNullString|
maxDataRows|Long|True| 0|


---
VBA Procedure: **populateJSON**  
Type: **Function**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Function populateJSON(job As cJobject, rstart As Range, Optional wClearContents As Boolean = True, Optional stopAtFirstEmptyRow As Boolean = True) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
rstart|Range|False||
wClearContents|Boolean|True| True|
stopAtFirstEmptyRow|Boolean|True| True|


---
VBA Procedure: **populateGoogleWire**  
Type: **Function**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Function populateGoogleWire(sWire As String, rstart As Range, Optional wClearContents As Boolean = True, Optional stopAtFirstEmptyRow As Boolean = True) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sWire|String|False||
rstart|Range|False||
wClearContents|Boolean|True| True|
stopAtFirstEmptyRow|Boolean|True| True|


---
VBA Procedure: **rePopulate**  
Type: **Function**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Function rePopulate() As cDataSet*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**


---
VBA Procedure: **load**  
Type: **Function**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Function load(sheetName As String, Optional parameterBlock As String = vbNullString) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sheetName|String|False||
parameterBlock|String|True| vbNullString|


---
VBA Procedure: **populateData**  
Type: **Function**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Function populateData(Optional rstart As Range = Nothing, Optional keepFresh As Boolean = False, Optional sn As String = vbNullString, Optional blab As Boolean = False, Optional blockstarts As String = vbNullString, Optional ps As cDataSets, Optional bLikely As Boolean = False, Optional sKey As String = vbNullString, Optional maxDataRows As Long = 0, Optional stopAtFirstEmptyRow As Boolean = True, Optional brecordFilter As Boolean = False) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rstart|Range|True| Nothing|
keepFresh|Boolean|True| False|
sn|String|True| vbNullString|
blab|Boolean|True| False|
blockstarts|String|True| vbNullString|
ps|[cDataSets](/libraries/cDataSets_cls.md "cDataSets")|True||
bLikely|Boolean|True| False|
sKey|String|True| vbNullString|
maxDataRows|Long|True| 0|
stopAtFirstEmptyRow|Boolean|True| True|
brecordFilter|Boolean|True| False|


---
VBA Procedure: **values**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get values(Optional bIncludeKey = False) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
bIncludeKey|Variant|True||


---
VBA Procedure: **find**  
Type: **Function**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Public**  
Description: ****  

*Public Function find(v As Variant, Optional bIncludeKey = False) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||
bIncludeKey|Variant|True||


---
VBA Procedure: **max**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function max(Optional bIncludeKey = False) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
bIncludeKey|Variant|True||


---
VBA Procedure: **min**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function min(Optional bIncludeKey = False) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
bIncludeKey|Variant|True||


---
VBA Procedure: **flushDirtyColumns**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function flushDirtyColumns()*  

**no arguments required for this procedure**


---
VBA Procedure: **bigCommit**  
Type: **Function**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Function bigCommit(Optional rout As Range = Nothing, Optional clearWs As Boolean = False, Optional headOrderArray As Variant = Empty, Optional filterHead As String = vbNullString, Optional filterValue As Variant = Empty, Optional filterApproximate As Boolean = True, Optional outputHeadings As Boolean = True, Optional filterUpperValue) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rout|Range|True| Nothing|
clearWs|Boolean|True| False|
headOrderArray|Variant|True| Empty|
filterHead|String|True| vbNullString|
filterValue|Variant|True| Empty|
filterApproximate|Boolean|True| True|
outputHeadings|Boolean|True| True|
filterUpperValue|Variant|True||


---
VBA Procedure: **filterOk**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Private**  
Description: ****  

*Private Function filterOk(dr As cDataRow, filterCol As Long, filterValue As Variant, filterApproximate As Boolean, Optional filterUpperValue As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dr|[cDataRow](/libraries/cDataRow_cls.md "cDataRow")|False||
filterCol|Long|False||
filterValue|Variant|False||
filterApproximate|Boolean|False||
filterUpperValue|Variant|True||


---
VBA Procedure: **exists**  
Type: **Function**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Scope: **Private**  
Description: ****  

*Private Function exists(sid As Variant) As cDataRow*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|False||


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**
