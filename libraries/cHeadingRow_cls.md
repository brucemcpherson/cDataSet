# VBA Project: **cDataSet**
## VBA Module: **[cHeadingRow](/libraries/cHeadingRow.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (cDataSet) was automatically created on 26/03/2015 10:03:40 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cHeadingRow

---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cDataSet*  

**no arguments required for this procedure**


---
VBA Procedure: **dataRow**  
Type: **Get**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Scope: **Public**  
Description: ****  

*Public Property Get dataRow() As cDataRow*  

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
VBA Procedure: **create**  
Type: **Function**  
Returns: **[cHeadingRow](/libraries/cHeadingRow_cls.md "cHeadingRow")**  
Scope: **Public**  
Description: ****  

*Public Function create(dset As cDataSet, rHeading As Range, Optional keepFresh As Boolean = False) As cHeadingRow*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dset|[cDataSet](/scripts/cDataSet_cls.md "cDataSet")|False||
rHeading|Range|False||
keepFresh|Boolean|True| False|


---
VBA Procedure: **exists**  
Type: **Function**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Public**  
Description: ****  

*Public Function exists(s As String) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **headingList**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get headingList() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **validate**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Function validate(complain As Boolean, ParamArray args() As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
complain|Boolean|False||
ParamArray|Variant|False||


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
