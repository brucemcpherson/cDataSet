# VBA Project: **cDataSet**
## VBA Module: **[cDataRow](/libraries/cDataRow.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (cDataSet) was automatically created on 26/03/2015 10:03:40 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cDataRow

---
VBA Procedure: **hidden**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get hidden()*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cDataSet*  

**no arguments required for this procedure**


---
VBA Procedure: **row**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get row() As Long*  

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
VBA Procedure: **where**  
Type: **Get**  
Returns: **Range**  
Scope: **Public**  
Description: ****  

*Public Property Get where() As Range*  

**no arguments required for this procedure**


---
VBA Procedure: **cell**  
Type: **Get**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Public**  
Description: ****  

*Public Property Get cell(sid As Variant, Optional complain As Boolean = False) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|False||
complain|Boolean|True| False|


---
VBA Procedure: **value**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get value(sid As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|False||


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
VBA Procedure: **refresh**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function refresh(Optional sid As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|True||


---
VBA Procedure: **Commit**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub Commit(Optional p As Variant, Optional sid As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|True||
sid|Variant|True||


---
VBA Procedure: **toString**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get toString(sid As Variant, Optional sFormat As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sid|Variant|False||
sFormat|String|True| vbNullString|


---
VBA Procedure: **create**  
Type: **Function**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Scope: **Public**  
Description: ****  

*Public Function create(dset As cDataSet, rDataRow As Range, nRow As Long, rv As Variant) As cDataRow*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dset|[cDataSet](/scripts/cDataSet_cls.md "cDataSet")|False||
rDataRow|Range|False||
nRow|Long|False||
rv|Variant|False||


---
VBA Procedure: **exists**  
Type: **Function**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Private**  
Description: ****  

*Private Function exists(sid As Variant) As cCell*  

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


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
