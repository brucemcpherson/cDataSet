# VBA Project: **cDataSet**
## VBA Module: **[cDataColumn](/libraries/cDataColumn.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (cDataSet) was automatically created on 26/03/2015 10:03:40 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cDataColumn

---
VBA Procedure: **googleType**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get googleType() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **dirty**  
Type: **Get**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Property Get dirty() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **dirty**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Let dirty(p As Boolean)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Boolean|False||


---
VBA Procedure: **typeofColumn**  
Type: **Get**  
Returns: **eTypeofColumn**  
Scope: **Public**  
Description: ****  

*Public Property Get typeofColumn() As eTypeofColumn*  

**no arguments required for this procedure**


---
VBA Procedure: **typeofColumn**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Let typeofColumn(p As eTypeofColumn)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|eTypeofColumn|False||


---
VBA Procedure: **column**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get column() As Long*  

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
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cDataSet](/scripts/cDataSet_cls.md "cDataSet")**  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cDataSet*  

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

*Public Property Get cell(rowID As Variant) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|False||


---
VBA Procedure: **value**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get value(rowID As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|False||


---
VBA Procedure: **refresh**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function refresh(Optional rowID As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowID|Variant|True||


---
VBA Procedure: **filtered**  
Type: **Function**  
Returns: **Collection**  
Scope: **Public**  
Description: ****  

*Public Function filtered(v As Variant) As Collection*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||


---
VBA Procedure: **uniqueValues**  
Type: **Get**  
Returns: **Collection**  
Scope: **Public**  
Description: ****  

*Public Property Get uniqueValues(Optional es As eSort = eSortNone) As Collection*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
es|eSort|True| eSortNone|


---
VBA Procedure: **Commit**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub Commit(Optional p As Variant, Optional rowID As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|True||
rowID|Variant|True||


---
VBA Procedure: **values**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get values() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **find**  
Type: **Function**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Public**  
Description: ****  

*Public Function find(v As Variant) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||


---
VBA Procedure: **max**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function max() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **min**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function min() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **toString**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get toString(rowNum As Long, Optional sFormat As String = vbNullString) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rowNum|Long|False||
sFormat|String|True| vbNullString|


---
VBA Procedure: **create**  
Type: **Function**  
Returns: **[cDataColumn](/libraries/cDataColumn_cls.md "cDataColumn")**  
Scope: **Public**  
Description: ****  

*Public Function create(dset As cDataSet, hcell As cCell, ncol As Long) As cDataColumn*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dset|[cDataSet](/scripts/cDataSet_cls.md "cDataSet")|False||
hcell|[cCell](/libraries/cCell_cls.md "cCell")|False||
ncol|Long|False||


---
VBA Procedure: **exists**  
Type: **Function**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Private**  
Description: ****  

*Private Function exists(vCollect As Collection, sid As Variant) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
vCollect|Collection|False||
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
