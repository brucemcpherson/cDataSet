# VBA Project: **cDataSet**
## VBA Module: **[cCell](/libraries/cCell.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (cDataSet) was automatically created on 26/03/2015 10:03:40 by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cCell

---
VBA Procedure: **row**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get row() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **column**  
Type: **Get**  
Returns: **Long**  
Scope: **Public**  
Description: ****  

*Public Property Get column() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cDataRow*  

**no arguments required for this procedure**


---
VBA Procedure: **myKey**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get myKey() As String*  

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
VBA Procedure: **refresh**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get refresh() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **toString**  
Type: **Get**  
Returns: **String**  
Scope: **Public**  
Description: ****  

*Public Property Get toString(Optional sFormat As String = vbNullString, Optional followFormat As Boolean = False, Optional deLocalize As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sFormat|String|True| vbNullString|
followFormat|Boolean|True| False|
deLocalize|Boolean|True| False|


---
VBA Procedure: **value**  
Type: **Get**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Property Get value() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **value**  
Type: **Let**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Property Let value(p As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|False||


---
VBA Procedure: **needSwap**  
Type: **Function**  
Returns: **Boolean**  
Scope: **Public**  
Description: ****  

*Public Function needSwap(cc As cCell, e As eSort) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cc|[cCell](/libraries/cCell_cls.md "cCell")|False||
e|eSort|False||


---
VBA Procedure: **Commit**  
Type: **Function**  
Returns: **Variant**  
Scope: **Public**  
Description: ****  

*Public Function Commit(Optional p As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|True||


---
VBA Procedure: **create**  
Type: **Function**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Scope: **Public**  
Description: ****  

*Public Function create(par As cDataRow, colNum As Long, rCell As Range, Optional v As Variant) As cCell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
par|[cDataRow](/libraries/cDataRow_cls.md "cDataRow")|False||
colNum|Long|False||
rCell|Range|False||
v|Variant|True||


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**
