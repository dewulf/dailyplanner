# VBA Project: **DailyPlanner**
## VBA Module: **[cStringChunker](/libraries/cStringChunker.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 8:22:31 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cStringChunker

---
VBA Procedure: **size**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get size() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **content**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get content() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getLeft**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get getLeft(howMany As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|False||


---
VBA Procedure: **getRight**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get getRight(howMany As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
howMany|Long|False||


---
VBA Procedure: **getMid**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get getMid(startPos As Long, Optional howMany As Long = -1) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
startPos|Long|False||
howMany|Long|True| -1|


---
VBA Procedure: **self**  
Type: **Get**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get self() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **clear**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function clear() As cStringChunker*  

**no arguments required for this procedure**


---
VBA Procedure: **add**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function add(addString As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
addString|String|False||


---
VBA Procedure: **addLine**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function addLine(addString As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
addString|String|False||


---
VBA Procedure: **insert**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function insert(Optional insertString As String = " ", Optional insertBefore As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
insertString|String|True| " "|
insertBefore|Long|True| 1|


---
VBA Procedure: **overWrite**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function overWrite(Optional overWriteString As String = " ", Optional overWriteAt As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
overWriteString|String|True| " "|
overWriteAt|Long|True| 1|


---
VBA Procedure: **Shift**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Shift(Optional startPos As Long = 1, Optional howManyChars As Long = 0, Optional replaceWith As String = vbNullString) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
startPos|Long|True| 1|
howManyChars|Long|True| 0|
replaceWith|String|True| vbNullString|


---
VBA Procedure: **chop**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function chop(Optional n As Long = 1) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
n|Long|True| 1|


---
VBA Procedure: **chopIf**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function chopIf(T As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
T|String|False||


---
VBA Procedure: **chopWhile**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function chopWhile(T As String) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
T|String|False||


---
VBA Procedure: **maxNumber**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function maxNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Long|False||
b|Long|False||


---
VBA Procedure: **minNumber**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function minNumber(a As Long, b As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Long|False||
b|Long|False||


---
VBA Procedure: **adjustSize**  
Type: **Function**  
Returns: **[cStringChunker](/libraries/cStringChunker_cls.md "cStringChunker")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function adjustSize(needMore As Long) As cStringChunker*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
needMore|Long|False||


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
