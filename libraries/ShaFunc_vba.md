# VBA Project: **DailyPlanner**
## VBA Module: **[ShaFunc](/libraries/ShaFunc.vba "source is here")**
### Type: StdModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 7:31:26 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in ShaFunc

---
VBA Procedure: **HexDefaultSHA1**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function HexDefaultSHA1(message() As Byte) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
message|Variant|False||


---
VBA Procedure: **HexSHA1**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function HexSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
message|Variant|False||
ByVal|Long|False||
ByVal|Long|False||
ByVal|Long|False||
ByVal|Long|False||


---
VBA Procedure: **DefaultSHA1**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub DefaultSHA1(message() As Byte, h1 As Long, h2 As Long, H3 As Long, H4 As Long, H5 As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
message|Variant|False||
h1|Long|False||
h2|Long|False||
H3|Long|False||
H4|Long|False||
H5|Long|False||


---
VBA Procedure: **xSHA1**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub xSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, h1 As Long, h2 As Long, H3 As Long, H4 As Long, H5 As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
message|Variant|False||
ByVal|Long|False||
ByVal|Long|False||
ByVal|Long|False||
ByVal|Long|False||
h1|Long|False||
h2|Long|False||
H3|Long|False||
H4|Long|False||
H5|Long|False||


---
VBA Procedure: **U32Add**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function U32Add(ByVal a As Long, ByVal b As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Long|False||
ByVal|Long|False||


---
VBA Procedure: **U32ShiftLeft3**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function U32ShiftLeft3(ByVal a As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Long|False||


---
VBA Procedure: **U32ShiftRight29**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function U32ShiftRight29(ByVal a As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Long|False||


---
VBA Procedure: **U32RotateLeft1**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function U32RotateLeft1(ByVal a As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Long|False||


---
VBA Procedure: **U32RotateLeft5**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function U32RotateLeft5(ByVal a As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Long|False||


---
VBA Procedure: **U32RotateLeft30**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function U32RotateLeft30(ByVal a As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Long|False||


---
VBA Procedure: **DecToHex5**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function DecToHex5(ByVal h1 As Long, ByVal h2 As Long, ByVal H3 As Long, ByVal H4 As Long, ByVal H5 As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Long|False||
ByVal|Long|False||
ByVal|Long|False||
ByVal|Long|False||
ByVal|Long|False||


---
VBA Procedure: **SHA1TRUNC**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function SHA1TRUNC(str)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
str|Variant|False||
