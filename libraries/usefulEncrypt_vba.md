# VBA Project: **DailyPlanner**
## VBA Module: **[usefulEncrypt](/libraries/usefulEncrypt.vba "source is here")**
### Type: StdModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 8:22:30 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in usefulEncrypt

---
VBA Procedure: **encryptMessage**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function encryptMessage(ByVal TobeEncrypted As String, ByVal salt As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|String|False||


---
VBA Procedure: **decryptMessage**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function decryptMessage(ByVal encrypted As String, ByVal salt As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|String|False||


---
VBA Procedure: **encryptSha1**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function encryptSha1(ByVal keyString As String, ByVal str As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
ByVal|String|False||


---
VBA Procedure: **tob64**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function tob64(ByRef arrData() As Byte) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByRef|Variant|False||


---
VBA Procedure: **decodeBase64**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function decodeBase64(ByVal strData As String) As Byte()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
