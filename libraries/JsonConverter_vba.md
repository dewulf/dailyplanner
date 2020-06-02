# VBA Project: **DailyPlanner**
## VBA Module: **[JsonConverter](/libraries/JsonConverter.vba "source is here")**
### Type: StdModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 7:31:26 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in JsonConverter

---
VBA Procedure: **ParseJson**  
Type: **Function**  
Returns: **Object**  
Return description: **(Dictionary or Collection)**  
Scope: **Public**  
Description: ****  

*Public Function ParseJson(ByVal JsonString As String) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **ConvertToJson**  
Type: **Function**  
Returns: **String**  
Return description: **''**  
Scope: **Public**  
Description: ****  

*Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||
ByVal|Variant|True||
ByVal|Variant|True||


---
VBA Procedure: **json_ParseObject**  
Type: **Function**  
Returns: **Dictionary**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_ParseObject(json_string As String, ByRef json_Index As Long) As Dictionary*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||


---
VBA Procedure: **json_ParseArray**  
Type: **Function**  
Returns: **Collection**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_ParseArray(json_string As String, ByRef json_Index As Long) As Collection*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||


---
VBA Procedure: **json_ParseValue**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_ParseValue(json_string As String, ByRef json_Index As Long) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||


---
VBA Procedure: **json_ParseString**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_ParseString(json_string As String, ByRef json_Index As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||


---
VBA Procedure: **json_ParseNumber**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_ParseNumber(json_string As String, ByRef json_Index As Long) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||


---
VBA Procedure: **json_ParseKey**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_ParseKey(json_string As String, ByRef json_Index As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||


---
VBA Procedure: **json_IsUndefined**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||


---
VBA Procedure: **json_Encode**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_Encode(ByVal json_Text As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||


---
VBA Procedure: **json_Peek**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_Peek(json_string As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByVal|Long|False||
json_NumberOfCharacters|Long|True| 1|


---
VBA Procedure: **json_SkipSpaces**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub json_SkipSpaces(json_string As String, ByRef json_Index As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||


---
VBA Procedure: **json_StringIsLargeNumber**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_StringIsLargeNumber(json_string As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|Variant|False||


---
VBA Procedure: **json_ParseErrorMessage**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_ParseErrorMessage(json_string As String, ByRef json_Index As Long, ErrorMessage As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
json_string|String|False||
ByRef|Long|False||
ErrorMessage|String|False||


---
VBA Procedure: **json_BufferAppend**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub json_BufferAppend(ByRef json_Buffer As String, ByRef json_Append As Variant, ByRef json_BufferPosition As Long, ByRef json_BufferLength As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByRef|String|False||
ByRef|Variant|False||
ByRef|Long|False||
ByRef|Long|False||


---
VBA Procedure: **json_BufferToString**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByRef|String|False||
ByVal|Long|False||


---
VBA Procedure: **ParseUtc**  
Type: **Function**  
Returns: **Date**  
Return description: **Local date**  
Scope: **Public**  
Description: ****  

*Public Function ParseUtc(utc_UtcDate As Date) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_UtcDate|Date|False||


---
VBA Procedure: **ConvertToUtc**  
Type: **Function**  
Returns: **Date**  
Return description: **UTC date**  
Scope: **Public**  
Description: ****  

*Public Function ConvertToUtc(utc_LocalDate As Date) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_LocalDate|Date|False||' @return {Date} UTC date


---
VBA Procedure: **ParseIso**  
Type: **Function**  
Returns: **Date**  
Return description: **Local date**  
Scope: **Public**  
Description: ****  

*Public Function ParseIso(utc_IsoString As String) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_IsoString|String|False||' @return {Date} Local date


---
VBA Procedure: **ConvertToIso**  
Type: **Function**  
Returns: **String**  
Return description: **ISO 8601 string**  
Scope: **Public**  
Description: ****  

*Public Function ConvertToIso(utc_LocalDate As Date) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_LocalDate|Date|False||' @return {Date} ISO 8601 string


---
VBA Procedure: **utc_ConvertDate**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_Value|Date|False||
utc_ConvertToUtc|Boolean|True| False|


---
VBA Procedure: **utc_ExecuteInShell**  
Type: **Function**  
Returns: **utc_ShellResult**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_ShellCommand|String|False||


---
VBA Procedure: **utc_DateToSystemTime**  
Type: **Function**  
Returns: **utc_SYSTEMTIME**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_Value|Date|False||


---
VBA Procedure: **utc_SystemTimeToDate**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
utc_Value|utc_SYSTEMTIME|False||
