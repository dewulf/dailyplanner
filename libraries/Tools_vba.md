# VBA Project: **DailyPlanner**
## VBA Module: **[Tools](/libraries/Tools.vba "source is here")**
### Type: StdModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 8:22:29 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in Tools

---
VBA Procedure: **CONCATENATE_MULTIPLE**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function CONCATENATE_MULTIPLE(ref As Range, Separator As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ref|Range|False||
Separator|String|False||


---
VBA Procedure: **Move_Row**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Move_Row(ws_from As Worksheet, row_from As Integer, ws_to As Worksheet, row_to As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws_from|Worksheet|False||
row_from|Integer|False||
ws_to|Worksheet|False||
row_to|Integer|False||


---
VBA Procedure: **Copy_Row**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Copy_Row(ws_from As Worksheet, row_from As Integer, ws_to As Worksheet, row_to As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws_from|Worksheet|False||
row_from|Integer|False||
ws_to|Worksheet|False||
row_to|Integer|False||


---
VBA Procedure: **Copy_Rowrange**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Copy_Rowrange(ws_from As Worksheet, row_range_from As Range, ws_to As Worksheet, row_to As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws_from|Worksheet|False||
row_range_from|Range|False||
ws_to|Worksheet|False||
row_to|Integer|False||


---
VBA Procedure: **Update_Task_Desc**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Update_Task_Desc(ws_name As String, proj_rng As String, task_rng As String, project_name As String, old_task_desc As String, new_task_desc As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws_name|String|False||
proj_rng|String|False||
task_rng|String|False||
project_name|String|False||
old_task_desc|String|False||
new_task_desc|String|False||


---
VBA Procedure: **Get_Final_Row**  
Type: **Function**  
Returns: **Integer**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Get_Final_Row(search_col As Range) As Integer*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
search_col|Range|False||


---
VBA Procedure: **IsArrayAllocated**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function IsArrayAllocated(arr As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
arr|Variant|False||
