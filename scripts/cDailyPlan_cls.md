# VBA Project: **DailyPlanner**
## VBA Module: **[cDailyPlan](/scripts/cDailyPlan.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 8:22:31 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cDailyPlan

---
VBA Procedure: **Ws**  
Type: **Get**  
Returns: **Worksheet**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Ws() As Worksheet*  

**no arguments required for this procedure**


---
VBA Procedure: **Col_Date**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Col_Date()*  

**no arguments required for this procedure**


---
VBA Procedure: **Col_Project**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Col_Project()*  

**no arguments required for this procedure**


---
VBA Procedure: **Col_Task**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Col_Task()*  

**no arguments required for this procedure**


---
VBA Procedure: **Col_Done**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Col_Done()*  

**no arguments required for this procedure**


---
VBA Procedure: **set_followup**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub set_followup(row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
row|Integer|False||


---
VBA Procedure: **Do_Analytics**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Do_Analytics(row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
row|Integer|False||


---
VBA Procedure: **Insert_New_Entry**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Insert_New_Entry(row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
row|Integer|False||


---
VBA Procedure: **Move_To_Tomorrow**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Move_To_Tomorrow(row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
row|Integer|False||


---
VBA Procedure: **Move_To_Todo_Nextdays**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Move_To_Todo_Nextdays(row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
row|Integer|False||


---
VBA Procedure: **Copy_To_Todo_Followups**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Copy_To_Todo_Followups(row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
row|Integer|False||


---
VBA Procedure: **Move_From_Todo_To_Someday**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Move_From_Todo_To_Someday(todo As cTodo, someday As Date, row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
todo|[cTodo](/scripts/cTodo_cls.md "cTodo")|False||
someday|Date|False||
row|Integer|False||


---
VBA Procedure: **Conditional_Formating**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Conditional_Formating()*  

**no arguments required for this procedure**


---
VBA Procedure: **Book_Spent_Time_To_Redmine**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Book_Spent_Time_To_Redmine(book_date As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
book_date|String|False||


---
VBA Procedure: **Insert_Cal_Entry**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Insert_Cal_Entry(ByVal cal_event As Object, cur_date_rows() As Integer, cal_engine_id As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Object|False||
cur_date_rows|Variant|False||
cal_engine_id|String|False||


---
VBA Procedure: **Update_Cal_Entry**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Update_Cal_Entry(ByVal cal_event As Object, cur_date_rows() As Integer, cal_engine_id As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Object|False||
cur_date_rows|Variant|False||
cal_engine_id|String|False||


---
VBA Procedure: **Remove_Cal_Entries**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Remove_Cal_Entries(iCalUIDs() As String, cur_date_rows() As Integer, cal_engine_id As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
iCalUIDs|Variant|False||
cur_date_rows|Variant|False||
cal_engine_id|String|False||


---
VBA Procedure: **Get_Rows_For_Date**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Get_Rows_For_Date(cal_date As String) As Integer()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cal_date|String|False||


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
