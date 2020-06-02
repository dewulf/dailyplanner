# VBA Project: **DailyPlanner**
## VBA Module: **[cTodo](/scripts/cTodo.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 7:31:28 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cTodo

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
VBA Procedure: **Move_From_Dailyplan_To_Nextdays**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Move_From_Dailyplan_To_Nextdays(dailyPlan As cDailyPlan, row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dailyPlan|[cDailyPlan](/scripts/cDailyPlan_cls.md "cDailyPlan")|False||
row|Integer|False||


---
VBA Procedure: **Copy_From_Dailyplan_To_Followups**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Copy_From_Dailyplan_To_Followups(dailyPlan As cDailyPlan, row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
dailyPlan|[cDailyPlan](/scripts/cDailyPlan_cls.md "cDailyPlan")|False||
row|Integer|False||


---
VBA Procedure: **Add_To_Dailyplan_Someday**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub Add_To_Dailyplan_Someday(someday As Date, row As Integer)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
someday|Date|False||
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
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
