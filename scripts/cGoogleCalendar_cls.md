# VBA Project: **DailyPlanner**
## VBA Module: **[cGoogleCalendar](/scripts/cGoogleCalendar.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 8:22:32 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cGoogleCalendar

---
VBA Procedure: **Get_Auth_Header**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function Get_Auth_Header()*  

**no arguments required for this procedure**


---
VBA Procedure: **Get_Google_Calendar_Events**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub Get_Google_Calendar_Events(cal_date As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cal_date|String|False||


---
VBA Procedure: **Get_Calendars**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Get_Calendars() As String()*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
As|Variant|False||


---
VBA Procedure: **get_google_time**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function get_google_time(timestring As String) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
timestring|String|False||


---
VBA Procedure: **getGoogled**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getGoogled(scope As String, Optional replacementpackage As cJobject = Nothing, Optional clientID As String = vbNullString, Optional clientSecret As String = vbNullString, Optional complain As Boolean = True, Optional cloneFromeScope As String = vbNullString) As cOauth2*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
scope|String|False||
replacementpackage|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|
clientID|String|True| vbNullString|
clientSecret|String|True| vbNullString|
complain|Boolean|True| True|
cloneFromeScope|String|True| vbNullString|


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
