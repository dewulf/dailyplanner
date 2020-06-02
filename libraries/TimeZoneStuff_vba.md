# VBA Project: **DailyPlanner**
## VBA Module: **[TimeZoneStuff](/libraries/TimeZoneStuff.vba "source is here")**
### Type: StdModule  

This procedure list for repo (DailyPlanner) was automatically created on 6/2/2020 8:22:29 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in TimeZoneStuff

---
VBA Procedure: **ConvertLocalToGMT**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function ConvertLocalToGMT(Optional LocalTime As Date, Optional AdjustForDST As Boolean = False) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
LocalTime|Date|True||
AdjustForDST|Boolean|True| False|


---
VBA Procedure: **GetLocalTimeFromGMT**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function GetLocalTimeFromGMT(Optional StartTime As Date) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
StartTime|Date|True||


---
VBA Procedure: **SystemTimeToVBTime**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SystemTimeToVBTime(SysTime As SYSTEMTIME) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
SysTime|SYSTEMTIME|False||


---
VBA Procedure: **LocalOffsetFromGMT**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function LocalOffsetFromGMT(Optional AsHours As Boolean = False, Optional AdjustForDST As Boolean = False) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
AsHours|Boolean|True| False|
AdjustForDST|Boolean|True| False|


---
VBA Procedure: **DaylightTime**  
Type: **Function**  
Returns: **TIME_ZONE**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function DaylightTime() As TIME_ZONE*  

**no arguments required for this procedure**
