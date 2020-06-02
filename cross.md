# VBA Project: DailyPlanner
This cross reference list for repo (DailyPlanner) was automatically created on 6/2/2020 7:31:28 PM by VBAGit.For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")
You can see [library and dependency information here](dependencies.md)

###Below is a cross reference showing which modules and procedures reference which others
*module*|*proc*|*referenced by module*|*proc*
---|---|---|---
cBrowser||cGoogleCalendar|Get_Google_Calendar_Events
cBrowser||cGoogleCalendar|Get_Calendars
cJobject||cOauth2|googlePackage
cJobject||cOauth2|skeletonPackage
cOauth2||cGoogleCalendar|getGoogled
cregXLib||regXLib|rxMakeRxLib
cStringChunker||usefulStuff|basicStyle
cStringChunker||usefulStuff|tableStyle
cStringChunker||usefulStuff|includeJQuery
cStringChunker||usefulStuff|includeGoogleCallBack
cStringChunker||usefulStuff|jScriptTag
cStringChunker||usefulStuff|jDivAtMouse
cStringChunker||usefulStuff|encloseTag
JsonConverter|ParseJson|cGoogleCalendar|Get_Calendars
JsonConverter|ParseJson|cGoogleCalendar|Get_Google_Calendar_Events
regXLib|rxReplace|usefulcJobject|cleanGoogleWire
ShaFunc|SHA1TRUNC|cDailyPlan|Book_Spent_Time_To_Redmine
TimeZoneStuff|ConvertLocalToGMT|cGoogleCalendar|Get_Google_Calendar_Events
Tools|CONCATENATE_MULTIPLE|cDailyPlan|Book_Spent_Time_To_Redmine
Tools|Copy_Rowrange|cDailyPlan|Insert_New_Entry
Tools|Move_Row|cDailyPlan|Move_From_Todo_To_Someday
Tools|Move_Row|cDailyPlan|Move_To_Tomorrow
usefulcJobject|JSONParse|cOauth2|makeBasicGoogleConsole
usefulcJobject|JSONParse|cOauth2|getToken
usefulcJobject|JSONParse|cOauth2|getRegistryPackage
usefulcJobject|JSONParse|cOauth2|describeDialog
usefulcJobject|JSONStringify|cOauth2|setRegistryPackage
usefulEncrypt|decryptMessage|cOauth2|decrypt
usefulEncrypt|encryptMessage|cOauth2|encrypt
usefulStuff|Base64Encode|cBrowser|httpGET
