<div align="center">

## logging events \- extended functionality


</div>

### Description

This is a module with one function that will give you some more functionality than what you can do with the App.LogEvent method in VB6.

1. You will be able to specify the EventLog (Application, Security or System)

2. You are able to specify the Source (Your own Application identifier instead of the VBRuntime)

This module is also used in the ExeptionHandler dll. that is posted earlier today.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Edwin Vermeer\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/edwin-vermeer.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/edwin-vermeer-logging-events-extended-functionality__1-62772/archive/master.zip)





### Source Code

```

' Using this module is easy. Just Call it like this:
' writelog EventLog_Application,"My Special APP",vbLogEventTypeError, "Oep, Something went wrong :)"
'Functions and type for logging events
Option Explicit
Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Private Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Integer, ByVal wCategory As Integer, ByVal dwEventID As Long, ByVal lpUserSid As Any, ByVal wNumStrings As Integer, ByVal dwDataSize As Long, plpStrings As Long, lpRawData As Any) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Enum EventLog
  EventLog_Application
  EventLog_Security
  EventLog_System
End Enum
' This function will give you some more functionality over the App.LogEvent method.
' You will be able to specify the EventLog (Application, Security or System)
' And you are able to specify the Source (Your own Application identifier instead of the VBRuntime)
Public Function WriteLog(intEventLogID As EventLog, strEventSource As String, intEventType As LogEventTypeConstants, strEventString As String) As Boolean
1   On Error GoTo ErrHandler
2 Dim intEventStringsCount As Integer
3 Dim hEventLog As Long
4 Dim hMsgs As Long
5 Dim lngEventStringSize As Long
6 Dim objRegistry As Object
7 Dim strEventLogDescription As String
8   WriteLog = False
  ' In case we have a new source we make sure it finds the VBRuntime DLL for handeling the event description.
9   Select Case intEventLogID
  Case EventLog_Application
10     strEventLogDescription = "Application"
11   Case EventLog_Security
12     strEventLogDescription = "Security"
13   Case EventLog_System
14     strEventLogDescription = "System"
15   End Select
16   Set objRegistry = CreateObject("Wscript.Shell")
17   objRegistry.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\" & strEventLogDescription & "\" & strEventSource & "\EventMessageFile", objRegistry.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\VBRuntime\EventMessageFile"), "REG_SZ"
18   objRegistry.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\" & strEventLogDescription & "\" & strEventSource & "\TypesSupported", 4, "REG_DWORD"
  ' Set the event source and report the event
19   hEventLog = RegisterEventSource("", strEventSource)
20   strEventString = ":" & vbCrLf & vbCrLf & strEventString
21   lngEventStringSize = Len(strEventString) + 1
22   hMsgs = GlobalAlloc(&H40, lngEventStringSize)
23   CopyMemory ByVal hMsgs, ByVal strEventString, lngEventStringSize
24   intEventStringsCount = 1
25   If ReportEvent(hEventLog, intEventType, 0, 1, 0&, intEventStringsCount, lngEventStringSize, hMsgs, hMsgs) = 0 Then
26     WriteLog = True
27   End If
28   GlobalFree (hMsgs)
29   DeregisterEventSource (hEventLog)
30 Exit Function
31 ErrHandler:
End Function
```

