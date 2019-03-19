Attribute VB_Name = "VBIB32"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 32-bit Visual Basic Language Interface
' Version 1.7
' Copyright 1998 National Instruments Corporation.
' All Rights Reserved.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This module contains the subroutine declarations,
'   function declarations and constants required to use
'   the National Instruments GPIB Dynamic Link Library
'   (DLL) for controlling IEEE-488 instrumentation.  This
'   file must be 'added' to your Visual Basic project
'   (by choosing Add File from the File menu or pressing
'   CTRL+F12) so that you can access the NI-488.2
'   subroutines and functions.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NI-488.2 DLL entry function declarations

Declare Function ibask32 Lib "Gpib-32.dll" Alias "ibask" (ByVal ud As Long, ByVal opt As Long, value As Long) As Long
Declare Function ibbna32 Lib "Gpib-32.dll" Alias "ibbnaA" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibcac32 Lib "Gpib-32.dll" Alias "ibcac" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibclr32 Lib "Gpib-32.dll" Alias "ibclr" (ByVal ud As Long) As Long
Declare Function ibcmd32 Lib "Gpib-32.dll" Alias "ibcmd" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibcmda32 Lib "Gpib-32.dll" Alias "ibcmda" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibconfig32 Lib "Gpib-32.dll" Alias "ibconfig" (ByVal ud As Long, ByVal opt As Long, ByVal v As Long) As Long
Declare Function ibdev32 Lib "Gpib-32.dll" Alias "ibdev" (ByVal bdid As Long, ByVal pad As Long, ByVal sad As Long, ByVal tmo As Long, ByVal eot As Long, ByVal eos As Long) As Long
Declare Function ibdma32 Lib "Gpib-32.dll" Alias "ibdma" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibeos32 Lib "Gpib-32.dll" Alias "ibeos" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibeot32 Lib "Gpib-32.dll" Alias "ibeot" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibfind32 Lib "Gpib-32.dll" Alias "ibfindA" (sstr As Any) As Long
Declare Function ibgts32 Lib "Gpib-32.dll" Alias "ibgts" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibist32 Lib "Gpib-32.dll" Alias "ibist" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function iblines32 Lib "Gpib-32.dll" Alias "iblines" (ByVal ud As Long, v As Long) As Long
Declare Function ibln32 Lib "Gpib-32.dll" Alias "ibln" (ByVal ud As Long, ByVal pad As Long, ByVal sad As Long, ln As Long) As Long
Declare Function ibloc32 Lib "Gpib-32.dll" Alias "ibloc" (ByVal ud As Long) As Long
Declare Function iblock32 Lib "Gpib-32.dll" Alias "iblock" (ByVal ud As Long) As Long
Declare Function ibonl32 Lib "Gpib-32.dll" Alias "ibonl" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibpad32 Lib "Gpib-32.dll" Alias "ibpad" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibpct32 Lib "Gpib-32.dll" Alias "ibpct" (ByVal ud As Long) As Long
Declare Function ibppc32 Lib "Gpib-32.dll" Alias "ibppc" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibrd32 Lib "Gpib-32.dll" Alias "ibrd" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrda32 Lib "Gpib-32.dll" Alias "ibrda" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrdf32 Lib "Gpib-32.dll" Alias "ibrdfA" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrpp32 Lib "Gpib-32.dll" Alias "ibrpp" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrsc32 Lib "Gpib-32.dll" Alias "ibrsc" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibrsp32 Lib "Gpib-32.dll" Alias "ibrsp" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrsv32 Lib "Gpib-32.dll" Alias "ibrsv" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibsad32 Lib "Gpib-32.dll" Alias "ibsad" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibsic32 Lib "Gpib-32.dll" Alias "ibsic" (ByVal ud As Long) As Long
Declare Function ibsre32 Lib "Gpib-32.dll" Alias "ibsre" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibstop32 Lib "Gpib-32.dll" Alias "ibstop" (ByVal ud As Long) As Long
Declare Function ibtmo32 Lib "Gpib-32.dll" Alias "ibtmo" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibtrg32 Lib "Gpib-32.dll" Alias "ibtrg" (ByVal ud As Long) As Long
Declare Function ibunlock32 Lib "Gpib-32.dll" Alias "ibunlock" (ByVal ud As Long) As Long
Declare Function ibwait32 Lib "Gpib-32.dll" Alias "ibwait" (ByVal ud As Long, ByVal mask As Long) As Long
Declare Function ibwrt32 Lib "Gpib-32.dll" Alias "ibwrt" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrta32 Lib "Gpib-32.dll" Alias "ibwrta" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrtf32 Lib "Gpib-32.dll" Alias "ibwrtfA" (ByVal ud As Long, sstr As Any) As Long
Declare Sub AllSpoll32 Lib "Gpib-32.dll" Alias "AllSpoll" (ByVal ud As Long, arg1 As Any, arg2 As Any)
Declare Sub DevClear32 Lib "Gpib-32.dll" Alias "DevClear" (ByVal ud As Long, ByVal v As Long)
Declare Sub DevClearList32 Lib "Gpib-32.dll" Alias "DevClearList" (ByVal ud As Long, arg1 As Any)
Declare Sub EnableLocal32 Lib "Gpib-32.dll" Alias "EnableLocal" (ByVal ud As Long, arg1 As Any)
Declare Sub EnableRemote32 Lib "Gpib-32.dll" Alias "EnableRemote" (ByVal ud As Long, arg1 As Any)
Declare Sub FindLstn32 Lib "Gpib-32.dll" Alias "FindLstn" (ByVal ud As Long, arg1 As Any, arg2 As Any, ByVal limit As Long)
Declare Sub FindRQS32 Lib "Gpib-32.dll" Alias "FindRQS" (ByVal ud As Long, arg1 As Any, result As Long)
Declare Sub PassControl32 Lib "Gpib-32.dll" Alias "PassControl" (ByVal ud As Long, ByVal addr As Long)
Declare Sub PPoll32 Lib "Gpib-32.dll" Alias "PPoll" (ByVal ud As Long, result As Long)
Declare Sub PPollConfig32 Lib "Gpib-32.dll" Alias "PPollConfig" (ByVal ud As Long, ByVal addr As Long, ByVal line As Long, ByVal sense As Long)
Declare Sub PPollUnconfig32 Lib "Gpib-32.dll" Alias "PPollUnconfig" (ByVal ud As Long, arg1 As Any)
Declare Sub RcvRespMsg32 Lib "Gpib-32.dll" Alias "RcvRespMsg" (ByVal ud As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReadStatusByte32 Lib "Gpib-32.dll" Alias "ReadStatusByte" (ByVal ud As Long, ByVal addr As Long, result As Long)
Declare Sub Receive32 Lib "Gpib-32.dll" Alias "Receive" (ByVal ud As Long, ByVal addr As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReceiveSetup32 Lib "Gpib-32.dll" Alias "ReceiveSetup" (ByVal ud As Long, ByVal addr As Long)
Declare Sub ResetSys32 Lib "Gpib-32.dll" Alias "ResetSys" (ByVal ud As Long, arg1 As Any)
Declare Sub Send32 Lib "Gpib-32.dll" Alias "Send" (ByVal ud As Long, ByVal addr As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendCmds32 Lib "Gpib-32.dll" Alias "SendCmds" (ByVal ud As Long, sstr As Any, ByVal cnt As Long)
Declare Sub SendDataBytes32 Lib "Gpib-32.dll" Alias "SendDataBytes" (ByVal ud As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendIFC32 Lib "Gpib-32.dll" Alias "SendIFC" (ByVal ud As Long)
Declare Sub SendList32 Lib "Gpib-32.dll" Alias "SendList" (ByVal ud As Long, arg1 As Any, arg2 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendLLO32 Lib "Gpib-32.dll" Alias "SendLLO" (ByVal ud As Long)
Declare Sub SendSetup32 Lib "Gpib-32.dll" Alias "SendSetup" (ByVal ud As Long, arg1 As Any)
Declare Sub SetRWLS32 Lib "Gpib-32.dll" Alias "SetRWLS" (ByVal ud As Long, arg1 As Any)
Declare Sub TestSys32 Lib "Gpib-32.dll" Alias "TestSys" (ByVal ud As Long, arg1 As Any, arg2 As Any)
Declare Sub Trigger32 Lib "Gpib-32.dll" Alias "Trigger" (ByVal ud As Long, ByVal addr As Long)
Declare Sub TriggerList32 Lib "Gpib-32.dll" Alias "TriggerList" (ByVal ud As Long, arg1 As Any)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIB global variables

Declare Function RegisterGpibGlobalsForThread Lib "Gpib-32.dll" (Longibsta As Long, Longiberr As Long, Longibcnt As Long, ibcntl As Long) As Long
Declare Function UnregisterGpibGlobalsForThread Lib "Gpib-32.dll" () As Long
Declare Function ThreadIbsta32 Lib "Gpib-32.dll" Alias "ThreadIbsta" () As Long
Declare Function ThreadIbcnt32 Lib "Gpib-32.dll" Alias "ThreadIbcnt" () As Long
Declare Function ThreadIbcntl32 Lib "Gpib-32.dll" Alias "ThreadIbcntl" () As Long
Declare Function ThreadIberr32 Lib "Gpib-32.dll" Alias "ThreadIberr" () As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIBnotify OLE control

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIB-ENET functions

Declare Function iblockx32 Lib "Gpib-32.dll" Alias "iblockxA" (ByVal ud As Long, ByVal LockWaitTime As Long, arg1 As Any) As Long
Declare Function ibunlockx32 Lib "Gpib-32.dll" Alias "ibunlockx" (ByVal ud As Long) As Long


' <VB WATCH>
Const VBWMODULE = "VBIB32"
' </VB WATCH>

Sub AllSpoll(ByVal ud As Integer, addrs() As Integer, results() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "VBIB32.AllSpoll"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
7                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
8                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("results", results) & ") "
9              End If
10             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
11         End If
' </VB WATCH>
12         If (GPIBglobalsRegistered = 0) Then
13           Call RegisterGPIBGlobals
14         End If

       ' Call the 32-bit DLL.
15         Call AllSpoll32(ud, addrs(0), results(0))

16         Call copy_ibvars
' <VB WATCH>
17         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
18         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AllSpoll"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportVariable "results", results
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub copy_ibvars()
' <VB WATCH>
19         On Error GoTo vbwErrHandler
20         Const VBWPROCNAME = "VBIB32.copy_ibvars"
21         If vbwProtector.vbwTraceProc Then
22             Dim vbwProtectorParameterString As String
23             If vbwProtector.vbwTraceParameters Then
24                 vbwProtectorParameterString = "()"
25             End If
26             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
27         End If
' </VB WATCH>
28         ibsta = ConvertLongToInt(Longibsta)
29         iberr = CInt(Longiberr)
30         ibcnt = ConvertLongToInt(ibcntl)
' <VB WATCH>
31         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
32         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "copy_ibvars"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub DevClear(ByVal ud As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
33         On Error GoTo vbwErrHandler
34         Const VBWPROCNAME = "VBIB32.DevClear"
35         If vbwProtector.vbwTraceProc Then
36             Dim vbwProtectorParameterString As String
37             If vbwProtector.vbwTraceParameters Then
38                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
39                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
40             End If
41             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
42         End If
' </VB WATCH>
43         If (GPIBglobalsRegistered = 0) Then
44           Call RegisterGPIBGlobals
45         End If

       ' Call the 32-bit DLL.
46         Call DevClear32(ud, addr)

47         Call copy_ibvars
' <VB WATCH>
48         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
49         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DevClear"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub DevClearList(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
50         On Error GoTo vbwErrHandler
51         Const VBWPROCNAME = "VBIB32.DevClearList"
52         If vbwProtector.vbwTraceProc Then
53             Dim vbwProtectorParameterString As String
54             If vbwProtector.vbwTraceParameters Then
55                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
56                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
57             End If
58             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
59         End If
' </VB WATCH>
60         If (GPIBglobalsRegistered = 0) Then
61           Call RegisterGPIBGlobals
62         End If

       ' Call the 32-bit DLL.
63         Call DevClearList32(ud, addrs(0))

64         Call copy_ibvars
' <VB WATCH>
65         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
66         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DevClearList"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub EnableLocal(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
67         On Error GoTo vbwErrHandler
68         Const VBWPROCNAME = "VBIB32.EnableLocal"
69         If vbwProtector.vbwTraceProc Then
70             Dim vbwProtectorParameterString As String
71             If vbwProtector.vbwTraceParameters Then
72                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
73                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
74             End If
75             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
76         End If
' </VB WATCH>
77         If (GPIBglobalsRegistered = 0) Then
78           Call RegisterGPIBGlobals
79         End If

       ' Call the 32-bit DLL.
80         Call EnableLocal32(ud, addrs(0))

81         Call copy_ibvars
' <VB WATCH>
82         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
83         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableLocal"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub EnableRemote(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
84         On Error GoTo vbwErrHandler
85         Const VBWPROCNAME = "VBIB32.EnableRemote"
86         If vbwProtector.vbwTraceProc Then
87             Dim vbwProtectorParameterString As String
88             If vbwProtector.vbwTraceParameters Then
89                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
90                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
91             End If
92             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
93         End If
' </VB WATCH>
94         If (GPIBglobalsRegistered = 0) Then
95           Call RegisterGPIBGlobals
96         End If

       ' Call the 32-bit DLL.
97         Call EnableRemote32(ud, addrs(0))

98         Call copy_ibvars
' <VB WATCH>
99         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
100        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableRemote"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub FindLstn(ByVal ud As Integer, addrs() As Integer, results() As Integer, ByVal limit As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
101        On Error GoTo vbwErrHandler
102        Const VBWPROCNAME = "VBIB32.FindLstn"
103        If vbwProtector.vbwTraceProc Then
104            Dim vbwProtectorParameterString As String
105            If vbwProtector.vbwTraceParameters Then
106                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
107                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
108                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("results", results) & ", "
109                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("limit", limit) & ") "
110            End If
111            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
112        End If
' </VB WATCH>
113        If (GPIBglobalsRegistered = 0) Then
114          Call RegisterGPIBGlobals
115        End If

       ' Call the 32-bit DLL.
116        Call FindLstn32(ud, addrs(0), results(0), limit)

117        Call copy_ibvars
' <VB WATCH>
118        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
119        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindLstn"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportVariable "results", results
            vbwReportVariable "limit", limit
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub FindRQS(ByVal ud As Integer, addrs() As Integer, result As Integer)
' <VB WATCH>
120        On Error GoTo vbwErrHandler
121        Const VBWPROCNAME = "VBIB32.FindRQS"
122        If vbwProtector.vbwTraceProc Then
123            Dim vbwProtectorParameterString As String
124            If vbwProtector.vbwTraceParameters Then
125                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
126                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
127                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
128            End If
129            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
130        End If
' </VB WATCH>
131       Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
132        If (GPIBglobalsRegistered = 0) Then
133          Call RegisterGPIBGlobals
134        End If

       ' Call the 32-bit DLL.
135        Call FindRQS32(ud, addrs(0), tmpresult)

136        result = ConvertLongToInt(tmpresult)

137        Call copy_ibvars
' <VB WATCH>
138        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
139        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindRQS"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportVariable "result", result
            vbwReportVariable "tmpresult", tmpresult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer)
' <VB WATCH>
140        On Error GoTo vbwErrHandler
141        Const VBWPROCNAME = "VBIB32.ibask"
142        If vbwProtector.vbwTraceProc Then
143            Dim vbwProtectorParameterString As String
144            If vbwProtector.vbwTraceParameters Then
145                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
146                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
147                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rval", rval) & ") "
148            End If
149            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
150        End If
' </VB WATCH>
151      Dim tmprval As Long

       ' Check to see if GPIB Global variables are registered
152        If (GPIBglobalsRegistered = 0) Then
153          Call RegisterGPIBGlobals
154        End If

       ' Call the 32-bit DLL.
155        Call ibask32(ud, opt, tmprval)

156        rval = ConvertLongToInt(tmprval)

157        Call copy_ibvars
' <VB WATCH>
158        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
159        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibask"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "opt", opt
            vbwReportVariable "rval", rval
            vbwReportVariable "tmprval", tmprval
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibbna(ByVal ud As Integer, ByVal udname As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
160        On Error GoTo vbwErrHandler
161        Const VBWPROCNAME = "VBIB32.ibbna"
162        If vbwProtector.vbwTraceProc Then
163            Dim vbwProtectorParameterString As String
164            If vbwProtector.vbwTraceParameters Then
165                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
166                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("udname", udname) & ") "
167            End If
168            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
169        End If
' </VB WATCH>
170        If (GPIBglobalsRegistered = 0) Then
171          Call RegisterGPIBGlobals
172        End If

       ' Call the 32-bit DLL.
173        Call ibbna32(ud, ByVal udname)

174        Call copy_ibvars
' <VB WATCH>
175        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
176        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibbna"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "udname", udname
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibcac(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
177        On Error GoTo vbwErrHandler
178        Const VBWPROCNAME = "VBIB32.ibcac"
179        If vbwProtector.vbwTraceProc Then
180            Dim vbwProtectorParameterString As String
181            If vbwProtector.vbwTraceParameters Then
182                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
183                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
184            End If
185            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
186        End If
' </VB WATCH>
187        If (GPIBglobalsRegistered = 0) Then
188          Call RegisterGPIBGlobals
189        End If

       ' Call the 32-bit DLL.
190        Call ibcac32(ud, v)

191        Call copy_ibvars
' <VB WATCH>
192        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
193        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibcac"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibclr(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
194        On Error GoTo vbwErrHandler
195        Const VBWPROCNAME = "VBIB32.ibclr"
196        If vbwProtector.vbwTraceProc Then
197            Dim vbwProtectorParameterString As String
198            If vbwProtector.vbwTraceParameters Then
199                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
200            End If
201            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
202        End If
' </VB WATCH>
203        If (GPIBglobalsRegistered = 0) Then
204          Call RegisterGPIBGlobals
205        End If

       ' Call the 32-bit DLL.
206        Call ibclr32(ud)

207        Call copy_ibvars
' <VB WATCH>
208        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
209        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibclr"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibcmd(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
210        On Error GoTo vbwErrHandler
211        Const VBWPROCNAME = "VBIB32.ibcmd"
212        If vbwProtector.vbwTraceProc Then
213            Dim vbwProtectorParameterString As String
214            If vbwProtector.vbwTraceParameters Then
215                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
216                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
217            End If
218            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
219        End If
' </VB WATCH>
220       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
221        If (GPIBglobalsRegistered = 0) Then
222          Call RegisterGPIBGlobals
223        End If

224        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
225        Call ibcmd32(ud, ByVal buf, cnt)

226        Call copy_ibvars
' <VB WATCH>
227        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
228        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibcmd"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibcmda(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
229        On Error GoTo vbwErrHandler
230        Const VBWPROCNAME = "VBIB32.ibcmda"
231        If vbwProtector.vbwTraceProc Then
232            Dim vbwProtectorParameterString As String
233            If vbwProtector.vbwTraceParameters Then
234                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
235                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
236            End If
237            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
238        End If
' </VB WATCH>
239        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
240        If (GPIBglobalsRegistered = 0) Then
241          Call RegisterGPIBGlobals
242        End If

243        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
244        Call ibcmd32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem solved, then use:
       '    call ibcmda32(ud, ByVal buf, cnt)

245        Call copy_ibvars
' <VB WATCH>
246        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
247        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibcmda"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
248        On Error GoTo vbwErrHandler
249        Const VBWPROCNAME = "VBIB32.ibconfig"
250        If vbwProtector.vbwTraceProc Then
251            Dim vbwProtectorParameterString As String
252            If vbwProtector.vbwTraceParameters Then
253                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
254                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
255                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
256            End If
257            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
258        End If
' </VB WATCH>
259        If (GPIBglobalsRegistered = 0) Then
260          Call RegisterGPIBGlobals
261        End If

       ' Call the 32-bit DLL.
262        Call ibconfig32(bdid, opt, v)

263        Call copy_ibvars
' <VB WATCH>
264        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
265        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibconfig"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "opt", opt
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibdev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer, ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
266        On Error GoTo vbwErrHandler
267        Const VBWPROCNAME = "VBIB32.ibdev"
268        If vbwProtector.vbwTraceProc Then
269            Dim vbwProtectorParameterString As String
270            If vbwProtector.vbwTraceParameters Then
271                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
272                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
273                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
274                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("tmo", tmo) & ", "
275                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eot", eot) & ", "
276                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eos", eos) & ", "
277                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ud", ud) & ") "
278            End If
279            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
280        End If
' </VB WATCH>
281        If (GPIBglobalsRegistered = 0) Then
282          Call RegisterGPIBGlobals
283        End If

       ' Call the 32-bit DLL.
284        ud = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

285        Call copy_ibvars
' <VB WATCH>
286        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
287        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibdev"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "tmo", tmo
            vbwReportVariable "eot", eot
            vbwReportVariable "eos", eos
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Sub ibdma(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
288        On Error GoTo vbwErrHandler
289        Const VBWPROCNAME = "VBIB32.ibdma"
290        If vbwProtector.vbwTraceProc Then
291            Dim vbwProtectorParameterString As String
292            If vbwProtector.vbwTraceParameters Then
293                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
294                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
295            End If
296            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
297        End If
' </VB WATCH>
298        If (GPIBglobalsRegistered = 0) Then
299          Call RegisterGPIBGlobals
300        End If

       ' Call the 32-bit DLL.
301        Call ibdma32(ud, v)

302        Call copy_ibvars
' <VB WATCH>
303        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
304        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibdma"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibeos(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
305        On Error GoTo vbwErrHandler
306        Const VBWPROCNAME = "VBIB32.ibeos"
307        If vbwProtector.vbwTraceProc Then
308            Dim vbwProtectorParameterString As String
309            If vbwProtector.vbwTraceParameters Then
310                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
311                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
312            End If
313            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
314        End If
' </VB WATCH>
315        If (GPIBglobalsRegistered = 0) Then
316          Call RegisterGPIBGlobals
317        End If

       ' Call the 32-bit DLL.
318        Call ibeos32(ud, v)

319        Call copy_ibvars
' <VB WATCH>
320        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
321        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibeos"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibeot(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
322        On Error GoTo vbwErrHandler
323        Const VBWPROCNAME = "VBIB32.ibeot"
324        If vbwProtector.vbwTraceProc Then
325            Dim vbwProtectorParameterString As String
326            If vbwProtector.vbwTraceParameters Then
327                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
328                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
329            End If
330            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
331        End If
' </VB WATCH>
332        If (GPIBglobalsRegistered = 0) Then
333          Call RegisterGPIBGlobals
334        End If

       ' Call the 32-bit DLL.
335        Call ibeot32(ud, v)

336        Call copy_ibvars
' <VB WATCH>
337        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
338        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibeot"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Sub ibfind(ByVal udname As String, ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
339        On Error GoTo vbwErrHandler
340        Const VBWPROCNAME = "VBIB32.ibfind"
341        If vbwProtector.vbwTraceProc Then
342            Dim vbwProtectorParameterString As String
343            If vbwProtector.vbwTraceParameters Then
344                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("udname", udname) & ", "
345                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ud", ud) & ") "
346            End If
347            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
348        End If
' </VB WATCH>
349        If (GPIBglobalsRegistered = 0) Then
350          Call RegisterGPIBGlobals
351        End If

       ' Call the 32-bit DLL.
352        ud = ConvertLongToInt(ibfind32(ByVal udname))

353        Call copy_ibvars
' <VB WATCH>
354        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
355        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibfind"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "udname", udname
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibgts(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
356        On Error GoTo vbwErrHandler
357        Const VBWPROCNAME = "VBIB32.ibgts"
358        If vbwProtector.vbwTraceProc Then
359            Dim vbwProtectorParameterString As String
360            If vbwProtector.vbwTraceParameters Then
361                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
362                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
363            End If
364            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
365        End If
' </VB WATCH>
366        If (GPIBglobalsRegistered = 0) Then
367          Call RegisterGPIBGlobals
368        End If

       ' Call the 32-bit DLL.
369        Call ibgts32(ud, v)

370        Call copy_ibvars
' <VB WATCH>
371        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
372        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibgts"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibist(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
373        On Error GoTo vbwErrHandler
374        Const VBWPROCNAME = "VBIB32.ibist"
375        If vbwProtector.vbwTraceProc Then
376            Dim vbwProtectorParameterString As String
377            If vbwProtector.vbwTraceParameters Then
378                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
379                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
380            End If
381            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
382        End If
' </VB WATCH>
383        If (GPIBglobalsRegistered = 0) Then
384          Call RegisterGPIBGlobals
385        End If

       ' Call the 32-bit DLL.
386        Call ibist32(ud, v)

387        Call copy_ibvars
' <VB WATCH>
388        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
389        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibist"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub iblines(ByVal ud As Integer, lines As Integer)
' <VB WATCH>
390        On Error GoTo vbwErrHandler
391        Const VBWPROCNAME = "VBIB32.iblines"
392        If vbwProtector.vbwTraceProc Then
393            Dim vbwProtectorParameterString As String
394            If vbwProtector.vbwTraceParameters Then
395                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
396                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("lines", lines) & ") "
397            End If
398            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
399        End If
' </VB WATCH>
400       Dim tmplines As Long

       ' Check to see if GPIB Global variables are registered
401        If (GPIBglobalsRegistered = 0) Then
402          Call RegisterGPIBGlobals
403        End If

       ' Call the 32-bit DLL.
404        Call iblines32(ud, tmplines)

405        lines = ConvertLongToInt(tmplines)

406        Call copy_ibvars
' <VB WATCH>
407        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
408        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iblines"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "lines", lines
            vbwReportVariable "tmplines", tmplines
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibln(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer)
' <VB WATCH>
409        On Error GoTo vbwErrHandler
410        Const VBWPROCNAME = "VBIB32.ibln"
411        If vbwProtector.vbwTraceProc Then
412            Dim vbwProtectorParameterString As String
413            If vbwProtector.vbwTraceParameters Then
414                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
415                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
416                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
417                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ln", ln) & ") "
418            End If
419            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
420        End If
' </VB WATCH>
421        Dim tmpln As Long

       ' Check to see if GPIB Global variables are registered
422        If (GPIBglobalsRegistered = 0) Then
423          Call RegisterGPIBGlobals
424        End If

       ' Call the 32-bit DLL.
425        Call ibln32(ud, pad, sad, tmpln)

426        ln = ConvertLongToInt(tmpln)

427        Call copy_ibvars
' <VB WATCH>
428        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
429        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibln"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "ln", ln
            vbwReportVariable "tmpln", tmpln
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibloc(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
430        On Error GoTo vbwErrHandler
431        Const VBWPROCNAME = "VBIB32.ibloc"
432        If vbwProtector.vbwTraceProc Then
433            Dim vbwProtectorParameterString As String
434            If vbwProtector.vbwTraceParameters Then
435                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
436            End If
437            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
438        End If
' </VB WATCH>
439        If (GPIBglobalsRegistered = 0) Then
440          Call RegisterGPIBGlobals
441        End If

       ' Call the 32-bit DLL.
442        Call ibloc32(ud)

443        Call copy_ibvars
' <VB WATCH>
444        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
445        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibloc"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibonl(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
446        On Error GoTo vbwErrHandler
447        Const VBWPROCNAME = "VBIB32.ibonl"
448        If vbwProtector.vbwTraceProc Then
449            Dim vbwProtectorParameterString As String
450            If vbwProtector.vbwTraceParameters Then
451                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
452                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
453            End If
454            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
455        End If
' </VB WATCH>
456        If (GPIBglobalsRegistered = 0) Then
457          Call RegisterGPIBGlobals
458        End If

       ' Call the 32-bit DLL.
459        Call ibonl32(ud, v)

460        Call copy_ibvars
' <VB WATCH>
461        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
462        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibonl"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibpad(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
463        On Error GoTo vbwErrHandler
464        Const VBWPROCNAME = "VBIB32.ibpad"
465        If vbwProtector.vbwTraceProc Then
466            Dim vbwProtectorParameterString As String
467            If vbwProtector.vbwTraceParameters Then
468                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
469                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
470            End If
471            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
472        End If
' </VB WATCH>
473        If (GPIBglobalsRegistered = 0) Then
474          Call RegisterGPIBGlobals
475        End If

       ' Call the 32-bit DLL.
476        Call ibpad32(ud, v)

477        Call copy_ibvars
' <VB WATCH>
478        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
479        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibpad"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibpct(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
480        On Error GoTo vbwErrHandler
481        Const VBWPROCNAME = "VBIB32.ibpct"
482        If vbwProtector.vbwTraceProc Then
483            Dim vbwProtectorParameterString As String
484            If vbwProtector.vbwTraceParameters Then
485                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
486            End If
487            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
488        End If
' </VB WATCH>
489        If (GPIBglobalsRegistered = 0) Then
490          Call RegisterGPIBGlobals
491        End If

       ' Call the 32-bit DLL.
492        Call ibpct32(ud)

493        Call copy_ibvars
' <VB WATCH>
494        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
495        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibpct"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Sub ibppc(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
496        On Error GoTo vbwErrHandler
497        Const VBWPROCNAME = "VBIB32.ibppc"
498        If vbwProtector.vbwTraceProc Then
499            Dim vbwProtectorParameterString As String
500            If vbwProtector.vbwTraceParameters Then
501                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
502                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
503            End If
504            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
505        End If
' </VB WATCH>
506        If (GPIBglobalsRegistered = 0) Then
507          Call RegisterGPIBGlobals
508        End If

       ' Call the 32-bit DLL.
509        Call ibppc32(ud, v)

510        Call copy_ibvars
' <VB WATCH>
511        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
512        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibppc"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrd(ByVal ud As Integer, buf As String)
' <VB WATCH>
513        On Error GoTo vbwErrHandler
514        Const VBWPROCNAME = "VBIB32.ibrd"
515        If vbwProtector.vbwTraceProc Then
516            Dim vbwProtectorParameterString As String
517            If vbwProtector.vbwTraceParameters Then
518                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
519                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
520            End If
521            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
522        End If
' </VB WATCH>
523        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
524        If (GPIBglobalsRegistered = 0) Then
525          Call RegisterGPIBGlobals
526        End If

527        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
528        Call ibrd32(ud, ByVal buf, cnt)

529        Call copy_ibvars
' <VB WATCH>
530        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
531        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrd"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrda(ByVal ud As Integer, buf As String)
' <VB WATCH>
532        On Error GoTo vbwErrHandler
533        Const VBWPROCNAME = "VBIB32.ibrda"
534        If vbwProtector.vbwTraceProc Then
535            Dim vbwProtectorParameterString As String
536            If vbwProtector.vbwTraceParameters Then
537                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
538                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
539            End If
540            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
541        End If
' </VB WATCH>
542        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
543        If (GPIBglobalsRegistered = 0) Then
544          Call RegisterGPIBGlobals
545        End If

546        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
547        Call ibrd32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem solved, use this:
       '    Call ibrda32(ud, ByVal buf, cnt)

548        Call copy_ibvars
' <VB WATCH>
549        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
550        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrda"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrdf(ByVal ud As Integer, ByVal filename As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
551        On Error GoTo vbwErrHandler
552        Const VBWPROCNAME = "VBIB32.ibrdf"
553        If vbwProtector.vbwTraceProc Then
554            Dim vbwProtectorParameterString As String
555            If vbwProtector.vbwTraceParameters Then
556                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
557                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
558            End If
559            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
560        End If
' </VB WATCH>
561        If (GPIBglobalsRegistered = 0) Then
562          Call RegisterGPIBGlobals
563        End If

       ' Call the 32-bit DLL.
564        Call ibrdf32(ud, ByVal filename)

565        Call copy_ibvars
' <VB WATCH>
566        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
567        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrdf"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
568        On Error GoTo vbwErrHandler
569        Const VBWPROCNAME = "VBIB32.ibrdi"
570        If vbwProtector.vbwTraceProc Then
571            Dim vbwProtectorParameterString As String
572            If vbwProtector.vbwTraceParameters Then
573                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
574                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
575                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
576            End If
577            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
578        End If
' </VB WATCH>
579        If (GPIBglobalsRegistered = 0) Then
580          Call RegisterGPIBGlobals
581        End If

       ' Call the 32-bit DLL.
582        Call ibrd32(ud, ibuf(0), cnt)

583        Call copy_ibvars
' <VB WATCH>
584        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
585        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrdi"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
586        On Error GoTo vbwErrHandler
587        Const VBWPROCNAME = "VBIB32.ibrdia"
588        If vbwProtector.vbwTraceProc Then
589            Dim vbwProtectorParameterString As String
590            If vbwProtector.vbwTraceParameters Then
591                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
592                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
593                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
594            End If
595            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
596        End If
' </VB WATCH>
597        If (GPIBglobalsRegistered = 0) Then
598          Call RegisterGPIBGlobals
599        End If

       ' Call the 32-bit DLL.
600        Call ibrd32(ud, ibuf(0), cnt)

       ' When Visual Basic remapping buffer problem is solved, then use:
       '    Call ibrda32(u, ibuf(0), cnt)

601        Call copy_ibvars
' <VB WATCH>
602        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
603        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrdia"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Sub ibrpp(ByVal ud As Integer, ppr As Integer)
' <VB WATCH>
604        On Error GoTo vbwErrHandler
605        Const VBWPROCNAME = "VBIB32.ibrpp"
606        If vbwProtector.vbwTraceProc Then
607            Dim vbwProtectorParameterString As String
608            If vbwProtector.vbwTraceParameters Then
609                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
610                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ppr", ppr) & ") "
611            End If
612            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
613        End If
' </VB WATCH>
614        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
615        If (GPIBglobalsRegistered = 0) Then
616          Call RegisterGPIBGlobals
617        End If

       ' Call the 32-bit DLL.
618        Call ibrpp32(ud, ByVal tmp_str)

619        ppr = Asc(tmp_str)

620        Call copy_ibvars
' <VB WATCH>
621        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
622        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrpp"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ppr", ppr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrsc(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
623        On Error GoTo vbwErrHandler
624        Const VBWPROCNAME = "VBIB32.ibrsc"
625        If vbwProtector.vbwTraceProc Then
626            Dim vbwProtectorParameterString As String
627            If vbwProtector.vbwTraceParameters Then
628                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
629                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
630            End If
631            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
632        End If
' </VB WATCH>
633        If (GPIBglobalsRegistered = 0) Then
634          Call RegisterGPIBGlobals
635        End If

       ' Call the 32-bit DLL.
636        Call ibrsc32(ud, v)

637        Call copy_ibvars
' <VB WATCH>
638        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
639        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrsc"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrsp(ByVal ud As Integer, spr As Integer)
' <VB WATCH>
640        On Error GoTo vbwErrHandler
641        Const VBWPROCNAME = "VBIB32.ibrsp"
642        If vbwProtector.vbwTraceProc Then
643            Dim vbwProtectorParameterString As String
644            If vbwProtector.vbwTraceParameters Then
645                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
646                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("spr", spr) & ") "
647            End If
648            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
649        End If
' </VB WATCH>
650        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
651        If (GPIBglobalsRegistered = 0) Then
652          Call RegisterGPIBGlobals
653        End If

       ' Call the 32-bit DLL
654        Call ibrsp32(ud, ByVal tmp_str)

655        spr = Asc(tmp_str)

656        Call copy_ibvars
' <VB WATCH>
657        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
658        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrsp"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "spr", spr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrsv(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
659        On Error GoTo vbwErrHandler
660        Const VBWPROCNAME = "VBIB32.ibrsv"
661        If vbwProtector.vbwTraceProc Then
662            Dim vbwProtectorParameterString As String
663            If vbwProtector.vbwTraceParameters Then
664                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
665                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
666            End If
667            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
668        End If
' </VB WATCH>
669        If (GPIBglobalsRegistered = 0) Then
670          Call RegisterGPIBGlobals
671        End If

       ' Call the 32-bit DLL.
672        Call ibrsv32(ud, v)

673        Call copy_ibvars
' <VB WATCH>
674        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
675        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrsv"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibsad(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
676        On Error GoTo vbwErrHandler
677        Const VBWPROCNAME = "VBIB32.ibsad"
678        If vbwProtector.vbwTraceProc Then
679            Dim vbwProtectorParameterString As String
680            If vbwProtector.vbwTraceParameters Then
681                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
682                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
683            End If
684            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
685        End If
' </VB WATCH>
686        If (GPIBglobalsRegistered = 0) Then
687          Call RegisterGPIBGlobals
688        End If

       ' Call the 32-bit DLL.
689        Call ibsad32(ud, v)

690        Call copy_ibvars
' <VB WATCH>
691        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
692        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibsad"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibsic(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
693        On Error GoTo vbwErrHandler
694        Const VBWPROCNAME = "VBIB32.ibsic"
695        If vbwProtector.vbwTraceProc Then
696            Dim vbwProtectorParameterString As String
697            If vbwProtector.vbwTraceParameters Then
698                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
699            End If
700            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
701        End If
' </VB WATCH>
702        If (GPIBglobalsRegistered = 0) Then
703          Call RegisterGPIBGlobals
704        End If

       ' Call the 32-bit DLL.
705        Call ibsic32(ud)

706        Call copy_ibvars
' <VB WATCH>
707        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
708        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibsic"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibsre(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
709        On Error GoTo vbwErrHandler
710        Const VBWPROCNAME = "VBIB32.ibsre"
711        If vbwProtector.vbwTraceProc Then
712            Dim vbwProtectorParameterString As String
713            If vbwProtector.vbwTraceParameters Then
714                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
715                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
716            End If
717            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
718        End If
' </VB WATCH>
719        If (GPIBglobalsRegistered = 0) Then
720          Call RegisterGPIBGlobals
721        End If

       ' Call the 32-bit DLL.
722        Call ibsre32(ud, v)

723        Call copy_ibvars
' <VB WATCH>
724        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
725        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibsre"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibstop(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
726        On Error GoTo vbwErrHandler
727        Const VBWPROCNAME = "VBIB32.ibstop"
728        If vbwProtector.vbwTraceProc Then
729            Dim vbwProtectorParameterString As String
730            If vbwProtector.vbwTraceParameters Then
731                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
732            End If
733            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
734        End If
' </VB WATCH>
735        If (GPIBglobalsRegistered = 0) Then
736          Call RegisterGPIBGlobals
737        End If

       ' Call the 32-bit DLL.
738        Call ibstop32(ud)

739        Call copy_ibvars
' <VB WATCH>
740        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
741        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibstop"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibtmo(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
742        On Error GoTo vbwErrHandler
743        Const VBWPROCNAME = "VBIB32.ibtmo"
744        If vbwProtector.vbwTraceProc Then
745            Dim vbwProtectorParameterString As String
746            If vbwProtector.vbwTraceParameters Then
747                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
748                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
749            End If
750            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
751        End If
' </VB WATCH>
752        If (GPIBglobalsRegistered = 0) Then
753          Call RegisterGPIBGlobals
754        End If

       ' Call the 32-bit DLL.
755        Call ibtmo32(ud, v)

756        Call copy_ibvars
' <VB WATCH>
757        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
758        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibtmo"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibtrg(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
759        On Error GoTo vbwErrHandler
760        Const VBWPROCNAME = "VBIB32.ibtrg"
761        If vbwProtector.vbwTraceProc Then
762            Dim vbwProtectorParameterString As String
763            If vbwProtector.vbwTraceParameters Then
764                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
765            End If
766            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
767        End If
' </VB WATCH>
768        If (GPIBglobalsRegistered = 0) Then
769          Call RegisterGPIBGlobals
770        End If

       ' Call 32-bit DLL.
771        Call ibtrg32(ud)

772        Call copy_ibvars
' <VB WATCH>
773        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
774        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibtrg"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwait(ByVal ud As Integer, ByVal mask As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
775        On Error GoTo vbwErrHandler
776        Const VBWPROCNAME = "VBIB32.ibwait"
777        If vbwProtector.vbwTraceProc Then
778            Dim vbwProtectorParameterString As String
779            If vbwProtector.vbwTraceParameters Then
780                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
781                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("mask", mask) & ") "
782            End If
783            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
784        End If
' </VB WATCH>
785        If (GPIBglobalsRegistered = 0) Then
786          Call RegisterGPIBGlobals
787        End If

       ' Call the 32-bit DLL.
788        Call ibwait32(ud, mask)

789        Call copy_ibvars
' <VB WATCH>
790        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
791        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwait"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "mask", mask
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrt(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
792        On Error GoTo vbwErrHandler
793        Const VBWPROCNAME = "VBIB32.ibwrt"
794        If vbwProtector.vbwTraceProc Then
795            Dim vbwProtectorParameterString As String
796            If vbwProtector.vbwTraceParameters Then
797                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
798                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
799            End If
800            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
801        End If
' </VB WATCH>
802        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
803        If (GPIBglobalsRegistered = 0) Then
804          Call RegisterGPIBGlobals
805        End If

806        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
807        Call ibwrt32(ud, ByVal buf, cnt)

808        Call copy_ibvars
' <VB WATCH>
809        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
810        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrt"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrta(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
811        On Error GoTo vbwErrHandler
812        Const VBWPROCNAME = "VBIB32.ibwrta"
813        If vbwProtector.vbwTraceProc Then
814            Dim vbwProtectorParameterString As String
815            If vbwProtector.vbwTraceParameters Then
816                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
817                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
818            End If
819            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
820        End If
' </VB WATCH>
821        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
822        If (GPIBglobalsRegistered = 0) Then
823          Call RegisterGPIBGlobals
824        End If

825        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
826        Call ibwrt32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    Call ibwrta32(ud, ByVal buf, cnt)

827        Call copy_ibvars
' <VB WATCH>
828        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
829        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrta"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrtf(ByVal ud As Integer, ByVal filename As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
830        On Error GoTo vbwErrHandler
831        Const VBWPROCNAME = "VBIB32.ibwrtf"
832        If vbwProtector.vbwTraceProc Then
833            Dim vbwProtectorParameterString As String
834            If vbwProtector.vbwTraceParameters Then
835                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
836                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
837            End If
838            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
839        End If
' </VB WATCH>
840        If (GPIBglobalsRegistered = 0) Then
841          Call RegisterGPIBGlobals
842        End If

       ' Call the 32-bit DLL.
843        Call ibwrtf32(ud, ByVal filename)

844        Call copy_ibvars
' <VB WATCH>
845        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
846        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrtf"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
847        On Error GoTo vbwErrHandler
848        Const VBWPROCNAME = "VBIB32.ibwrti"
849        If vbwProtector.vbwTraceProc Then
850            Dim vbwProtectorParameterString As String
851            If vbwProtector.vbwTraceParameters Then
852                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
853                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
854                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
855            End If
856            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
857        End If
' </VB WATCH>
858        If (GPIBglobalsRegistered = 0) Then
859          Call RegisterGPIBGlobals
860        End If

       ' Call the 32-bit DLL.
861        Call ibwrt32(ud, ibuf(0), cnt)

862        Call copy_ibvars
' <VB WATCH>
863        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
864        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrti"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
865        On Error GoTo vbwErrHandler
866        Const VBWPROCNAME = "VBIB32.ibwrtia"
867        If vbwProtector.vbwTraceProc Then
868            Dim vbwProtectorParameterString As String
869            If vbwProtector.vbwTraceParameters Then
870                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
871                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
872                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
873            End If
874            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
875        End If
' </VB WATCH>
876        If (GPIBglobalsRegistered = 0) Then
877          Call RegisterGPIBGlobals
878        End If

       ' Call the 32-bit DLL.
879        Call ibwrt32(ud, ibuf(0), cnt)

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    Call ibwrta32(ud, ibuf(0), cnt)

880        Call copy_ibvars
' <VB WATCH>
881        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
882        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrtia"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Function ilask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer) As Integer
' <VB WATCH>
883        On Error GoTo vbwErrHandler
884        Const VBWPROCNAME = "VBIB32.ilask"
885        If vbwProtector.vbwTraceProc Then
886            Dim vbwProtectorParameterString As String
887            If vbwProtector.vbwTraceParameters Then
888                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
889                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
890                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rval", rval) & ") "
891            End If
892            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
893        End If
' </VB WATCH>
894        Dim tmprval As Long

       ' Check to see if GPIB Global variables are registered
895        If (GPIBglobalsRegistered = 0) Then
896          Call RegisterGPIBGlobals
897        End If

       ' Call the 32-bit DLL.
898        ilask = ConvertLongToInt(ibask32(ud, opt, tmprval))

899        rval = ConvertLongToInt(tmprval)

900        Call copy_ibvars
' <VB WATCH>
901        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
902        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilask"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "opt", opt
            vbwReportVariable "rval", rval
            vbwReportVariable "tmprval", tmprval
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilbna(ByVal ud As Integer, ByVal udname As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
903        On Error GoTo vbwErrHandler
904        Const VBWPROCNAME = "VBIB32.ilbna"
905        If vbwProtector.vbwTraceProc Then
906            Dim vbwProtectorParameterString As String
907            If vbwProtector.vbwTraceParameters Then
908                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
909                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("udname", udname) & ") "
910            End If
911            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
912        End If
' </VB WATCH>
913        If (GPIBglobalsRegistered = 0) Then
914          Call RegisterGPIBGlobals
915        End If

       ' Call the 32-bit DLL.
916        ilbna = ConvertLongToInt(ibbna32(ud, ByVal udname))

917        Call copy_ibvars
' <VB WATCH>
918        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
919        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilbna"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "udname", udname
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilcac(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
920        On Error GoTo vbwErrHandler
921        Const VBWPROCNAME = "VBIB32.ilcac"
922        If vbwProtector.vbwTraceProc Then
923            Dim vbwProtectorParameterString As String
924            If vbwProtector.vbwTraceParameters Then
925                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
926                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
927            End If
928            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
929        End If
' </VB WATCH>
930        If (GPIBglobalsRegistered = 0) Then
931          Call RegisterGPIBGlobals
932        End If

       ' Call the 32-bit DLL.
933        ilcac = ConvertLongToInt(ibcac32(ud, v))

934        Call copy_ibvars
' <VB WATCH>
935        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
936        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilcac"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilclr(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
937        On Error GoTo vbwErrHandler
938        Const VBWPROCNAME = "VBIB32.ilclr"
939        If vbwProtector.vbwTraceProc Then
940            Dim vbwProtectorParameterString As String
941            If vbwProtector.vbwTraceParameters Then
942                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
943            End If
944            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
945        End If
' </VB WATCH>
946        If (GPIBglobalsRegistered = 0) Then
947          Call RegisterGPIBGlobals
948        End If

       ' Call the 32-bit DLL.
949        ilclr = ConvertLongToInt(ibclr32(ud))

950        Call copy_ibvars
' <VB WATCH>
951        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
952        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilclr"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilcmd(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
953        On Error GoTo vbwErrHandler
954        Const VBWPROCNAME = "VBIB32.ilcmd"
955        If vbwProtector.vbwTraceProc Then
956            Dim vbwProtectorParameterString As String
957            If vbwProtector.vbwTraceParameters Then
958                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
959                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
960                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
961            End If
962            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
963        End If
' </VB WATCH>
964        If (GPIBglobalsRegistered = 0) Then
965          Call RegisterGPIBGlobals
966        End If

       ' Call the 32-bit DLL.
967        ilcmd = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))

968        Call copy_ibvars
' <VB WATCH>
969        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
970        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilcmd"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilcmda(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
971        On Error GoTo vbwErrHandler
972        Const VBWPROCNAME = "VBIB32.ilcmda"
973        If vbwProtector.vbwTraceProc Then
974            Dim vbwProtectorParameterString As String
975            If vbwProtector.vbwTraceParameters Then
976                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
977                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
978                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
979            End If
980            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
981        End If
' </VB WATCH>
982        If (GPIBglobalsRegistered = 0) Then
983          Call RegisterGPIBGlobals
984        End If

       ' Call the 32-bit DLL.
985        ilcmda = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    ilcmda = ConvertLongToInt(ibcmda32(ud, ByVal buf, cnt))

986        Call copy_ibvars
' <VB WATCH>
987        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
988        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilcmda"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
989        On Error GoTo vbwErrHandler
990        Const VBWPROCNAME = "VBIB32.ilconfig"
991        If vbwProtector.vbwTraceProc Then
992            Dim vbwProtectorParameterString As String
993            If vbwProtector.vbwTraceParameters Then
994                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
995                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
996                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
997            End If
998            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
999        End If
' </VB WATCH>
1000       If (GPIBglobalsRegistered = 0) Then
1001         Call RegisterGPIBGlobals
1002       End If

       ' Call the 32-bit DLL.
1003       ilconfig = ConvertLongToInt(ibconfig32(bdid, opt, v))

1004       Call copy_ibvars
' <VB WATCH>
1005       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1006       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilconfig"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "opt", opt
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ildev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1007       On Error GoTo vbwErrHandler
1008       Const VBWPROCNAME = "VBIB32.ildev"
1009       If vbwProtector.vbwTraceProc Then
1010           Dim vbwProtectorParameterString As String
1011           If vbwProtector.vbwTraceParameters Then
1012               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
1013               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
1014               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
1015               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("tmo", tmo) & ", "
1016               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eot", eot) & ", "
1017               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eos", eos) & ") "
1018           End If
1019           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1020       End If
' </VB WATCH>
1021       If (GPIBglobalsRegistered = 0) Then
1022         Call RegisterGPIBGlobals
1023       End If

       ' Call the 32-bit DLL.
1024       ildev = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

1025       Call copy_ibvars
' <VB WATCH>
1026       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1027       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ildev"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "tmo", tmo
            vbwReportVariable "eot", eot
            vbwReportVariable "eos", eos
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function


Function ildma(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1028       On Error GoTo vbwErrHandler
1029       Const VBWPROCNAME = "VBIB32.ildma"
1030       If vbwProtector.vbwTraceProc Then
1031           Dim vbwProtectorParameterString As String
1032           If vbwProtector.vbwTraceParameters Then
1033               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1034               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1035           End If
1036           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1037       End If
' </VB WATCH>
1038       If (GPIBglobalsRegistered = 0) Then
1039         Call RegisterGPIBGlobals
1040       End If

       ' Call the 32-bit DLL.
1041       ildma = ConvertLongToInt(ibdma32(ud, v))

1042       Call copy_ibvars
' <VB WATCH>
1043       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1044       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ildma"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ileos(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1045       On Error GoTo vbwErrHandler
1046       Const VBWPROCNAME = "VBIB32.ileos"
1047       If vbwProtector.vbwTraceProc Then
1048           Dim vbwProtectorParameterString As String
1049           If vbwProtector.vbwTraceParameters Then
1050               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1051               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1052           End If
1053           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1054       End If
' </VB WATCH>
1055       If (GPIBglobalsRegistered = 0) Then
1056         Call RegisterGPIBGlobals
1057       End If

       ' Call the 32-bit DLL.
1058       ileos = ConvertLongToInt(ibeos32(ud, v))

1059       Call copy_ibvars
' <VB WATCH>
1060       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1061       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ileos"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ileot(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1062       On Error GoTo vbwErrHandler
1063       Const VBWPROCNAME = "VBIB32.ileot"
1064       If vbwProtector.vbwTraceProc Then
1065           Dim vbwProtectorParameterString As String
1066           If vbwProtector.vbwTraceParameters Then
1067               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1068               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1069           End If
1070           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1071       End If
' </VB WATCH>
1072       If (GPIBglobalsRegistered = 0) Then
1073         Call RegisterGPIBGlobals
1074       End If

       ' Call the 32-bit DLL.
1075       ileot = ConvertLongToInt(ibeot32(ud, v))

1076       Call copy_ibvars
' <VB WATCH>
1077       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1078       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ileot"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function


Function ilfind(ByVal udname As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1079       On Error GoTo vbwErrHandler
1080       Const VBWPROCNAME = "VBIB32.ilfind"
1081       If vbwProtector.vbwTraceProc Then
1082           Dim vbwProtectorParameterString As String
1083           If vbwProtector.vbwTraceParameters Then
1084               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("udname", udname) & ") "
1085           End If
1086           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1087       End If
' </VB WATCH>
1088       If (GPIBglobalsRegistered = 0) Then
1089         Call RegisterGPIBGlobals
1090       End If

       ' Call the 32-bit DLL.
1091       ilfind = ConvertLongToInt(ibfind32(ByVal udname))

1092       Call copy_ibvars
' <VB WATCH>
1093       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1094       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilfind"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "udname", udname
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilgts(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1095       On Error GoTo vbwErrHandler
1096       Const VBWPROCNAME = "VBIB32.ilgts"
1097       If vbwProtector.vbwTraceProc Then
1098           Dim vbwProtectorParameterString As String
1099           If vbwProtector.vbwTraceParameters Then
1100               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1101               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1102           End If
1103           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1104       End If
' </VB WATCH>
1105       If (GPIBglobalsRegistered = 0) Then
1106         Call RegisterGPIBGlobals
1107       End If

       ' Call the 32-bit DLL.
1108       ilgts = ConvertLongToInt(ibgts32(ud, v))

1109       Call copy_ibvars
' <VB WATCH>
1110       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1111       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilgts"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilist(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1112       On Error GoTo vbwErrHandler
1113       Const VBWPROCNAME = "VBIB32.ilist"
1114       If vbwProtector.vbwTraceProc Then
1115           Dim vbwProtectorParameterString As String
1116           If vbwProtector.vbwTraceParameters Then
1117               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1118               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1119           End If
1120           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1121       End If
' </VB WATCH>
1122       If (GPIBglobalsRegistered = 0) Then
1123         Call RegisterGPIBGlobals
1124       End If

       ' Call the 32-bit DLL.
1125       ilist = ConvertLongToInt(ibist32(ud, v))

1126       Call copy_ibvars
' <VB WATCH>
1127       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1128       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilist"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function


Function illines(ByVal ud As Integer, lines As Integer) As Integer
' <VB WATCH>
1129       On Error GoTo vbwErrHandler
1130       Const VBWPROCNAME = "VBIB32.illines"
1131       If vbwProtector.vbwTraceProc Then
1132           Dim vbwProtectorParameterString As String
1133           If vbwProtector.vbwTraceParameters Then
1134               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1135               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("lines", lines) & ") "
1136           End If
1137           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1138       End If
' </VB WATCH>
1139       Dim tmplines As Long

       ' Check to see if GPIB Global variables are registered
1140       If (GPIBglobalsRegistered = 0) Then
1141         Call RegisterGPIBGlobals
1142       End If

       ' Call the 32-bit DLL.
1143       illines = ConvertLongToInt(iblines32(ud, tmplines))

1144       lines = ConvertLongToInt(tmplines)

1145       Call copy_ibvars
' <VB WATCH>
1146       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1147       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illines"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "lines", lines
            vbwReportVariable "tmplines", tmplines
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function illn(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer) As Integer
' <VB WATCH>
1148       On Error GoTo vbwErrHandler
1149       Const VBWPROCNAME = "VBIB32.illn"
1150       If vbwProtector.vbwTraceProc Then
1151           Dim vbwProtectorParameterString As String
1152           If vbwProtector.vbwTraceParameters Then
1153               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1154               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
1155               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
1156               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ln", ln) & ") "
1157           End If
1158           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1159       End If
' </VB WATCH>
1160       Dim tmpln As Long

       ' Check to see if GPIB Global variables are registered
1161       If (GPIBglobalsRegistered = 0) Then
1162         Call RegisterGPIBGlobals
1163       End If

       ' Call the 32-bit DLL.
1164       illn = ConvertLongToInt(ibln32(ud, pad, sad, tmpln))

1165       ln = ConvertLongToInt(tmpln)

1166       Call copy_ibvars
' <VB WATCH>
1167       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1168       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illn"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "ln", ln
            vbwReportVariable "tmpln", tmpln
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function illoc(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1169       On Error GoTo vbwErrHandler
1170       Const VBWPROCNAME = "VBIB32.illoc"
1171       If vbwProtector.vbwTraceProc Then
1172           Dim vbwProtectorParameterString As String
1173           If vbwProtector.vbwTraceParameters Then
1174               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1175           End If
1176           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1177       End If
' </VB WATCH>
1178       If (GPIBglobalsRegistered = 0) Then
1179         Call RegisterGPIBGlobals
1180       End If

       ' Call the 32-bit DLL.
1181       illoc = ConvertLongToInt(ibloc32(ud))

1182       Call copy_ibvars
' <VB WATCH>
1183       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1184       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illoc"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilonl(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1185       On Error GoTo vbwErrHandler
1186       Const VBWPROCNAME = "VBIB32.ilonl"
1187       If vbwProtector.vbwTraceProc Then
1188           Dim vbwProtectorParameterString As String
1189           If vbwProtector.vbwTraceParameters Then
1190               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1191               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1192           End If
1193           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1194       End If
' </VB WATCH>
1195       If (GPIBglobalsRegistered = 0) Then
1196         Call RegisterGPIBGlobals
1197       End If

       ' Call the 32-bit DLL.
1198       ilonl = ConvertLongToInt(ibonl32(ud, v))

1199       Call copy_ibvars
' <VB WATCH>
1200       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1201       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilonl"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilpad(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1202       On Error GoTo vbwErrHandler
1203       Const VBWPROCNAME = "VBIB32.ilpad"
1204       If vbwProtector.vbwTraceProc Then
1205           Dim vbwProtectorParameterString As String
1206           If vbwProtector.vbwTraceParameters Then
1207               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1208               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1209           End If
1210           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1211       End If
' </VB WATCH>
1212       If (GPIBglobalsRegistered = 0) Then
1213         Call RegisterGPIBGlobals
1214       End If

       ' Call the 32-bit DLL.
1215       ilpad = ConvertLongToInt(ibpad32(ud, v))

1216       Call copy_ibvars
' <VB WATCH>
1217       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1218       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilpad"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilpct(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1219       On Error GoTo vbwErrHandler
1220       Const VBWPROCNAME = "VBIB32.ilpct"
1221       If vbwProtector.vbwTraceProc Then
1222           Dim vbwProtectorParameterString As String
1223           If vbwProtector.vbwTraceParameters Then
1224               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1225           End If
1226           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1227       End If
' </VB WATCH>
1228       If (GPIBglobalsRegistered = 0) Then
1229         Call RegisterGPIBGlobals
1230       End If

       ' Call the 32-bit DLL.
1231       ilpct = ConvertLongToInt(ibpct32(ud))

1232       Call copy_ibvars
' <VB WATCH>
1233       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1234       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilpct"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function



Function ilppc(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1235       On Error GoTo vbwErrHandler
1236       Const VBWPROCNAME = "VBIB32.ilppc"
1237       If vbwProtector.vbwTraceProc Then
1238           Dim vbwProtectorParameterString As String
1239           If vbwProtector.vbwTraceParameters Then
1240               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1241               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1242           End If
1243           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1244       End If
' </VB WATCH>
1245       If (GPIBglobalsRegistered = 0) Then
1246         Call RegisterGPIBGlobals
1247       End If

       ' Call the 32-bit DLL.
1248       ilppc = ConvertLongToInt(ibppc32(ud, v))

1249       Call copy_ibvars
' <VB WATCH>
1250       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1251       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilppc"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrd(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1252       On Error GoTo vbwErrHandler
1253       Const VBWPROCNAME = "VBIB32.ilrd"
1254       If vbwProtector.vbwTraceProc Then
1255           Dim vbwProtectorParameterString As String
1256           If vbwProtector.vbwTraceParameters Then
1257               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1258               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1259               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1260           End If
1261           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1262       End If
' </VB WATCH>
1263       If (GPIBglobalsRegistered = 0) Then
1264         Call RegisterGPIBGlobals
1265       End If

       ' Call the 32-bit DLL.
1266       ilrd = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))

1267       Call copy_ibvars
' <VB WATCH>
1268       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1269       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrd"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrda(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1270       On Error GoTo vbwErrHandler
1271       Const VBWPROCNAME = "VBIB32.ilrda"
1272       If vbwProtector.vbwTraceProc Then
1273           Dim vbwProtectorParameterString As String
1274           If vbwProtector.vbwTraceParameters Then
1275               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1276               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1277               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1278           End If
1279           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1280       End If
' </VB WATCH>
1281       If (GPIBglobalsRegistered = 0) Then
1282         Call RegisterGPIBGlobals
1283       End If

       ' Call the 32-bit DLL.
1284       ilrda = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilrda = ConvertLongToInt(ibrda32(ud, ByVal buf, cnt))

1285       Call copy_ibvars
' <VB WATCH>
1286       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1287       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrda"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrdf(ByVal ud As Integer, ByVal filename As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1288       On Error GoTo vbwErrHandler
1289       Const VBWPROCNAME = "VBIB32.ilrdf"
1290       If vbwProtector.vbwTraceProc Then
1291           Dim vbwProtectorParameterString As String
1292           If vbwProtector.vbwTraceParameters Then
1293               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1294               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
1295           End If
1296           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1297       End If
' </VB WATCH>
1298       If (GPIBglobalsRegistered = 0) Then
1299         Call RegisterGPIBGlobals
1300       End If

       ' Call the 32-bit DLL.
1301       ilrdf = ConvertLongToInt(ibrdf32(ud, ByVal filename))

1302       Call copy_ibvars
' <VB WATCH>
1303       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1304       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrdf"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1305       On Error GoTo vbwErrHandler
1306       Const VBWPROCNAME = "VBIB32.ilrdi"
1307       If vbwProtector.vbwTraceProc Then
1308           Dim vbwProtectorParameterString As String
1309           If vbwProtector.vbwTraceParameters Then
1310               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1311               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1312               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1313           End If
1314           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1315       End If
' </VB WATCH>
1316       If (GPIBglobalsRegistered = 0) Then
1317         Call RegisterGPIBGlobals
1318       End If

       ' Call the 32-bit DLL.
1319       ilrdi = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

1320       Call copy_ibvars
' <VB WATCH>
1321       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1322       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrdi"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1323       On Error GoTo vbwErrHandler
1324       Const VBWPROCNAME = "VBIB32.ilrdia"
1325       If vbwProtector.vbwTraceProc Then
1326           Dim vbwProtectorParameterString As String
1327           If vbwProtector.vbwTraceParameters Then
1328               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1329               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1330               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1331           End If
1332           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1333       End If
' </VB WATCH>
1334       If (GPIBglobalsRegistered = 0) Then
1335         Call RegisterGPIBGlobals
1336       End If

       ' Call the 32-bit DLL.
1337       ilrdia = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilrdia = ConvertLongToInt(ibrda32(ud, ibuf(0), cnt))

1338       Call copy_ibvars
' <VB WATCH>
1339       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1340       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrdia"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function



Function ilrpp(ByVal ud As Integer, ppr As Integer) As Integer
' <VB WATCH>
1341       On Error GoTo vbwErrHandler
1342       Const VBWPROCNAME = "VBIB32.ilrpp"
1343       If vbwProtector.vbwTraceProc Then
1344           Dim vbwProtectorParameterString As String
1345           If vbwProtector.vbwTraceParameters Then
1346               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1347               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ppr", ppr) & ") "
1348           End If
1349           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1350       End If
' </VB WATCH>
1351       Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
1352       If (GPIBglobalsRegistered = 0) Then
1353         Call RegisterGPIBGlobals
1354       End If

       ' Call the 32-bit DLL.
1355       ilrpp = ConvertLongToInt(ibrpp32(ud, ByVal tmp_str))

1356       ppr = Asc(tmp_str)

1357       Call copy_ibvars
' <VB WATCH>
1358       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1359       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrpp"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ppr", ppr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrsc(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1360       On Error GoTo vbwErrHandler
1361       Const VBWPROCNAME = "VBIB32.ilrsc"
1362       If vbwProtector.vbwTraceProc Then
1363           Dim vbwProtectorParameterString As String
1364           If vbwProtector.vbwTraceParameters Then
1365               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1366               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1367           End If
1368           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1369       End If
' </VB WATCH>
1370       If (GPIBglobalsRegistered = 0) Then
1371         Call RegisterGPIBGlobals
1372       End If

       '  Call the 32-bit DLL.
1373       ilrsc = ConvertLongToInt(ibrsc32(ud, v))

1374       Call copy_ibvars
' <VB WATCH>
1375       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1376       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrsc"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrsp(ByVal ud As Integer, spr As Integer) As Integer
' <VB WATCH>
1377       On Error GoTo vbwErrHandler
1378       Const VBWPROCNAME = "VBIB32.ilrsp"
1379       If vbwProtector.vbwTraceProc Then
1380           Dim vbwProtectorParameterString As String
1381           If vbwProtector.vbwTraceParameters Then
1382               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1383               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("spr", spr) & ") "
1384           End If
1385           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1386       End If
' </VB WATCH>
1387       Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
1388       If (GPIBglobalsRegistered = 0) Then
1389         Call RegisterGPIBGlobals
1390       End If

       ' Call the 32-bit DLL
1391       ilrsp = ConvertLongToInt(ibrsp32(ud, ByVal tmp_str))

1392       spr = Asc(tmp_str)

1393       Call copy_ibvars
' <VB WATCH>
1394       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1395       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrsp"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "spr", spr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrsv(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1396       On Error GoTo vbwErrHandler
1397       Const VBWPROCNAME = "VBIB32.ilrsv"
1398       If vbwProtector.vbwTraceProc Then
1399           Dim vbwProtectorParameterString As String
1400           If vbwProtector.vbwTraceParameters Then
1401               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1402               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1403           End If
1404           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1405       End If
' </VB WATCH>
1406       If (GPIBglobalsRegistered = 0) Then
1407         Call RegisterGPIBGlobals
1408       End If

       ' Call the 32-bit DLL.
1409       ilrsv = ConvertLongToInt(ibrsv32(ud, v))

1410       Call copy_ibvars
' <VB WATCH>
1411       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1412       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrsv"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilsad(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1413       On Error GoTo vbwErrHandler
1414       Const VBWPROCNAME = "VBIB32.ilsad"
1415       If vbwProtector.vbwTraceProc Then
1416           Dim vbwProtectorParameterString As String
1417           If vbwProtector.vbwTraceParameters Then
1418               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1419               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1420           End If
1421           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1422       End If
' </VB WATCH>
1423       If (GPIBglobalsRegistered = 0) Then
1424         Call RegisterGPIBGlobals
1425       End If

       '  Call the 32-bit DLL.
1426       ilsad = ConvertLongToInt(ibsad32(ud, v))

1427       Call copy_ibvars
' <VB WATCH>
1428       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1429       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilsad"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilsic(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1430       On Error GoTo vbwErrHandler
1431       Const VBWPROCNAME = "VBIB32.ilsic"
1432       If vbwProtector.vbwTraceProc Then
1433           Dim vbwProtectorParameterString As String
1434           If vbwProtector.vbwTraceParameters Then
1435               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1436           End If
1437           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1438       End If
' </VB WATCH>
1439       If (GPIBglobalsRegistered = 0) Then
1440         Call RegisterGPIBGlobals
1441       End If

       '  Call the 32-bit DLL.
1442       ilsic = ConvertLongToInt(ibsic32(ud))

1443       Call copy_ibvars
' <VB WATCH>
1444       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1445       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilsic"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilsre(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1446       On Error GoTo vbwErrHandler
1447       Const VBWPROCNAME = "VBIB32.ilsre"
1448       If vbwProtector.vbwTraceProc Then
1449           Dim vbwProtectorParameterString As String
1450           If vbwProtector.vbwTraceParameters Then
1451               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1452               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1453           End If
1454           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1455       End If
' </VB WATCH>
1456       If (GPIBglobalsRegistered = 0) Then
1457         Call RegisterGPIBGlobals
1458       End If

       '  Call the 32-bit DLL.
1459       ilsre = ConvertLongToInt(ibsre32(ud, v))

1460       Call copy_ibvars
' <VB WATCH>
1461       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1462       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilsre"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilstop(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1463       On Error GoTo vbwErrHandler
1464       Const VBWPROCNAME = "VBIB32.ilstop"
1465       If vbwProtector.vbwTraceProc Then
1466           Dim vbwProtectorParameterString As String
1467           If vbwProtector.vbwTraceParameters Then
1468               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1469           End If
1470           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1471       End If
' </VB WATCH>
1472       If (GPIBglobalsRegistered = 0) Then
1473         Call RegisterGPIBGlobals
1474       End If

       '  Call the 32-bit DLL.
1475       ilstop = ConvertLongToInt(ibstop32(ud))

1476       Call copy_ibvars
' <VB WATCH>
1477       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1478       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilstop"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function iltmo(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1479       On Error GoTo vbwErrHandler
1480       Const VBWPROCNAME = "VBIB32.iltmo"
1481       If vbwProtector.vbwTraceProc Then
1482           Dim vbwProtectorParameterString As String
1483           If vbwProtector.vbwTraceParameters Then
1484               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1485               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1486           End If
1487           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1488       End If
' </VB WATCH>
1489       If (GPIBglobalsRegistered = 0) Then
1490         Call RegisterGPIBGlobals
1491       End If

       '  Call the 32-bit DLL.
1492       iltmo = ConvertLongToInt(ibtmo32(ud, v))

1493       Call copy_ibvars
' <VB WATCH>
1494       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1495       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iltmo"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function iltrg(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1496       On Error GoTo vbwErrHandler
1497       Const VBWPROCNAME = "VBIB32.iltrg"
1498       If vbwProtector.vbwTraceProc Then
1499           Dim vbwProtectorParameterString As String
1500           If vbwProtector.vbwTraceParameters Then
1501               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1502           End If
1503           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1504       End If
' </VB WATCH>
1505       If (GPIBglobalsRegistered = 0) Then
1506         Call RegisterGPIBGlobals
1507       End If

       ' Call 32-bit DLL.
1508       iltrg = ConvertLongToInt(ibtrg32(ud))

1509       Call copy_ibvars
' <VB WATCH>
1510       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1511       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iltrg"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwait(ByVal ud As Integer, ByVal mask As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1512       On Error GoTo vbwErrHandler
1513       Const VBWPROCNAME = "VBIB32.ilwait"
1514       If vbwProtector.vbwTraceProc Then
1515           Dim vbwProtectorParameterString As String
1516           If vbwProtector.vbwTraceParameters Then
1517               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1518               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("mask", mask) & ") "
1519           End If
1520           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1521       End If
' </VB WATCH>
1522       If (GPIBglobalsRegistered = 0) Then
1523         Call RegisterGPIBGlobals
1524       End If

       ' Call the 32-bit DLL.
1525       ilwait = ConvertLongToInt(ibwait32(ud, mask))

1526       Call copy_ibvars
' <VB WATCH>
1527       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1528       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwait"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "mask", mask
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrt(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1529       On Error GoTo vbwErrHandler
1530       Const VBWPROCNAME = "VBIB32.ilwrt"
1531       If vbwProtector.vbwTraceProc Then
1532           Dim vbwProtectorParameterString As String
1533           If vbwProtector.vbwTraceParameters Then
1534               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1535               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1536               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1537           End If
1538           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1539       End If
' </VB WATCH>
1540       If (GPIBglobalsRegistered = 0) Then
1541         Call RegisterGPIBGlobals
1542       End If

       ' Call the 32-bit DLL.
1543       ilwrt = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))

1544       Call copy_ibvars
' <VB WATCH>
1545       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1546       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrt"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrta(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1547       On Error GoTo vbwErrHandler
1548       Const VBWPROCNAME = "VBIB32.ilwrta"
1549       If vbwProtector.vbwTraceProc Then
1550           Dim vbwProtectorParameterString As String
1551           If vbwProtector.vbwTraceParameters Then
1552               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1553               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1554               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1555           End If
1556           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1557       End If
' </VB WATCH>
1558       If (GPIBglobalsRegistered = 0) Then
1559         Call RegisterGPIBGlobals
1560       End If

       ' Call the 32-bit DLL.
1561       ilwrta = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))

       ' When the Visual Basic remapping solved, use this:
       '    ilwrta = ConvertLongToInt(ibwrta32(ud, ByVal buf, cnt))

1562       Call copy_ibvars

' <VB WATCH>
1563       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1564       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrta"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrtf(ByVal ud As Integer, ByVal filename As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1565       On Error GoTo vbwErrHandler
1566       Const VBWPROCNAME = "VBIB32.ilwrtf"
1567       If vbwProtector.vbwTraceProc Then
1568           Dim vbwProtectorParameterString As String
1569           If vbwProtector.vbwTraceParameters Then
1570               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1571               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
1572           End If
1573           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1574       End If
' </VB WATCH>
1575       If (GPIBglobalsRegistered = 0) Then
1576         Call RegisterGPIBGlobals
1577       End If

       ' Call the 32-bit DLL.
1578       ilwrtf = ConvertLongToInt(ibwrtf32(ud, ByVal filename))

1579       Call copy_ibvars
' <VB WATCH>
1580       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1581       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrtf"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1582       On Error GoTo vbwErrHandler
1583       Const VBWPROCNAME = "VBIB32.ilwrti"
1584       If vbwProtector.vbwTraceProc Then
1585           Dim vbwProtectorParameterString As String
1586           If vbwProtector.vbwTraceParameters Then
1587               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1588               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1589               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1590           End If
1591           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1592       End If
' </VB WATCH>
1593       If (GPIBglobalsRegistered = 0) Then
1594         Call RegisterGPIBGlobals
1595       End If

       ' Call the 32-bit DLL.
1596       ilwrti = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

1597       Call copy_ibvars
' <VB WATCH>
1598       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1599       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrti"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1600       On Error GoTo vbwErrHandler
1601       Const VBWPROCNAME = "VBIB32.ilwrtia"
1602       If vbwProtector.vbwTraceProc Then
1603           Dim vbwProtectorParameterString As String
1604           If vbwProtector.vbwTraceParameters Then
1605               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1606               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1607               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1608           End If
1609           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1610       End If
' </VB WATCH>
1611       If (GPIBglobalsRegistered = 0) Then
1612         Call RegisterGPIBGlobals
1613       End If

       ' Call the 32-bit DLL.
1614       ilwrtia = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilwrtia = ConvertLongToInt(ibwrta32(ud, ibuf(0), cnt))

1615       Call copy_ibvars
' <VB WATCH>
1616       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1617       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrtia"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function



Sub PassControl(ByVal ud As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1618       On Error GoTo vbwErrHandler
1619       Const VBWPROCNAME = "VBIB32.PassControl"
1620       If vbwProtector.vbwTraceProc Then
1621           Dim vbwProtectorParameterString As String
1622           If vbwProtector.vbwTraceParameters Then
1623               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1624               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
1625           End If
1626           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1627       End If
' </VB WATCH>
1628       If (GPIBglobalsRegistered = 0) Then
1629         Call RegisterGPIBGlobals
1630       End If

       ' Call the 32-bit DLL.
1631       Call PassControl32(ud, addr)

1632       Call copy_ibvars
' <VB WATCH>
1633       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1634       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PassControl"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Ppoll(ByVal ud As Integer, result As Integer)
' <VB WATCH>
1635       On Error GoTo vbwErrHandler
1636       Const VBWPROCNAME = "VBIB32.Ppoll"
1637       If vbwProtector.vbwTraceProc Then
1638           Dim vbwProtectorParameterString As String
1639           If vbwProtector.vbwTraceParameters Then
1640               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1641               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
1642           End If
1643           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1644       End If
' </VB WATCH>
1645       Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
1646       If (GPIBglobalsRegistered = 0) Then
1647         Call RegisterGPIBGlobals
1648       End If

       ' Call the 32-bit DLL.
1649       Call PPoll32(ud, tmpresult)

1650       result = ConvertLongToInt(tmpresult)

1651       Call copy_ibvars
' <VB WATCH>
1652       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1653       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Ppoll"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "result", result
            vbwReportVariable "tmpresult", tmpresult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub PpollConfig(ByVal ud As Integer, ByVal addr As Integer, ByVal lline As Integer, ByVal sense As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1654       On Error GoTo vbwErrHandler
1655       Const VBWPROCNAME = "VBIB32.PpollConfig"
1656       If vbwProtector.vbwTraceProc Then
1657           Dim vbwProtectorParameterString As String
1658           If vbwProtector.vbwTraceParameters Then
1659               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1660               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1661               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("lline", lline) & ", "
1662               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sense", sense) & ") "
1663           End If
1664           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1665       End If
' </VB WATCH>
1666       If (GPIBglobalsRegistered = 0) Then
1667         Call RegisterGPIBGlobals
1668       End If

       ' Call the 32-bit DLL.
1669       Call PPollConfig32(ud, addr, lline, sense)

1670       Call copy_ibvars
' <VB WATCH>
1671       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1672       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PpollConfig"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportVariable "lline", lline
            vbwReportVariable "sense", sense
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub PpollUnconfig(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1673       On Error GoTo vbwErrHandler
1674       Const VBWPROCNAME = "VBIB32.PpollUnconfig"
1675       If vbwProtector.vbwTraceProc Then
1676           Dim vbwProtectorParameterString As String
1677           If vbwProtector.vbwTraceParameters Then
1678               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1679               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1680           End If
1681           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1682       End If
' </VB WATCH>
1683       If (GPIBglobalsRegistered = 0) Then
1684         Call RegisterGPIBGlobals
1685       End If

       ' Call the 32-bit DLL.
1686       Call PPollUnconfig32(ud, addrs(0))

1687       Call copy_ibvars
' <VB WATCH>
1688       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1689       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PpollUnconfig"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub RcvRespMsg(ByVal ud As Integer, buf As String, ByVal term As Integer)
' <VB WATCH>
1690       On Error GoTo vbwErrHandler
1691       Const VBWPROCNAME = "VBIB32.RcvRespMsg"
1692       If vbwProtector.vbwTraceProc Then
1693           Dim vbwProtectorParameterString As String
1694           If vbwProtector.vbwTraceParameters Then
1695               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1696               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1697               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1698           End If
1699           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1700       End If
' </VB WATCH>
1701       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1702       If (GPIBglobalsRegistered = 0) Then
1703         Call RegisterGPIBGlobals
1704       End If

1705       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1706       Call RcvRespMsg32(ud, ByVal buf, cnt, term)

1707       Call copy_ibvars
' <VB WATCH>
1708       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1709       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "RcvRespMsg"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ReadStatusByte(ByVal ud As Integer, ByVal addr As Integer, result As Integer)
' <VB WATCH>
1710       On Error GoTo vbwErrHandler
1711       Const VBWPROCNAME = "VBIB32.ReadStatusByte"
1712       If vbwProtector.vbwTraceProc Then
1713           Dim vbwProtectorParameterString As String
1714           If vbwProtector.vbwTraceParameters Then
1715               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1716               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1717               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
1718           End If
1719           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1720       End If
' </VB WATCH>
1721       Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
1722       If (GPIBglobalsRegistered = 0) Then
1723         Call RegisterGPIBGlobals
1724       End If

       ' Call the 32-bit DLL.
1725       Call ReadStatusByte32(ud, addr, tmpresult)

1726       result = ConvertLongToInt(tmpresult)

1727       Call copy_ibvars
' <VB WATCH>
1728       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1729       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ReadStatusByte"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportVariable "result", result
            vbwReportVariable "tmpresult", tmpresult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Receive(ByVal ud As Integer, ByVal addr As Integer, buf As String, ByVal term As Integer)
' <VB WATCH>
1730       On Error GoTo vbwErrHandler
1731       Const VBWPROCNAME = "VBIB32.Receive"
1732       If vbwProtector.vbwTraceProc Then
1733           Dim vbwProtectorParameterString As String
1734           If vbwProtector.vbwTraceParameters Then
1735               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1736               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1737               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1738               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1739           End If
1740           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1741       End If
' </VB WATCH>
1742       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1743       If (GPIBglobalsRegistered = 0) Then
1744         Call RegisterGPIBGlobals
1745       End If

1746       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1747       Call Receive32(ud, addr, ByVal buf, cnt, term)

1748       Call copy_ibvars
' <VB WATCH>
1749       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1750       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Receive"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ReceiveSetup(ByVal ud As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1751       On Error GoTo vbwErrHandler
1752       Const VBWPROCNAME = "VBIB32.ReceiveSetup"
1753       If vbwProtector.vbwTraceProc Then
1754           Dim vbwProtectorParameterString As String
1755           If vbwProtector.vbwTraceParameters Then
1756               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1757               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
1758           End If
1759           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1760       End If
' </VB WATCH>
1761       If (GPIBglobalsRegistered = 0) Then
1762         Call RegisterGPIBGlobals
1763       End If

       ' Call the 32-bit DLL.
1764       Call ReceiveSetup32(ud, addr)

1765       Call copy_ibvars
' <VB WATCH>
1766       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1767       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ReceiveSetup"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ResetSys(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1768       On Error GoTo vbwErrHandler
1769       Const VBWPROCNAME = "VBIB32.ResetSys"
1770       If vbwProtector.vbwTraceProc Then
1771           Dim vbwProtectorParameterString As String
1772           If vbwProtector.vbwTraceParameters Then
1773               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1774               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1775           End If
1776           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1777       End If
' </VB WATCH>
1778       If (GPIBglobalsRegistered = 0) Then
1779         Call RegisterGPIBGlobals
1780       End If

       ' Call the 32-bit DLL.
1781       Call ResetSys32(ud, addrs(0))

1782       Call copy_ibvars
' <VB WATCH>
1783       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1784       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ResetSys"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Send(ByVal ud As Integer, ByVal addr As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
1785       On Error GoTo vbwErrHandler
1786       Const VBWPROCNAME = "VBIB32.Send"
1787       If vbwProtector.vbwTraceProc Then
1788           Dim vbwProtectorParameterString As String
1789           If vbwProtector.vbwTraceParameters Then
1790               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1791               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1792               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1793               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1794           End If
1795           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1796       End If
' </VB WATCH>
1797       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1798       If (GPIBglobalsRegistered = 0) Then
1799         Call RegisterGPIBGlobals
1800       End If

1801       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1802       Call Send32(ud, addr, ByVal buf, cnt, term)

1803       Call copy_ibvars
' <VB WATCH>
1804       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1805       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Send"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendCmds(ByVal ud As Integer, ByVal cmdbuf As String)
' <VB WATCH>
1806       On Error GoTo vbwErrHandler
1807       Const VBWPROCNAME = "VBIB32.SendCmds"
1808       If vbwProtector.vbwTraceProc Then
1809           Dim vbwProtectorParameterString As String
1810           If vbwProtector.vbwTraceParameters Then
1811               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1812               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cmdbuf", cmdbuf) & ") "
1813           End If
1814           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1815       End If
' </VB WATCH>
1816       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1817       If (GPIBglobalsRegistered = 0) Then
1818         Call RegisterGPIBGlobals
1819       End If

1820       cnt = CLng(Len(cmdbuf))

       ' Call the 32-bit DLL.
1821       Call SendCmds32(ud, ByVal cmdbuf, cnt)

1822       Call copy_ibvars
' <VB WATCH>
1823       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1824       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendCmds"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "cmdbuf", cmdbuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendDataBytes(ByVal ud As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
1825       On Error GoTo vbwErrHandler
1826       Const VBWPROCNAME = "VBIB32.SendDataBytes"
1827       If vbwProtector.vbwTraceProc Then
1828           Dim vbwProtectorParameterString As String
1829           If vbwProtector.vbwTraceParameters Then
1830               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1831               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1832               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1833           End If
1834           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1835       End If
' </VB WATCH>
1836       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1837       If (GPIBglobalsRegistered = 0) Then
1838         Call RegisterGPIBGlobals
1839       End If

1840       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1841       Call SendDataBytes32(ud, ByVal buf, cnt, term)

1842       Call copy_ibvars
' <VB WATCH>
1843       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1844       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendDataBytes"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendIFC(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1845       On Error GoTo vbwErrHandler
1846       Const VBWPROCNAME = "VBIB32.SendIFC"
1847       If vbwProtector.vbwTraceProc Then
1848           Dim vbwProtectorParameterString As String
1849           If vbwProtector.vbwTraceParameters Then
1850               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1851           End If
1852           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1853       End If
' </VB WATCH>
1854       If (GPIBglobalsRegistered = 0) Then
1855         Call RegisterGPIBGlobals
1856       End If

       ' Call the 32-bit DLL.
1857       Call SendIFC32(ud)

1858       Call copy_ibvars
' <VB WATCH>
1859       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1860       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendIFC"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendList(ByVal ud As Integer, addr() As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
1861       On Error GoTo vbwErrHandler
1862       Const VBWPROCNAME = "VBIB32.SendList"
1863       If vbwProtector.vbwTraceProc Then
1864           Dim vbwProtectorParameterString As String
1865           If vbwProtector.vbwTraceParameters Then
1866               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1867               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1868               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1869               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1870           End If
1871           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1872       End If
' </VB WATCH>
1873       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1874       If (GPIBglobalsRegistered = 0) Then
1875         Call RegisterGPIBGlobals
1876       End If

1877       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1878       Call SendList32(ud, addr(0), ByVal buf, cnt, term)

1879       Call copy_ibvars
' <VB WATCH>
1880       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1881       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendList"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendLLO(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1882       On Error GoTo vbwErrHandler
1883       Const VBWPROCNAME = "VBIB32.SendLLO"
1884       If vbwProtector.vbwTraceProc Then
1885           Dim vbwProtectorParameterString As String
1886           If vbwProtector.vbwTraceParameters Then
1887               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1888           End If
1889           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1890       End If
' </VB WATCH>
1891       If (GPIBglobalsRegistered = 0) Then
1892         Call RegisterGPIBGlobals
1893       End If

       ' Call the 32-bit DLL.
1894       Call SendLLO32(ud)

1895       Call copy_ibvars
' <VB WATCH>
1896       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1897       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendLLO"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendSetup(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1898       On Error GoTo vbwErrHandler
1899       Const VBWPROCNAME = "VBIB32.SendSetup"
1900       If vbwProtector.vbwTraceProc Then
1901           Dim vbwProtectorParameterString As String
1902           If vbwProtector.vbwTraceParameters Then
1903               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1904               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1905           End If
1906           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1907       End If
' </VB WATCH>
1908       If (GPIBglobalsRegistered = 0) Then
1909         Call RegisterGPIBGlobals
1910       End If

       ' Call the 32-bit DLL.
1911       Call SendSetup32(ud, addrs(0))

1912       Call copy_ibvars
' <VB WATCH>
1913       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1914       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendSetup"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SetRWLS(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1915       On Error GoTo vbwErrHandler
1916       Const VBWPROCNAME = "VBIB32.SetRWLS"
1917       If vbwProtector.vbwTraceProc Then
1918           Dim vbwProtectorParameterString As String
1919           If vbwProtector.vbwTraceParameters Then
1920               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1921               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1922           End If
1923           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1924       End If
' </VB WATCH>
1925       If (GPIBglobalsRegistered = 0) Then
1926         Call RegisterGPIBGlobals
1927       End If

       ' Call the 32-bit DLL.
1928       Call SetRWLS32(ud, addrs(0))

1929       Call copy_ibvars
' <VB WATCH>
1930       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1931       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetRWLS"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub TestSRQ(ByVal ud As Integer, result As Integer)
' <VB WATCH>
1932       On Error GoTo vbwErrHandler
1933       Const VBWPROCNAME = "VBIB32.TestSRQ"
1934       If vbwProtector.vbwTraceProc Then
1935           Dim vbwProtectorParameterString As String
1936           If vbwProtector.vbwTraceParameters Then
1937               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1938               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
1939           End If
1940           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1941       End If
' </VB WATCH>
1942       Call ibwait(ud, 0)

1943       If ibsta And &H1000 Then
1944           result = 1
1945       Else
1946           result = 0
1947       End If

' <VB WATCH>
1948       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1949       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TestSRQ"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "result", result
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub TestSys(ByVal ud As Integer, addrs() As Integer, results() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1950       On Error GoTo vbwErrHandler
1951       Const VBWPROCNAME = "VBIB32.TestSys"
1952       If vbwProtector.vbwTraceProc Then
1953           Dim vbwProtectorParameterString As String
1954           If vbwProtector.vbwTraceParameters Then
1955               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1956               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
1957               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("results", results) & ") "
1958           End If
1959           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1960       End If
' </VB WATCH>
1961       If (GPIBglobalsRegistered = 0) Then
1962         Call RegisterGPIBGlobals
1963       End If

       ' Call the 32-bit DLL.
1964       Call TestSys32(ud, addrs(0), results(0))

1965       Call copy_ibvars
' <VB WATCH>
1966       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1967       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TestSys"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportVariable "results", results
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Trigger(ByVal ud As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1968       On Error GoTo vbwErrHandler
1969       Const VBWPROCNAME = "VBIB32.Trigger"
1970       If vbwProtector.vbwTraceProc Then
1971           Dim vbwProtectorParameterString As String
1972           If vbwProtector.vbwTraceParameters Then
1973               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1974               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
1975           End If
1976           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1977       End If
' </VB WATCH>
1978       If (GPIBglobalsRegistered = 0) Then
1979         Call RegisterGPIBGlobals
1980       End If

       ' Call the 32-bit DLL.
1981       Call Trigger32(ud, addr)

1982       Call copy_ibvars
' <VB WATCH>
1983       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1984       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Trigger"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub TriggerList(ByVal ud As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1985       On Error GoTo vbwErrHandler
1986       Const VBWPROCNAME = "VBIB32.TriggerList"
1987       If vbwProtector.vbwTraceProc Then
1988           Dim vbwProtectorParameterString As String
1989           If vbwProtector.vbwTraceParameters Then
1990               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1991               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1992           End If
1993           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1994       End If
' </VB WATCH>
1995       If (GPIBglobalsRegistered = 0) Then
1996         Call RegisterGPIBGlobals
1997       End If

       ' Call the 32-bit DLL.
1998       Call TriggerList32(ud, addrs(0))

1999       Call copy_ibvars
' <VB WATCH>
2000       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2001       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TriggerList"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub WaitSRQ(ByVal ud As Integer, result As Integer)
' <VB WATCH>
2002       On Error GoTo vbwErrHandler
2003       Const VBWPROCNAME = "VBIB32.WaitSRQ"
2004       If vbwProtector.vbwTraceProc Then
2005           Dim vbwProtectorParameterString As String
2006           If vbwProtector.vbwTraceParameters Then
2007               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
2008               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
2009           End If
2010           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2011       End If
' </VB WATCH>
2012       Call ibwait(ud, &H5000)

2013       If ibsta And &H1000 Then
2014           result = 1
2015       Else
2016           result = 0
2017       End If
' <VB WATCH>
2018       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2019       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WaitSRQ"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "result", result
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Private Function ConvertLongToInt(LongNumb As Long) As Integer
' <VB WATCH>
2020       On Error GoTo vbwErrHandler
2021       Const VBWPROCNAME = "VBIB32.ConvertLongToInt"
2022       If vbwProtector.vbwTraceProc Then
2023           Dim vbwProtectorParameterString As String
2024           If vbwProtector.vbwTraceParameters Then
2025               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("LongNumb", LongNumb) & ") "
2026           End If
2027           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2028       End If
' </VB WATCH>

2029     If (LongNumb And &H8000&) = 0 Then
2030         ConvertLongToInt = LongNumb And &HFFFF&
2031     Else
2032       ConvertLongToInt = &H8000 Or (LongNumb And &H7FFF&)
2033     End If

' <VB WATCH>
2034       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2035       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ConvertLongToInt"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "LongNumb", LongNumb
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Sub RegisterGPIBGlobals()
' <VB WATCH>
2036       On Error GoTo vbwErrHandler
2037       Const VBWPROCNAME = "VBIB32.RegisterGPIBGlobals"
2038       If vbwProtector.vbwTraceProc Then
2039           Dim vbwProtectorParameterString As String
2040           If vbwProtector.vbwTraceParameters Then
2041               vbwProtectorParameterString = "()"
2042           End If
2043           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2044       End If
' </VB WATCH>
2045       Dim rc As Long

2046       rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
2047       If (rc = 0) Then
2048         GPIBglobalsRegistered = 1
2049       ElseIf (rc = 1) Then
2050         rc = UnregisterGpibGlobalsForThread
2051         rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
2052         GPIBglobalsRegistered = 1
2053       ElseIf (rc = 2) Then
2054         rc = UnregisterGpibGlobalsForThread
2055         ibsta = &H8000
2056         iberr = EDVR
2057         ibcntl = &HDEAD37F0
2058       ElseIf (rc = 3) Then
2059         rc = UnregisterGpibGlobalsForThread
2060         ibsta = &H8000
2061         iberr = EDVR
2062         ibcntl = &HDEAD37F0
2063       Else
2064         ibsta = &H8000
2065         iberr = EDVR
2066         ibcntl = &HDEAD37F0
2067       End If
' <VB WATCH>
2068       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2069       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "RegisterGPIBGlobals"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "rc", rc
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub UnregisterGPIBGlobals()
' <VB WATCH>
2070       On Error GoTo vbwErrHandler
2071       Const VBWPROCNAME = "VBIB32.UnregisterGPIBGlobals"
2072       If vbwProtector.vbwTraceProc Then
2073           Dim vbwProtectorParameterString As String
2074           If vbwProtector.vbwTraceParameters Then
2075               vbwProtectorParameterString = "()"
2076           End If
2077           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2078       End If
' </VB WATCH>
2079       Dim rc As Long

2080       rc = UnregisterGpibGlobalsForThread
2081       GPIBglobalsRegistered = 0

' <VB WATCH>
2082       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2083       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "UnregisterGPIBGlobals"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "rc", rc
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Public Function ThreadIbsta() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
2084       On Error GoTo vbwErrHandler
2085       Const VBWPROCNAME = "VBIB32.ThreadIbsta"
2086       If vbwProtector.vbwTraceProc Then
2087           Dim vbwProtectorParameterString As String
2088           If vbwProtector.vbwTraceParameters Then
2089               vbwProtectorParameterString = "()"
2090           End If
2091           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2092       End If
' </VB WATCH>
2093       ThreadIbsta = ConvertLongToInt(ThreadIbsta32())
' <VB WATCH>
2094       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2095       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIbsta"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function ThreadIberr() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
2096       On Error GoTo vbwErrHandler
2097       Const VBWPROCNAME = "VBIB32.ThreadIberr"
2098       If vbwProtector.vbwTraceProc Then
2099           Dim vbwProtectorParameterString As String
2100           If vbwProtector.vbwTraceParameters Then
2101               vbwProtectorParameterString = "()"
2102           End If
2103           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2104       End If
' </VB WATCH>
2105       ThreadIberr = ConvertLongToInt(ThreadIberr32())
' <VB WATCH>
2106       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2107       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIberr"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function ThreadIbcnt() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
2108       On Error GoTo vbwErrHandler
2109       Const VBWPROCNAME = "VBIB32.ThreadIbcnt"
2110       If vbwProtector.vbwTraceProc Then
2111           Dim vbwProtectorParameterString As String
2112           If vbwProtector.vbwTraceParameters Then
2113               vbwProtectorParameterString = "()"
2114           End If
2115           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2116       End If
' </VB WATCH>
2117       ThreadIbcnt = ConvertLongToInt(ThreadIbcnt32())
' <VB WATCH>
2118       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2119       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIbcnt"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function ThreadIbcntl() As Long
       ' Call the 32-bit DLL.
' <VB WATCH>
2120       On Error GoTo vbwErrHandler
2121       Const VBWPROCNAME = "VBIB32.ThreadIbcntl"
2122       If vbwProtector.vbwTraceProc Then
2123           Dim vbwProtectorParameterString As String
2124           If vbwProtector.vbwTraceParameters Then
2125               vbwProtectorParameterString = "()"
2126           End If
2127           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2128       End If
' </VB WATCH>
2129       ThreadIbcntl = ThreadIbcntl32()
' <VB WATCH>
2130       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2131       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIbcntl"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function illock(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2132       On Error GoTo vbwErrHandler
2133       Const VBWPROCNAME = "VBIB32.illock"
2134       If vbwProtector.vbwTraceProc Then
2135           Dim vbwProtectorParameterString As String
2136           If vbwProtector.vbwTraceParameters Then
2137               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2138           End If
2139           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2140       End If
' </VB WATCH>
2141       If (GPIBglobalsRegistered = 0) Then
2142         Call RegisterGPIBGlobals
2143       End If

       ' Call the 32-bit DLL.
2144       illock = ConvertLongToInt(iblock32(ud))

2145       Call copy_ibvars
' <VB WATCH>
2146       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2147       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illock"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function ilunlock(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2148       On Error GoTo vbwErrHandler
2149       Const VBWPROCNAME = "VBIB32.ilunlock"
2150       If vbwProtector.vbwTraceProc Then
2151           Dim vbwProtectorParameterString As String
2152           If vbwProtector.vbwTraceParameters Then
2153               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2154           End If
2155           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2156       End If
' </VB WATCH>
2157       If (GPIBglobalsRegistered = 0) Then
2158         Call RegisterGPIBGlobals
2159       End If

       ' Call the 32-bit DLL.
2160       ilunlock = ConvertLongToInt(ibunlock32(ud))

2161       Call copy_ibvars
' <VB WATCH>
2162       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2163       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilunlock"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Sub iblock(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2164       On Error GoTo vbwErrHandler
2165       Const VBWPROCNAME = "VBIB32.iblock"
2166       If vbwProtector.vbwTraceProc Then
2167           Dim vbwProtectorParameterString As String
2168           If vbwProtector.vbwTraceParameters Then
2169               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2170           End If
2171           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2172       End If
' </VB WATCH>
2173       If (GPIBglobalsRegistered = 0) Then
2174         Call RegisterGPIBGlobals
2175       End If

       ' Call the 32-bit DLL.
2176       Call iblock32(ud)

2177       Call copy_ibvars
' <VB WATCH>
2178       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2179       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iblock"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub ibunlock(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2180       On Error GoTo vbwErrHandler
2181       Const VBWPROCNAME = "VBIB32.ibunlock"
2182       If vbwProtector.vbwTraceProc Then
2183           Dim vbwProtectorParameterString As String
2184           If vbwProtector.vbwTraceParameters Then
2185               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2186           End If
2187           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2188       End If
' </VB WATCH>
2189       If (GPIBglobalsRegistered = 0) Then
2190         Call RegisterGPIBGlobals
2191       End If

       ' Call the 32-bit DLL.
2192       Call ibunlock32(ud)

2193       Call copy_ibvars
' <VB WATCH>
2194       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2195       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibunlock"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Function illockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2196       On Error GoTo vbwErrHandler
2197       Const VBWPROCNAME = "VBIB32.illockx"
2198       If vbwProtector.vbwTraceProc Then
2199           Dim vbwProtectorParameterString As String
2200           If vbwProtector.vbwTraceParameters Then
2201               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
2202               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LockWaitTime", LockWaitTime) & ", "
2203               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
2204           End If
2205           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2206       End If
' </VB WATCH>
2207       If (GPIBglobalsRegistered = 0) Then
2208         Call RegisterGPIBGlobals
2209       End If

       ' Call the 32-bit DLL.
2210       illockx = ConvertLongToInt(iblockx32(ud, LockWaitTime, buf))

2211       Call copy_ibvars
' <VB WATCH>
2212       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2213       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illockx"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "LockWaitTime", LockWaitTime
            vbwReportVariable "buf", buf
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function ilunlockx(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2214       On Error GoTo vbwErrHandler
2215       Const VBWPROCNAME = "VBIB32.ilunlockx"
2216       If vbwProtector.vbwTraceProc Then
2217           Dim vbwProtectorParameterString As String
2218           If vbwProtector.vbwTraceParameters Then
2219               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2220           End If
2221           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2222       End If
' </VB WATCH>
2223       If (GPIBglobalsRegistered = 0) Then
2224         Call RegisterGPIBGlobals
2225       End If

       ' Call the 32-bit DLL.
2226       ilunlockx = ConvertLongToInt(ibunlockx32(ud))

2227       Call copy_ibvars
' <VB WATCH>
2228       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2229       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilunlockx"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Sub iblockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2230       On Error GoTo vbwErrHandler
2231       Const VBWPROCNAME = "VBIB32.iblockx"
2232       If vbwProtector.vbwTraceProc Then
2233           Dim vbwProtectorParameterString As String
2234           If vbwProtector.vbwTraceParameters Then
2235               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
2236               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LockWaitTime", LockWaitTime) & ", "
2237               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
2238           End If
2239           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2240       End If
' </VB WATCH>
2241       If (GPIBglobalsRegistered = 0) Then
2242         Call RegisterGPIBGlobals
2243       End If

       ' Call the 32-bit DLL.
2244       Call iblockx32(ud, LockWaitTime, buf)

2245       Call copy_ibvars
' <VB WATCH>
2246       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2247       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iblockx"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "LockWaitTime", LockWaitTime
            vbwReportVariable "buf", buf
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub ibunlockx(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2248       On Error GoTo vbwErrHandler
2249       Const VBWPROCNAME = "VBIB32.ibunlockx"
2250       If vbwProtector.vbwTraceProc Then
2251           Dim vbwProtectorParameterString As String
2252           If vbwProtector.vbwTraceParameters Then
2253               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2254           End If
2255           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2256       End If
' </VB WATCH>
2257       If (GPIBglobalsRegistered = 0) Then
2258         Call RegisterGPIBGlobals
2259       End If

       ' Call the 32-bit DLL.
2260       Call ibunlockx32(ud)

2261       Call copy_ibvars
' <VB WATCH>
2262       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2263       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibunlockx"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
