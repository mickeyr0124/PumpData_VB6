VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEmail 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4680
      TabIndex        =   19
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send E-mail When There Is A Change To:"
      Height          =   3495
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   9495
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Performance Modifications"
         Height          =   2655
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   4695
         Begin VB.CheckBox chkDischDia 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discharge Orifice Value"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1920
            Width           =   2775
         End
         Begin VB.CheckBox chkImpDia 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Diameter Value"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CheckBox chkPumpDisch 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pump Discharge Orifice Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox chkImpTrimmed 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Trimmed Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox chkImpFeathered 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Feathered Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Thrust Balance Modifications"
         Height          =   2655
         Left            =   5400
         TabIndex        =   4
         Top             =   480
         Width           =   3855
         Begin VB.CheckBox chkOtherMods 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Other Mods"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CheckBox chkCircFlowDia 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flow Orifice Value"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   3135
         End
         Begin VB.CheckBox ChkBalHoleMod 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Balance Holes Modifications"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CheckBox chkCircOrifice 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flow Orifice Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1200
            Width           =   2775
         End
         Begin VB.CheckBox ChkBalHoles 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Balance Holes Modified Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   2895
         End
         Begin VB.CheckBox chkEndPlay 
            BackColor       =   &H00FFFFC0&
            Caption         =   "End Play Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   1815
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "List of E-mails"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset
Dim qyData As New ADODB.Command
Dim rsData1 As New ADODB.Recordset
Dim qyData1 As New ADODB.Command

' <VB WATCH>
Const VBWMODULE = "frmEmail"
' </VB WATCH>

Private Sub cmdAdd_Click()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmEmail.cmdAdd_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>

10         rsData1.AddNew "email", "New Addition"
11         qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
12         Set MSHFlexGrid1.DataSource = rsData
13         MShFlexGrid1_Click
' <VB WATCH>
14         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
15         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdAdd_Click"

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

Private Sub cmdDelete_Click()
' <VB WATCH>
16         On Error GoTo vbwErrHandler
17         Const VBWPROCNAME = "frmEmail.cmdDelete_Click"
18         If vbwProtector.vbwTraceProc Then
19             Dim vbwProtectorParameterString As String
20             If vbwProtector.vbwTraceParameters Then
21                 vbwProtectorParameterString = "()"
22             End If
23             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
24         End If
' </VB WATCH>
25         rsData1.Delete
26         qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
27         Set MSHFlexGrid1.DataSource = rsData
28         MShFlexGrid1_Click
' <VB WATCH>
29         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
30         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdDelete_Click"

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

Private Sub cmdUpdate_Click()
' <VB WATCH>
31         On Error GoTo vbwErrHandler
32         Const VBWPROCNAME = "frmEmail.cmdUpdate_Click"
33         If vbwProtector.vbwTraceProc Then
34             Dim vbwProtectorParameterString As String
35             If vbwProtector.vbwTraceParameters Then
36                 vbwProtectorParameterString = "()"
37             End If
38             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
39         End If
' </VB WATCH>
40         rsData1.Update
41         qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
42         Set MSHFlexGrid1.DataSource = rsData
43         MShFlexGrid1_Click
' <VB WATCH>
44         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
45         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdUpdate_Click"

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

Private Sub Form_Load()
' <VB WATCH>
46         On Error GoTo vbwErrHandler
47         Const VBWPROCNAME = "frmEmail.Form_Load"
48         If vbwProtector.vbwTraceProc Then
49             Dim vbwProtectorParameterString As String
50             If vbwProtector.vbwTraceParameters Then
51                 vbwProtectorParameterString = "()"
52             End If
53             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
54         End If
' </VB WATCH>
55         MSHFlexGrid1.ColWidth(0) = 3000
56         MSHFlexGrid1.RowSel = 1

57         qyData.ActiveConnection = cnPumpData
58         qyData1.ActiveConnection = cnPumpData
59         qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
60         rsData.CursorType = adOpenStatic
61         rsData.CursorLocation = adUseClient
62         rsData1.CursorType = adOpenStatic
63         rsData1.CursorLocation = adUseClient
64         rsData1.LockType = adLockPessimistic
65         rsData.Open qyData

66         Set MSHFlexGrid1.DataSource = rsData
67         MSHFlexGrid1.RowSel = 1
68         MShFlexGrid1_Click

' <VB WATCH>
69         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
70         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Load"

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

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
71         On Error GoTo vbwErrHandler
72         Const VBWPROCNAME = "frmEmail.Form_Unload"
73         If vbwProtector.vbwTraceProc Then
74             Dim vbwProtectorParameterString As String
75             If vbwProtector.vbwTraceParameters Then
76                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
77             End If
78             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
79         End If
' </VB WATCH>
80         If rsData1.State = adStateOpen Then
81             rsData1.Close
82             Set rsData1 = Nothing
83             Set qyData1 = Nothing
84         End If
85         If rsData.State = adStateOpen Then
86             rsData.Close
87             Set rsData = Nothing
88             Set qyData = Nothing
89         End If

' <VB WATCH>
90         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
91         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Unload"

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
            vbwReportVariable "Cancel", Cancel
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub MShFlexGrid1_Click()
' <VB WATCH>
92         On Error GoTo vbwErrHandler
93         Const VBWPROCNAME = "frmEmail.MShFlexGrid1_Click"
94         If vbwProtector.vbwTraceProc Then
95             Dim vbwProtectorParameterString As String
96             If vbwProtector.vbwTraceParameters Then
97                 vbwProtectorParameterString = "()"
98             End If
99             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
100        End If
' </VB WATCH>


101        MSHFlexGrid1.Row = MSHFlexGrid1.RowSel
102        MSHFlexGrid1.Col = 0
103        FixHighlight (MSHFlexGrid1.Row)

104        If rsData1.State = 1 Then
105            rsData1.Close
106        End If

107        qyData1.CommandText = "SELECT * FROM ModificationEmail WHERE ModificationEmail.Email = '" & MSHFlexGrid1.Text & "' ORDER BY Email ;"
108        rsData1.Open qyData1

109        Set Text1.DataSource = rsData1
110        Text1.DataField = "Email"
111        Set chkImpFeathered.DataSource = rsData1
112        chkImpFeathered.DataField = "ImpellerFeathered"
113        Set chkImpTrimmed.DataSource = rsData1
114        chkImpTrimmed.DataField = "ImpellerTrimmed"
115        Set chkPumpDisch.DataSource = rsData1
116        chkPumpDisch.DataField = "DischargeOrifice"
117        Set chkImpDia.DataSource = rsData1
118        chkImpDia.DataField = "ImpellerDiameter"
119        Set chkDischDia.DataSource = rsData1
120        chkDischDia.DataField = "OrificeDiameter"

121        Set chkEndPlay.DataSource = rsData1
122        chkEndPlay.DataField = "Endplay"
123        Set ChkBalHoles.DataSource = rsData1
124        ChkBalHoles.DataField = "BalanceHolesModified"
125        Set chkCircOrifice.DataSource = rsData1
126        chkCircOrifice.DataField = "CirculationFlowOrifice"
127        Set chkCircFlowDia.DataSource = rsData1
128        chkCircFlowDia.DataField = "CirculationFlowDiameter"
129        Set chkOtherMods.DataSource = rsData1
130        chkOtherMods.DataField = "OtherMods"
131        Set ChkBalHoleMod.DataSource = rsData1
132        ChkBalHoleMod.DataField = "BalanceHoleModifications"

' <VB WATCH>
133        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
134        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "MShFlexGrid1_Click"

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
Sub FixHighlight(Row As Integer)
' <VB WATCH>
135        On Error GoTo vbwErrHandler
136        Const VBWPROCNAME = "frmEmail.FixHighlight"
137        If vbwProtector.vbwTraceProc Then
138            Dim vbwProtectorParameterString As String
139            If vbwProtector.vbwTraceParameters Then
140                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Row", Row) & ") "
141            End If
142            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
143        End If
' </VB WATCH>
144        Dim I As Integer
145        Const PresentColor = -2147483643

146        MSHFlexGrid1.Col = 0

147        For I = 1 To MSHFlexGrid1.Rows - 1
148            MSHFlexGrid1.Row = I
149            MSHFlexGrid1.CellBackColor = PresentColor
150        Next I

151        MSHFlexGrid1.Row = Row
152        MSHFlexGrid1.CellBackColor = vbYellow
' <VB WATCH>
153        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
154        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FixHighlight"

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
            vbwReportVariable "Row", Row
            vbwReportVariable "I", I
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
    vbwReportVariable "rsData", rsData
    vbwReportVariable "qyData", qyData
    vbwReportVariable "rsData1", rsData1
    vbwReportVariable "qyData1", qyData1
End Sub
' </VB WATCH>
