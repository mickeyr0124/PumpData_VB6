VERSION 5.00
Begin VB.Form FrmSODetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Details"
   ClientHeight    =   5265
   ClientLeft      =   6885
   ClientTop       =   5535
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtSOData 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"FrmSODetails.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "FrmSODetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <VB WATCH>
Const VBWMODULE = "FrmSODetails"
' </VB WATCH>

Private Sub cmdClose_Click()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "FrmSODetails.cmdClose_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>
10         Me.Hide
' <VB WATCH>
11         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
12         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdClose_Click"

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

Private Sub Form_Activate()
' <VB WATCH>
13         On Error GoTo vbwErrHandler
14         Const VBWPROCNAME = "FrmSODetails.Form_Activate"
15         If vbwProtector.vbwTraceProc Then
16             Dim vbwProtectorParameterString As String
17             If vbwProtector.vbwTraceParameters Then
18                 vbwProtectorParameterString = "()"
19             End If
20             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
21         End If
' </VB WATCH>
22     Const HWND_TOPMOST As Integer = -1
       'Const HWND_NOTOPMOST As Integer = -2
23     Const SWP_NOSIZE As Integer = &H1
24     Const SWP_NOMOVE As Integer = &H2
25     Const SWP_NOACTIVATE As Integer = &H10
26     Const SWP_SHOWWINDOW As Integer = &H40

27         SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

       '    SetWindowPos Me.hWnd, -1, 0, 0, 520, 400, &H40
' <VB WATCH>
28         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
29         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Activate"

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


' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
