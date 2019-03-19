VERSION 5.00
Begin VB.Form frmReportOptions 
   Caption         =   "Select Report Options"
   ClientHeight    =   3915
   ClientLeft      =   9960
   ClientTop       =   6000
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4680
   Begin VB.CheckBox chkTRG 
      Caption         =   "TRG Reading"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox chkVibration 
      Caption         =   "Vibration"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdOptionsOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.CheckBox chkSelectCircFlow 
      Caption         =   "Circulation Flow"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1880
      Width           =   1575
   End
   Begin VB.CheckBox chkSelectAxPos 
      Caption         =   "Axial Position"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1360
      Width           =   1455
   End
   Begin VB.CheckBox chkSelectRPM 
      Caption         =   "RPM"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please select the options that you want on the report."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmReportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' <VB WATCH>
Const VBWMODULE = "frmReportOptions"
' </VB WATCH>

Private Sub cmdOptionsOK_Click()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmReportOptions.cmdOptionsOK_Click"
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
    Const VBWPROCEDURE = "cmdOptionsOK_Click"

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
