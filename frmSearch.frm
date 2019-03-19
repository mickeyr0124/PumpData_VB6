VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search for Pumps"
   ClientHeight    =   12990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16545
   LinkTopic       =   "Form1"
   ScaleHeight     =   12990
   ScaleWidth      =   16545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmWildCard 
      Caption         =   "Search By Portion Of Model Number"
      Height          =   1335
      Left            =   0
      TabIndex        =   28
      Top             =   9960
      Width           =   15015
      Begin VB.TextBox txtModelNumberString 
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Text            =   "Enter Characters and Search with Return"
         Top             =   360
         Width           =   3135
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgWildCard 
         Height          =   975
         Left            =   3840
         TabIndex        =   29
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1720
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdResetSizes 
      Caption         =   "Reset Sizes"
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   11460
      Width           =   1575
   End
   Begin VB.Frame frmShipTo 
      Caption         =   "Search By Ship To Customer"
      Height          =   1335
      Left            =   0
      TabIndex        =   17
      Top             =   8520
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgShipTo 
         Height          =   975
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1720
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo cmbSearchShipTo 
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Select Customer"
      End
   End
   Begin VB.Frame frmCustomer 
      Caption         =   "Search By Bill To Customer"
      Height          =   1335
      Left            =   0
      TabIndex        =   12
      Top             =   7200
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchCustomer 
         Height          =   315
         Left            =   600
         TabIndex        =   13
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Select Customer"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCustomer 
         Height          =   975
         Left            =   3840
         TabIndex        =   20
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1720
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frmSalesOrder 
      Caption         =   "Search By Sales Order"
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchSalesOrder 
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Select Sales Order"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgSalesOrder 
         Height          =   975
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frmTEMCFrameNo 
      Caption         =   "Search By Teikoku Frame Number"
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   5880
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchTEMCFrameNumber 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "Select TEMC Frame No"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgTEMCFrameNo 
         Height          =   975
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1720
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frmModel 
      Caption         =   "Search By Chempump Model"
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   4560
      Width           =   15015
      Begin VB.ComboBox cmbSearchModel 
         Height          =   315
         Left            =   600
         TabIndex        =   7
         Text            =   "Select Model"
         Top             =   600
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgModel 
         Height          =   975
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1720
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame frmDate 
      Caption         =   "Search By Date"
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox cmbStartDate 
         Height          =   315
         Left            =   600
         TabIndex        =   26
         Text            =   "Select Start Date"
         Top             =   720
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo cmbSearchEndDate 
         Height          =   315
         Left            =   600
         TabIndex        =   14
         Top             =   1440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Select End Date"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDate 
         Height          =   1575
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame frmSN 
      Caption         =   "Search By Serial Number"
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchSN 
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Select Serial Number"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgSN 
         Height          =   975
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1720
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   11460
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Close"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   11460
      Width           =   1335
   End
   Begin VB.Label lblNoOfPumps 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   11640
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsData As New ADODB.Recordset
Dim qyData As New ADODB.Command
Dim rsData1 As New ADODB.Recordset
Dim qyData1 As New ADODB.Command
Dim qyData2 As New ADODB.Command
Dim rsData2 As New ADODB.Recordset
Dim rsDataDate As New ADODB.Recordset
Dim qyDataDate As New ADODB.Command
Dim rsDataModel As New ADODB.Recordset
Dim qyDataModel As New ADODB.Command
Dim rsDataSN As New ADODB.Recordset
Dim qyDataSN As New ADODB.Command
Dim rsDataSalesOrder As New ADODB.Recordset
Dim qySalesOrderData As New ADODB.Command
Dim rsSalesOrderData As New ADODB.Recordset
Dim qyDataSalesOrder As New ADODB.Command
Dim rsDataTEMCModel As New ADODB.Recordset
Dim qyDataTEMCModel As New ADODB.Command
Dim rsDataTEMCFrameNumber As New ADODB.Recordset
Dim qyDataTEMCFrameNumber As New ADODB.Command
Dim rsDataCustomer As New ADODB.Recordset
Dim qyDataCustomer As New ADODB.Command
Dim rsCustomerData As New ADODB.Recordset
Dim qyCustomerData As New ADODB.Command
Dim rsDataShipTo As New ADODB.Recordset
Dim qyDataShipto As New ADODB.Command
Dim rsShipToData As New ADODB.Recordset
Dim qyShipToData As New ADODB.Command
' <VB WATCH>
Const VBWMODULE = "frmSearch"
' </VB WATCH>

Private Sub cmbSalesOrder_Click(Area As Integer)
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmSearch.cmbSalesOrder_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>

' <VB WATCH>
10         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
11         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSalesOrder_Click"

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
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchCustomer_Click(Area As Integer)
' <VB WATCH>
12         On Error GoTo vbwErrHandler
13         Const VBWPROCNAME = "frmSearch.cmbSearchCustomer_Click"
14         If vbwProtector.vbwTraceProc Then
15             Dim vbwProtectorParameterString As String
16             If vbwProtector.vbwTraceParameters Then
17                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
18             End If
19             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
20         End If
' </VB WATCH>
21         If rsCustomerData.State = adStateOpen Then
22             rsCustomerData.Close
23         End If

24         If cmbSearchCustomer.SelectedItem = 1 Then
' <VB WATCH>
25         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
26             Exit Sub
27         End If

28         qyCustomerData.CommandText = "SELECT DISTINCT " & _
           " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], TempPumpData.ShiptoCustomer " & _
           " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
           " WHERE (((TempPumpData.BillToCustomer)= '" & cmbSearchCustomer.BoundText & "'));"

29         rsCustomerData.Open qyCustomerData

30         If rsCustomerData.RecordCount = 0 Then
' <VB WATCH>
31         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
32             Exit Sub
33         End If

           'bind the datalist to the recordset
34         Set fgCustomer.DataSource = rsCustomerData

35         fgCustomer.ColWidth(0) = 1400
36         fgCustomer.ColWidth(1) = 2000
37         fgCustomer.ColWidth(2) = 1200
38         fgCustomer.ColWidth(3) = 2000
39         fgCustomer.ColWidth(4) = 3200
40         fgCustomer.TextMatrix(0, 0) = "S/N"
41         fgCustomer.TextMatrix(0, 1) = "Date"
42         fgCustomer.TextMatrix(0, 2) = "Sales Order"
43         fgCustomer.TextMatrix(0, 3) = "Model No"
44         fgCustomer.TextMatrix(0, 4) = "Ship To"

45         frmModel.Visible = False
46         frmTEMCFrameNo.Visible = False
47         frmCustomer.Top = 7200 - (4000 - 1335)
48         frmCustomer.Height = 4000
49         fgCustomer.Height = 4000 - 360
50         frmCustomer.FontBold = True

' <VB WATCH>
51         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
52         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchCustomer_Click"

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
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Private Sub cmbSearchEndDate_Click(Area As Integer)
' <VB WATCH>
53         On Error GoTo vbwErrHandler
54         Const VBWPROCNAME = "frmSearch.cmbSearchEndDate_Click"
55         If vbwProtector.vbwTraceProc Then
56             Dim vbwProtectorParameterString As String
57             If vbwProtector.vbwTraceParameters Then
58                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
59             End If
60             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
61         End If
' </VB WATCH>
62         cmbStartDate_Click
' <VB WATCH>
63         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
64         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchEndDate_Click"

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
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchModel_Click()
' <VB WATCH>
65         On Error GoTo vbwErrHandler
66         Const VBWPROCNAME = "frmSearch.cmbSearchModel_Click"
67         If vbwProtector.vbwTraceProc Then
68             Dim vbwProtectorParameterString As String
69             If vbwProtector.vbwTraceParameters Then
70                 vbwProtectorParameterString = "()"
71             End If
72             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
73         End If
' </VB WATCH>
74         If rsDataModel.State = adStateOpen Then
75             rsDataModel.Close
76         End If

77         qyDataModel.CommandText = "SELECT DISTINCT " & _
           " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
           "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
           " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
           " WHERE (((TempPumpData.Model)= " & cmbSearchModel.ItemData(cmbSearchModel.ListIndex) & "));"

78         rsDataModel.Open qyDataModel

79         If rsDataModel.RecordCount = 0 Then
' <VB WATCH>
80         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
81             Exit Sub
82         End If

83         Set fgModel.DataSource = rsDataModel

84         Dim f As String

85         f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
86         fgModel.FormatString = f
87         fgModel.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
88         fgModel.ColWidth(0) = 1200
89         fgModel.ColWidth(1) = 2000
90         fgModel.ColWidth(2) = 1200
91         fgModel.ColWidth(3) = 2000
92         fgModel.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
93         fgModel.ColWidth(5) = 3200
94         fgModel.ColWidth(6) = 3200
95         fgModel.TextMatrix(0, 0) = "S/N"
96         fgModel.TextMatrix(0, 1) = "Date"
97         fgModel.TextMatrix(0, 2) = "Sales Order"
98         fgModel.TextMatrix(0, 3) = "Model No"
99         fgModel.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
100        fgModel.TextMatrix(0, 5) = "Bill To"
101        fgModel.TextMatrix(0, 6) = "Ship To"

102        Dim X As Long
103        With fgModel
104            For X = .FixedRows To .Rows - 1
105            .TextMatrix(X, 4) = Format(.TextMatrix(X, 4), "#0.000")
106            Next X
107        End With


108        frmTEMCFrameNo.Visible = False
109        frmCustomer.Visible = False
110        frmModel.Height = 4000
111        fgModel.Height = 4000 - 360
112        frmModel.FontBold = True

' <VB WATCH>
113        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
114        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchModel_Click"

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
            vbwReportVariable "f", f
            vbwReportVariable "X", X
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchSalesOrder_Change()
' <VB WATCH>
115        On Error GoTo vbwErrHandler
116        Const VBWPROCNAME = "frmSearch.cmbSearchSalesOrder_Change"
117        If vbwProtector.vbwTraceProc Then
118            Dim vbwProtectorParameterString As String
119            If vbwProtector.vbwTraceParameters Then
120                vbwProtectorParameterString = "()"
121            End If
122            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
123        End If
' </VB WATCH>

124        Text1.Text = cmbSearchSalesOrder.BoundText

125        If rsSalesOrderData.State = adStateOpen Then
126            rsSalesOrderData.Close
127        End If

           'find all dates and models for the selected serial number
128        qySalesOrderData.CommandText = "SELECT DISTINCT " & _
           " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
           " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
           " WHERE (((TempPumpData.SalesOrderNumber)= '" & cmbSearchSalesOrder.BoundText & "'));"

129        rsSalesOrderData.Open qySalesOrderData

130        If rsSalesOrderData.RecordCount = 0 Then
' <VB WATCH>
131        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
132            Exit Sub
133        End If

       '    'bind the datalist to the other recordset
134        Set fgSalesOrder.DataSource = rsSalesOrderData

135        fgSalesOrder.ColWidth(0) = 1400               'serial number
136        fgSalesOrder.ColWidth(1) = 0               'chempumppump
137        fgSalesOrder.ColWidth(2) = 2000
138        If rsSalesOrderData.Fields(1) = True Then       'show model number
139            fgSalesOrder.ColWidth(3) = 1800
140            fgSalesOrder.ColWidth(4) = 0
141        Else                                    'else, show TEMC Frame number
142            fgSalesOrder.ColWidth(3) = 0
143            fgSalesOrder.ColWidth(4) = 1800
144        End If
145        fgSalesOrder.ColWidth(5) = 3200
146        fgSalesOrder.ColWidth(6) = 3200
147        fgSalesOrder.TextMatrix(0, 0) = "S/N"
148        fgSalesOrder.TextMatrix(0, 2) = "Date"
149        fgSalesOrder.TextMatrix(0, 3) = "Model No"
150        fgSalesOrder.TextMatrix(0, 4) = "TEMC Frame"
151        fgSalesOrder.TextMatrix(0, 5) = "Bill To"
152        fgSalesOrder.TextMatrix(0, 6) = "Ship To"

153        frmSN.Visible = False
154        frmSalesOrder.Height = 4000
155        fgSalesOrder.Height = 4000 - 360
156        frmSalesOrder.FontBold = True

' <VB WATCH>
157        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
158        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchSalesOrder_Change"

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

Private Sub cmbSearchShipTo_Click(Area As Integer)
' <VB WATCH>
159        On Error GoTo vbwErrHandler
160        Const VBWPROCNAME = "frmSearch.cmbSearchShipTo_Click"
161        If vbwProtector.vbwTraceProc Then
162            Dim vbwProtectorParameterString As String
163            If vbwProtector.vbwTraceParameters Then
164                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
165            End If
166            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
167        End If
' </VB WATCH>

168        If rsShipToData.State = adStateOpen Then
169            rsShipToData.Close
170        End If

171        If cmbSearchShipTo.SelectedItem = 1 Then
' <VB WATCH>
172        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
173            Exit Sub
174        End If

175        qyShipToData.CommandText = "SELECT DISTINCT " & _
           " [TempPumpData]![SerialNumber],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.ModelNumber, TempPumpData.BillToCustomer " & _
           " FROM  (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
           " WHERE (((TempPumpData.ShipToCustomer)= '" & cmbSearchShipTo.BoundText & "'));"

176        rsShipToData.Open qyShipToData

177        If rsShipToData.RecordCount = 0 Then
' <VB WATCH>
178        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
179            Exit Sub
180        End If

181        Set fgShipTo.DataSource = rsShipToData


182        fgShipTo.ColWidth(0) = 1400
183        fgShipTo.ColWidth(1) = 2000
184        fgShipTo.ColWidth(2) = 1200
185        fgShipTo.ColWidth(3) = 2000
186        fgShipTo.ColWidth(4) = 3200
187        fgShipTo.TextMatrix(0, 0) = "S/N"
188        fgShipTo.TextMatrix(0, 1) = "Date"
189        fgShipTo.TextMatrix(0, 2) = "Sales Order"
190        fgShipTo.TextMatrix(0, 3) = "Model No"
191        fgShipTo.TextMatrix(0, 4) = "Bill To"

192        frmTEMCFrameNo.Visible = False
193        frmCustomer.Visible = False
194        frmShipTo.Top = 8520 - (4000 - 1335)
195        frmShipTo.Height = 4000
196        fgShipTo.Height = 4000 - 360
197        frmShipTo.FontBold = True

' <VB WATCH>
198        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
199        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchShipTo_Click"

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
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchSN_Change()
' <VB WATCH>
200        On Error GoTo vbwErrHandler
201        Const VBWPROCNAME = "frmSearch.cmbSearchSN_Change"
202        If vbwProtector.vbwTraceProc Then
203            Dim vbwProtectorParameterString As String
204            If vbwProtector.vbwTraceParameters Then
205                vbwProtectorParameterString = "()"
206            End If
207            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
208        End If
' </VB WATCH>

209        Text1.Text = cmbSearchSN.BoundText

210        If rsDataSN.State = adStateOpen Then
211            rsDataSN.Close
212        End If

           'find all dates and models for the selected serial number
213        qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
           " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
           " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
           " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "'));"

214        rsDataSN.Open qyDataSN

           'if we didn't find any records, see if we have any serial numbers that are close
215        If rsDataSN.RecordCount = 0 Then
216            rsDataSN.Close
217            qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
               " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
               " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
               " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "%'));"
218            rsDataSN.Open qyDataSN
219        End If

220        If rsDataSN.RecordCount = 0 Then
' <VB WATCH>
221        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
222            Exit Sub
223        End If

224        lblNoOfPumps.Caption = rsDataSN.RecordCount & " Pumps Found"

           'bind the datalist to the other recordset
225        Set fgSN.DataSource = rsDataSN

226        fgSN.ColWidth(0) = 0               'serial number
227        fgSN.ColWidth(1) = 0               'chempumppump
228        fgSN.ColWidth(2) = 2000
229        fgSN.ColWidth(3) = 1200
230        If rsDataSN.Fields(1) = True Then       'show model number
231            fgSN.ColWidth(4) = 1800
232            fgSN.ColWidth(5) = 0
233        Else                                    'else, show TEMC Frame number
234            fgSN.ColWidth(4) = 0
235            fgSN.ColWidth(5) = 1800
236        End If
237        fgSN.ColWidth(6) = 3000
238        fgSN.ColWidth(7) = 3000
239        fgSN.TextMatrix(0, 2) = "Date"
240        fgSN.TextMatrix(0, 3) = "Sales Order"
241        fgSN.TextMatrix(0, 4) = "Model No"
242        fgSN.TextMatrix(0, 5) = "TEMC Frame"
243        fgSN.TextMatrix(0, 6) = "Bill To"
244        fgSN.TextMatrix(0, 7) = "Ship To"

           'put the serial number into the find pump textbox
245        frmPLCData.txtSN.Text = rsDataSN.Fields("SerialNumber")

246        frmModel.Visible = False
247        frmSN.Height = 4000
248        fgSN.Height = 4000 - 360
249        frmSN.FontBold = True

' <VB WATCH>
250        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
251        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchSN_Change"

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

Private Sub ORIGINALcmbSearchTEMCFrameNumber_Click(Area As Integer)
' <VB WATCH>
252        On Error GoTo vbwErrHandler
253        Const VBWPROCNAME = "frmSearch.ORIGINALcmbSearchTEMCFrameNumber_Click"
254        If vbwProtector.vbwTraceProc Then
255            Dim vbwProtectorParameterString As String
256            If vbwProtector.vbwTraceParameters Then
257                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
258            End If
259            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
260        End If
' </VB WATCH>
261        If rsDataTEMCFrameNumber.State = adStateOpen Then
262            rsDataTEMCFrameNumber.Close
263        End If

           'find all pumps with the selected temc frame number
264        qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
           " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
           " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber" & _
           " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

265        rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

266        If rsDataTEMCFrameNumber.RecordCount = 0 Then
' <VB WATCH>
267        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
268            Exit Sub
269        End If

           'bind the datalist to the recordset
270        Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

271        fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
272        fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
273        fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
274        fgTEMCFrameNo.TextMatrix(0, 3) = "Bill To"
275        fgTEMCFrameNo.TextMatrix(0, 4) = "Ship To"
276        fgTEMCFrameNo.ColWidth(0) = 1400
277        fgTEMCFrameNo.ColWidth(1) = 2000
278        fgTEMCFrameNo.ColWidth(2) = 1200
279        fgTEMCFrameNo.ColWidth(3) = 3200
280        fgTEMCFrameNo.ColWidth(4) = 3200

281        frmCustomer.Visible = False
282        frmShipTo.Visible = False
283        frmTEMCFrameNo.Height = 4000
284        fgTEMCFrameNo.Height = 4000 - 360
285        frmTEMCFrameNo.FontBold = True

' <VB WATCH>
286        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
287        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ORIGINALcmbSearchTEMCFrameNumber_Click"

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
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmbSearchTEMCFrameNumber_Click(Area As Integer)
' <VB WATCH>
288        On Error GoTo vbwErrHandler
289        Const VBWPROCNAME = "frmSearch.cmbSearchTEMCFrameNumber_Click"
290        If vbwProtector.vbwTraceProc Then
291            Dim vbwProtectorParameterString As String
292            If vbwProtector.vbwTraceParameters Then
293                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
294            End If
295            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
296        End If
' </VB WATCH>
297        If rsDataTEMCFrameNumber.State = adStateOpen Then
298            rsDataTEMCFrameNumber.Close
299        End If

           'find all pumps with the selected temc frame number
300        qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
           " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
           "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
           " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
           " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

301        rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

302        If rsDataTEMCFrameNumber.RecordCount = 0 Then
' <VB WATCH>
303        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
304            Exit Sub
305        End If

306        Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

307        Dim f As String

308        f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
309        fgTEMCFrameNo.FormatString = f
310        fgTEMCFrameNo.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
311        fgTEMCFrameNo.ColWidth(0) = 1200
312        fgTEMCFrameNo.ColWidth(1) = 2000
313        fgTEMCFrameNo.ColWidth(2) = 1200
314        fgTEMCFrameNo.ColWidth(3) = 2000
315        fgTEMCFrameNo.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
316        fgTEMCFrameNo.ColWidth(5) = 3200
317        fgTEMCFrameNo.ColWidth(6) = 3200
318        fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
319        fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
320        fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
321        fgTEMCFrameNo.TextMatrix(0, 3) = "Model No"
322        fgTEMCFrameNo.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
323        fgTEMCFrameNo.TextMatrix(0, 5) = "Bill To"
324        fgTEMCFrameNo.TextMatrix(0, 6) = "Ship To"

325        Dim X As Long
326        With fgTEMCFrameNo
327            For X = .FixedRows To .Rows - 1
328            .TextMatrix(X, 4) = Format(.TextMatrix(X, 4), "#0.000")
329            Next X
330        End With


331        frmCustomer.Visible = False
332        frmShipTo.Visible = False
333        frmTEMCFrameNo.Height = 4000
334        fgTEMCFrameNo.Height = 4000 - 360
335        frmTEMCFrameNo.FontBold = True

' <VB WATCH>
336        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
337        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchTEMCFrameNumber_Click"

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
            vbwReportVariable "Area", Area
            vbwReportVariable "f", f
            vbwReportVariable "X", X
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbStartDate_Click()
' <VB WATCH>
338        On Error GoTo vbwErrHandler
339        Const VBWPROCNAME = "frmSearch.cmbStartDate_Click"
340        If vbwProtector.vbwTraceProc Then
341            Dim vbwProtectorParameterString As String
342            If vbwProtector.vbwTraceParameters Then
343                vbwProtectorParameterString = "()"
344            End If
345            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
346        End If
' </VB WATCH>
347        Dim StartDate As Date
348        Dim EndDate As Date
349        Dim I As Integer

350        Text1.Text = cmbStartDate.List(cmbStartDate.ListIndex)
351        If rsDataDate.State = adStateOpen Then
352            rsDataDate.Close
353        End If

354        StartDate = FormatDateTime(Text1.Text)

           'see if there's and end date
355        If Left$(cmbSearchEndDate.BoundText, 1) = "S" Then
356            EndDate = StartDate
357        Else
358            EndDate = FormatDateTime(cmbSearchEndDate.BoundText)
359        End If


360        I = InStr(EndDate, " ")
361        If I <> 0 Then
362            EndDate = Left$(EndDate, I)
363        End If
364        StartDate = StartDate & " 00:00:00"
365        EndDate = EndDate & " 23:59:59"

           'look for all tests that were done on that date, regardless of the time
366        qyDataDate.CommandText = "SELECT " & _
               " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
               " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
               " WHERE ((TempTestSetupData.Date >= #" & StartDate & "#) AND (TempTestSetupData.Date <= #" & EndDate & "#)) " & _
               " ORDER BY TempTestSetupData.Date;"
367        qyData2.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]  " & _
              " FROM TempTestSetupData " & _
              " WHERE Date >= #" & StartDate & "#" & _
              " ORDER BY Date;"
       '    qyData2.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempTestSetupData.Date, TempPumpData.SalesOrderNumber, TempPumpData.ModelNumber " & _
'       " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber " & _
'       " WHERE Date >= #" & StartDate & "#" & _
'       " ORDER BY Date;"

368        If rsData2.State = adStateOpen Then
369            rsData2.Close
370        End If

371        rsData2.Open qyData2
372        Set cmbSearchEndDate.DataSource = rsData2
373        cmbSearchEndDate.ListField = "Expr2"
374        Set cmbSearchEndDate.RowSource = rsData2

375        cmbSearchEndDate.Enabled = True

376        rsDataDate.Open qyDataDate

377        If rsDataDate.RecordCount = 0 Then
' <VB WATCH>
378        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
379            Exit Sub
380        End If

381        lblNoOfPumps.Caption = rsDataDate.RecordCount & " Pumps Found"

           'bind the dldate datalist to the recordset
382        Set fgDate.DataSource = rsDataDate

383        fgDate.TextMatrix(0, 0) = "S/N"
384        fgDate.TextMatrix(0, 2) = "Date"
385        fgDate.TextMatrix(0, 3) = "Sales Order"
386        fgDate.TextMatrix(0, 4) = "Model No"
387        fgDate.TextMatrix(0, 5) = "TEMC Frame"
388        fgDate.TextMatrix(0, 6) = "Bill To"
389        fgDate.TextMatrix(0, 7) = "Ship To"
390        fgDate.ColWidth(0) = 1400               'serial number
391        fgDate.ColWidth(1) = 0               'chempumppump
392        fgDate.ColWidth(2) = 2000
393        fgDate.ColWidth(3) = 1200
394        If rsDataDate.Fields(1) = True Then       'show model number
395            fgDate.ColWidth(4) = 1800
396            fgDate.ColWidth(5) = 0
397        Else                                    'else, show TEMC Frame number
398            fgDate.ColWidth(4) = 0
399            fgDate.ColWidth(5) = 1800
400        End If
401        fgDate.ColWidth(6) = 3000
402        fgDate.ColWidth(7) = 3000

403        frmSalesOrder.Visible = False
404        frmSN.Visible = False
405        frmDate.Height = 4000
406        fgDate.Height = 4000 - 360
407        frmDate.FontBold = True

' <VB WATCH>
408        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
409        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbStartDate_Click"

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
            vbwReportVariable "StartDate", StartDate
            vbwReportVariable "EndDate", EndDate
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

Private Sub cmdClose_Click()
' <VB WATCH>
410        On Error GoTo vbwErrHandler
411        Const VBWPROCNAME = "frmSearch.cmdClose_Click"
412        If vbwProtector.vbwTraceProc Then
413            Dim vbwProtectorParameterString As String
414            If vbwProtector.vbwTraceParameters Then
415                vbwProtectorParameterString = "()"
416            End If
417            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
418        End If
' </VB WATCH>
419        Unload Me   'unload the form
' <VB WATCH>
420        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
421        Exit Sub
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

Private Sub cmdResetSizes_Click()
' <VB WATCH>
422        On Error GoTo vbwErrHandler
423        Const VBWPROCNAME = "frmSearch.cmdResetSizes_Click"
424        If vbwProtector.vbwTraceProc Then
425            Dim vbwProtectorParameterString As String
426            If vbwProtector.vbwTraceParameters Then
427                vbwProtectorParameterString = "()"
428            End If
429            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
430        End If
' </VB WATCH>
431        frmDate.Top = 0
432        frmDate.Height = 1935
433        fgDate.Height = 1575
434        frmDate.Visible = True
435        frmDate.FontBold = False

436        frmSalesOrder.Top = 1920
437        frmSalesOrder.Height = 1335
438        fgSalesOrder.Height = 975
439        frmSalesOrder.Visible = True
440        frmSalesOrder.FontBold = False

441        frmSN.Top = 3240
442        frmSN.Height = 1335
443        fgSN.Height = 975
444        frmSN.Visible = True
445        frmSN.FontBold = False

446        frmModel.Top = 4560
447        frmModel.Height = 1335
448        fgModel.Height = 975
449        frmModel.Visible = True
450        frmModel.FontBold = False

451        frmTEMCFrameNo.Top = 5880
452        frmTEMCFrameNo.Height = 1335
453        fgTEMCFrameNo.Height = 975
454        frmTEMCFrameNo.Visible = True
455        frmTEMCFrameNo.FontBold = False

456        frmCustomer.Top = 7200
457        frmCustomer.Height = 1335
458        fgCustomer.Height = 975
459        frmCustomer.Visible = True
460        frmCustomer.FontBold = False

461        frmShipTo.Top = 8520
462        frmShipTo.Height = 1335
463        fgShipTo.Height = 1335
464        frmShipTo.Visible = True
465        frmShipTo.FontBold = False

466        frmWildCard.Top = 9960
467        frmWildCard.Height = 1335
468        fgWildCard.Height = 1335
469        frmWildCard.Visible = True
470        frmWildCard.FontBold = False
471        txtModelNumberString.Text = "Enter Characters and Search with Return"

' <VB WATCH>
472        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
473        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdResetSizes_Click"

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

Private Sub fgCustomer_Click()
' <VB WATCH>
474        On Error GoTo vbwErrHandler
475        Const VBWPROCNAME = "frmSearch.fgCustomer_Click"
476        If vbwProtector.vbwTraceProc Then
477            Dim vbwProtectorParameterString As String
478            If vbwProtector.vbwTraceParameters Then
479                vbwProtectorParameterString = "()"
480            End If
481            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
482        End If
' </VB WATCH>
483        fgCustomer.Col = 0
484        frmPLCData.txtSN.Text = fgCustomer.Text
' <VB WATCH>
485        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
486        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgCustomer_Click"

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

Private Sub fgDate_Click()
' <VB WATCH>
487        On Error GoTo vbwErrHandler
488        Const VBWPROCNAME = "frmSearch.fgDate_Click"
489        If vbwProtector.vbwTraceProc Then
490            Dim vbwProtectorParameterString As String
491            If vbwProtector.vbwTraceParameters Then
492                vbwProtectorParameterString = "()"
493            End If
494            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
495        End If
' </VB WATCH>
496        fgDate.Col = 0
497        frmPLCData.txtSN.Text = fgDate.Text
' <VB WATCH>
498        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
499        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgDate_Click"

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
Private Sub fgmodel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' <VB WATCH>
500        On Error GoTo vbwErrHandler
501        Const VBWPROCNAME = "frmSearch.fgmodel_MouseDown"
502        If vbwProtector.vbwTraceProc Then
503            Dim vbwProtectorParameterString As String
504            If vbwProtector.vbwTraceParameters Then
505                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
506                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
507                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("X", X) & ", "
508                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Y", Y) & ") "
509            End If
510            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
511        End If
' </VB WATCH>

          'MSFlexGrid has the strange feature of not being able to recognize
          'when the heading of 1st fixed column is clicked on, it calls it Row 1,
          'the Row below this also returns as Row 1.
          'This bit of code below singles out the heading row which is required in
          'this app for sorting the data.

512        Static LastCol As Integer       'last column clicked on
513        Static Direction As Boolean     'sorting ascending or descending

514        If Y < fgModel.RowHeight(fgModel.Row) Then  'if user clicked in header row
515            fgModel.Row = 1
516            fgModel.RowSel = 1
517            fgModel.ColSel = fgModel.Col
518            If LastCol = fgModel.Col Then   'if user clicked on same column, reverse sort
519                Direction = Not Direction
520            Else
521                Direction = True            'if new column, sort ascending
522            End If

523            If Direction Then
524                fgModel.Sort = flexSortGenericAscending
525            Else
526                fgModel.Sort = flexSortGenericDescending
527            End If

528            LastCol = fgModel.Col   'save column number
529        Else                            'user did not click on header, select serial number for main screen
530            fgModel.Col = 0
531            frmPLCData.txtSN.Text = fgModel.Text
532        End If

' <VB WATCH>
533        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
534        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgmodel_MouseDown"

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
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "X", X
            vbwReportVariable "Y", Y
            vbwReportVariable "LastCol", LastCol
            vbwReportVariable "Direction", Direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgSalesOrder_Click()
' <VB WATCH>
535        On Error GoTo vbwErrHandler
536        Const VBWPROCNAME = "frmSearch.fgSalesOrder_Click"
537        If vbwProtector.vbwTraceProc Then
538            Dim vbwProtectorParameterString As String
539            If vbwProtector.vbwTraceParameters Then
540                vbwProtectorParameterString = "()"
541            End If
542            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
543        End If
' </VB WATCH>
544        fgSalesOrder.Col = 0
545        frmPLCData.txtSN.Text = fgSalesOrder.Text
' <VB WATCH>
546        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
547        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgSalesOrder_Click"

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

Private Sub fgshipto_Click()
' <VB WATCH>
548        On Error GoTo vbwErrHandler
549        Const VBWPROCNAME = "frmSearch.fgshipto_Click"
550        If vbwProtector.vbwTraceProc Then
551            Dim vbwProtectorParameterString As String
552            If vbwProtector.vbwTraceParameters Then
553                vbwProtectorParameterString = "()"
554            End If
555            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
556        End If
' </VB WATCH>
557        fgShipTo.Col = 0
558        frmPLCData.txtSN.Text = fgShipTo.Text
' <VB WATCH>
559        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
560        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgshipto_Click"

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

Private Sub fgSN_Click()
' <VB WATCH>
561        On Error GoTo vbwErrHandler
562        Const VBWPROCNAME = "frmSearch.fgSN_Click"
563        If vbwProtector.vbwTraceProc Then
564            Dim vbwProtectorParameterString As String
565            If vbwProtector.vbwTraceParameters Then
566                vbwProtectorParameterString = "()"
567            End If
568            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
569        End If
' </VB WATCH>
570        fgSN.Col = 0
571        frmPLCData.txtSN.Text = fgSN.Text
' <VB WATCH>
572        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
573        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgSN_Click"

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

Private Sub fgTEMCFrameNo_Click()
' <VB WATCH>
574        On Error GoTo vbwErrHandler
575        Const VBWPROCNAME = "frmSearch.fgTEMCFrameNo_Click"
576        If vbwProtector.vbwTraceProc Then
577            Dim vbwProtectorParameterString As String
578            If vbwProtector.vbwTraceParameters Then
579                vbwProtectorParameterString = "()"
580            End If
581            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
582        End If
' </VB WATCH>
583        fgTEMCFrameNo.Col = 0
584        frmPLCData.txtSN.Text = fgTEMCFrameNo.Text
' <VB WATCH>
585        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
586        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgTEMCFrameNo_Click"

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
587        On Error GoTo vbwErrHandler
588        Const VBWPROCNAME = "frmSearch.Form_Activate"
589        If vbwProtector.vbwTraceProc Then
590            Dim vbwProtectorParameterString As String
591            If vbwProtector.vbwTraceParameters Then
592                vbwProtectorParameterString = "()"
593            End If
594            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
595        End If
' </VB WATCH>
596        Const HWND_TOPMOST As Integer = -1
597        Const SWP_NOSIZE As Integer = &H1
598        Const SWP_NOMOVE As Integer = &H2
599        Const SWP_NOACTIVATE As Integer = &H10
600        Const SWP_SHOWWINDOW As Integer = &H40

           'window always on top
       '    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

' <VB WATCH>
601        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
602        Exit Sub
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

Private Sub Form_Load()
           'open several recordsets for searching
' <VB WATCH>
603        On Error GoTo vbwErrHandler
604        Const VBWPROCNAME = "frmSearch.Form_Load"
605        If vbwProtector.vbwTraceProc Then
606            Dim vbwProtectorParameterString As String
607            If vbwProtector.vbwTraceParameters Then
608                vbwProtectorParameterString = "()"
609            End If
610            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
611        End If
' </VB WATCH>

           'qydata/rsdata is for the serial number dropdown
612        qyData.ActiveConnection = cnPumpData
613        qyData.CommandText = "SELECT DISTINCT SerialNumber FROM TempPumpData ORDER BY SerialNumber;"
614        rsData.CursorType = adOpenStatic
615        rsData.CursorLocation = adUseClient
616        rsData.Index = "SerialNumber"
617        rsData.Open qyData

           'bind the serial number dropdown
618        Set cmbSearchSN.DataSource = rsData
619        cmbSearchSN.ListField = "SerialNumber"
620        Set cmbSearchSN.RowSource = rsData

           'qydata1 and 2/rsdata1 and 2 is for the date dropdown
621        qyData1.ActiveConnection = cnPumpData
622        rsData1.CursorType = adOpenStatic
623        rsData1.CursorLocation = adUseClient
624        rsData1.Index = "SerialNumber"

625        qyData2.ActiveConnection = cnPumpData
626        rsData2.CursorType = adOpenStatic
627        rsData2.CursorLocation = adUseClient
628        rsData2.Index = "SerialNumber"

           'find dates without times
       '    qyData1.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempPumpData.Model, TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]" & _
'        " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber ORDER BY Date;"
629        qyData1.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2] " & _
               " FROM TempTestSetupData ORDER BY Date;"
630        rsData1.Open qyData1

631        Dim I As Integer
632        Dim TempDate As Date
633        Dim LastDate As Date
634        rsData1.MoveFirst

635        LastDate = FormatDateTime(Now(), vbShortDate)
636        For I = 1 To rsData1.RecordCount
637            TempDate = FormatDateTime(rsData1.Fields(0), vbShortDate)
638            If TempDate <> LastDate Then
639                cmbStartDate.AddItem TempDate
640                LastDate = TempDate
641            End If
642            rsData1.MoveNext
643        Next I

           'qydatadate/rsdatadate for date datalist
644        qyDataDate.ActiveConnection = cnPumpData
645        rsDataDate.CursorType = adOpenStatic
646        rsDataDate.CursorLocation = adUseClient

           'qydatamodel/rsdatamodel for model dropdown
647        qyDataModel.ActiveConnection = cnPumpData
648        rsDataModel.CursorType = adOpenStatic
649        rsDataModel.CursorLocation = adUseClient

           'qydatasalesorder/rsdatasalesorder for sales order dropdown
650        qyDataSalesOrder.ActiveConnection = cnPumpData
651        rsDataSalesOrder.CursorType = adOpenStatic
652        rsDataSalesOrder.CursorLocation = adUseClient
653        qyDataSalesOrder.CommandText = "SELECT DISTINCT TempPumpData.SalesOrderNumber FROM TempPumpData ORDER BY TempPumpData.SalesOrderNumber;"
654        rsDataSalesOrder.Open qyDataSalesOrder

655        qySalesOrderData.ActiveConnection = cnPumpData
656        rsSalesOrderData.CursorType = adOpenStatic
657        rsSalesOrderData.CursorLocation = adUseClient

           'bind to temc frame number dropdown
658        Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder
659        cmbSearchSalesOrder.ListField = "SalesOrderNumber"
660        Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder

           'qydatasn/rsdatasn for serial numbers
661        qyDataSN.ActiveConnection = cnPumpData
662        rsDataSN.CursorType = adOpenStatic
663        rsDataSN.CursorLocation = adUseClient

           'qydatatemcmodel/rsdatatemcmodel for temc frame number
664        qyDataTEMCModel.ActiveConnection = cnPumpData
665        rsDataTEMCModel.CursorType = adOpenStatic
666        rsDataTEMCModel.CursorLocation = adUseClient
667        qyDataTEMCModel.CommandText = "SELECT DISTINCT TempPumpData.TEMCFrameNumber FROM TempPumpData WHERE (TempPumpData.ChempumpPump = FALSE) ORDER BY TempPumpData.TEMCFrameNumber;"
668        rsDataTEMCModel.Open qyDataTEMCModel

           'bind to temc frame number dropdown
669        Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel
670        cmbSearchTEMCFrameNumber.ListField = "TEMCFrameNumber"
671        Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel

672        qyDataTEMCFrameNumber.ActiveConnection = cnPumpData
673        rsDataTEMCFrameNumber.CursorType = adOpenStatic
674        rsDataTEMCFrameNumber.CursorLocation = adUseClient

           'customer
675        qyDataCustomer.ActiveConnection = cnPumpData
676        rsDataCustomer.CursorType = adOpenStatic
677        rsDataCustomer.CursorLocation = adUseClient

678        qyDataCustomer.CommandText = "SELECT DISTINCT TempPumpData.BillToCustomer FROM TempPumpData ORDER BY TempPumpData.BillToCustomer;"
679        rsDataCustomer.Open qyDataCustomer

680        qyCustomerData.ActiveConnection = cnPumpData
681        rsCustomerData.CursorType = adOpenStatic
682        rsCustomerData.CursorLocation = adUseClient

           'bind to customer dropdown
683        Set cmbSearchCustomer.RowSource = rsDataCustomer
684        cmbSearchCustomer.ListField = "BillToCustomer"

           ' ship to customer
685        qyDataShipto.ActiveConnection = cnPumpData
686        rsDataShipTo.CursorType = adOpenStatic
687        rsDataShipTo.CursorLocation = adUseClient

688        qyDataShipto.CommandText = "SELECT DISTINCT TempPumpData.shipToCustomer FROM TempPumpData ORDER BY TempPumpData.shipToCustomer;"
689        rsDataShipTo.Open qyDataShipto

690        qyShipToData.ActiveConnection = cnPumpData
691        rsShipToData.CursorType = adOpenStatic
692        rsShipToData.CursorLocation = adUseClient

           'bind to customer dropdown
693        Set cmbSearchShipTo.RowSource = rsDataShipTo
694        cmbSearchShipTo.ListField = "ShipToCustomer"

695        cmbSearchEndDate.Enabled = False

' <VB WATCH>
696        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
697        Exit Sub
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
            vbwReportVariable "I", I
            vbwReportVariable "TempDate", TempDate
            vbwReportVariable "LastDate", LastDate
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
698        On Error GoTo vbwErrHandler
699        Const VBWPROCNAME = "frmSearch.Form_Unload"
700        If vbwProtector.vbwTraceProc Then
701            Dim vbwProtectorParameterString As String
702            If vbwProtector.vbwTraceParameters Then
703                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
704            End If
705            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
706        End If
' </VB WATCH>

           'close all the datasets and release the connections
707        If rsData.State = adStateOpen Then
708            rsData.Close
709        End If
710        If rsData1.State = adStateOpen Then
711            rsData1.Close
712        End If
713        If rsData2.State = adStateOpen Then
714            rsData2.Close
715        End If
716        If rsDataDate.State = adStateOpen Then
717            rsDataDate.Close
718        End If
719        If rsDataModel.State = adStateOpen Then
720            rsDataModel.Close
721        End If
722        If rsDataSalesOrder.State = adStateOpen Then
723            rsDataSalesOrder.Close
724        End If
725        If rsSalesOrderData.State = adStateOpen Then
726            rsSalesOrderData.Close
727        End If
728        If rsDataSN.State = adStateOpen Then
729            rsDataSN.Close
730        End If
731        If rsDataTEMCModel.State = adStateOpen Then
732            rsDataTEMCModel.Close
733        End If
734        If rsDataTEMCFrameNumber.State = adStateOpen Then
735            rsDataTEMCFrameNumber.Close
736        End If
737        If rsDataCustomer.State = adStateOpen Then
738            rsDataCustomer.Close
739        End If
740        If rsCustomerData.State = adStateOpen Then
741            rsCustomerData.Close
742        End If
743        If rsDataShipTo.State = adStateOpen Then
744            rsDataShipTo.Close
745        End If
746        If rsShipToData.State = adStateOpen Then
747            rsShipToData.Close
748        End If

749        Set rsData = Nothing
750        Set rsData1 = Nothing
751        Set rsData2 = Nothing
752        Set rsDataDate = Nothing
753        Set rsDataModel = Nothing
754        Set rsDataSalesOrder = Nothing
755        Set rsSalesOrderData = Nothing
756        Set rsDataSN = Nothing
757        Set rsDataTEMCModel = Nothing
758        Set rsDataTEMCFrameNumber = Nothing
759        Set rsDataCustomer = Nothing
760        Set rsCustomerData = Nothing
761        Set rsDataShipTo = Nothing
762        Set rsShipToData = Nothing

' <VB WATCH>
763        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
764        Exit Sub
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

Private Sub WildCardSearch()
' <VB WATCH>
765        On Error GoTo vbwErrHandler
766        Const VBWPROCNAME = "frmSearch.WildCardSearch"
767        If vbwProtector.vbwTraceProc Then
768            Dim vbwProtectorParameterString As String
769            If vbwProtector.vbwTraceParameters Then
770                vbwProtectorParameterString = "()"
771            End If
772            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
773        End If
' </VB WATCH>
774        If rsDataModel.State = adStateOpen Then
775            rsDataModel.Close
776        End If

777        qyDataModel.CommandText = "SELECT DISTINCT " & _
           " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
           "  TempPumpData.BillToCustomer,  Model.Description " & _
           " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
           " WHERE (((TempPumpData.ModelNumber) LIKE '%" & txtModelNumberString.Text & "%'));"


778        rsDataModel.Open qyDataModel

779        If rsDataModel.RecordCount = 0 Then
' <VB WATCH>
780        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
781            Exit Sub
782        End If

783        Set fgWildCard.DataSource = rsDataModel

784        Dim f As String

785        f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
786        fgWildCard.FormatString = f
787        fgWildCard.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
788        fgWildCard.ColWidth(0) = 1200
789        fgWildCard.ColWidth(1) = 2000
790        fgWildCard.ColWidth(2) = 1200
791        fgWildCard.ColWidth(3) = 2000
792        fgWildCard.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
793        fgWildCard.ColWidth(5) = 3200
794        fgWildCard.ColWidth(6) = 3200
795        fgWildCard.TextMatrix(0, 0) = "S/N"
796        fgWildCard.TextMatrix(0, 1) = "Date"
797        fgWildCard.TextMatrix(0, 2) = "Sales Order"
798        fgWildCard.TextMatrix(0, 3) = "Model No"
799        fgWildCard.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
800        fgWildCard.TextMatrix(0, 5) = "Bill To"
801        fgWildCard.TextMatrix(0, 6) = "Ship To"

802        Dim X As Long
803        With fgWildCard
804            For X = .FixedRows To .Rows - 1
805            .TextMatrix(X, 4) = Format(.TextMatrix(X, 4), "#0.000")
806            Next X
807        End With


808        frmTEMCFrameNo.Visible = False
809        frmCustomer.Visible = False
810        frmWildCard.Top = 4560
811        frmWildCard.Height = 4000
812        fgWildCard.Height = 4000 - 360
813        frmWildCard.FontBold = True

' <VB WATCH>
814        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
815        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WildCardSearch"

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
            vbwReportVariable "f", f
            vbwReportVariable "X", X
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub txtModelNumberString_KeyPress(KeyAscii As Integer)
' <VB WATCH>
816        On Error GoTo vbwErrHandler
817        Const VBWPROCNAME = "frmSearch.txtModelNumberString_KeyPress"
818        If vbwProtector.vbwTraceProc Then
819            Dim vbwProtectorParameterString As String
820            If vbwProtector.vbwTraceParameters Then
821                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("KeyAscii", KeyAscii) & ") "
822            End If
823            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
824        End If
' </VB WATCH>
825        If KeyAscii = 13 Then
826            WildCardSearch
827        End If
' <VB WATCH>
828        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
829        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNumberString_KeyPress"

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
            vbwReportVariable "KeyAscii", KeyAscii
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgwildcard_Click()
' <VB WATCH>
830        On Error GoTo vbwErrHandler
831        Const VBWPROCNAME = "frmSearch.fgwildcard_Click"
832        If vbwProtector.vbwTraceProc Then
833            Dim vbwProtectorParameterString As String
834            If vbwProtector.vbwTraceParameters Then
835                vbwProtectorParameterString = "()"
836            End If
837            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
838        End If
' </VB WATCH>
839        fgWildCard.Col = 0
840        frmPLCData.txtSN.Text = fgWildCard.Text
' <VB WATCH>
841        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
842        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgwildcard_Click"

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

Private Sub txtModelNumberString_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' <VB WATCH>
843        On Error GoTo vbwErrHandler
844        Const VBWPROCNAME = "frmSearch.txtModelNumberString_MouseDown"
845        If vbwProtector.vbwTraceProc Then
846            Dim vbwProtectorParameterString As String
847            If vbwProtector.vbwTraceParameters Then
848                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
849                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
850                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("X", X) & ", "
851                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Y", Y) & ") "
852            End If
853            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
854        End If
' </VB WATCH>
855        If txtModelNumberString.Text = "Enter Characters and Search with Return" Then
856            txtModelNumberString.Text = ""
857        End If

' <VB WATCH>
858        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
859        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNumberString_MouseDown"

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
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "X", X
            vbwReportVariable "Y", Y
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
    vbwReportVariable "qyData2", qyData2
    vbwReportVariable "rsData2", rsData2
    vbwReportVariable "rsDataDate", rsDataDate
    vbwReportVariable "qyDataDate", qyDataDate
    vbwReportVariable "rsDataModel", rsDataModel
    vbwReportVariable "qyDataModel", qyDataModel
    vbwReportVariable "rsDataSN", rsDataSN
    vbwReportVariable "qyDataSN", qyDataSN
    vbwReportVariable "rsDataSalesOrder", rsDataSalesOrder
    vbwReportVariable "qySalesOrderData", qySalesOrderData
    vbwReportVariable "rsSalesOrderData", rsSalesOrderData
    vbwReportVariable "qyDataSalesOrder", qyDataSalesOrder
    vbwReportVariable "rsDataTEMCModel", rsDataTEMCModel
    vbwReportVariable "qyDataTEMCModel", qyDataTEMCModel
    vbwReportVariable "rsDataTEMCFrameNumber", rsDataTEMCFrameNumber
    vbwReportVariable "qyDataTEMCFrameNumber", qyDataTEMCFrameNumber
    vbwReportVariable "rsDataCustomer", rsDataCustomer
    vbwReportVariable "qyDataCustomer", qyDataCustomer
    vbwReportVariable "rsCustomerData", rsCustomerData
    vbwReportVariable "qyCustomerData", qyCustomerData
    vbwReportVariable "rsDataShipTo", rsDataShipTo
    vbwReportVariable "qyDataShipto", qyDataShipto
    vbwReportVariable "rsShipToData", rsShipToData
    vbwReportVariable "qyShipToData", qyShipToData
End Sub
' </VB WATCH>
