VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPLCData 
   Caption         =   "v"
   ClientHeight    =   11760
   ClientLeft      =   4110
   ClientTop       =   1650
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   11760
   ScaleWidth      =   15375
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearchForPump 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search for Pump"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   283
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalibrate 
      Caption         =   "Calibrate Software"
      Height          =   495
      Left            =   11280
      TabIndex        =   208
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbTestDate 
      Height          =   315
      Left            =   6720
      TabIndex        =   48
      Top             =   120
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc adohp 
      Height          =   330
      Left            =   11160
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=True;Data Source=HP-3000/32;Mode=Read"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=True;Data Source=HP-3000/32;Mode=Read"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFindPump 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Find Pump"
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtSN 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer tmrStartUp 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10440
      Top             =   0
   End
   Begin VB.TextBox txtUpdateInterval 
      Height          =   495
      Left            =   8760
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrGetDDE 
      Interval        =   2000
      Left            =   9960
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10440
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   15012
      _ExtentX        =   26485
      _ExtentY        =   18415
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pump Data"
      TabPicture(0)   =   "Main.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTab1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTab1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTab1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTab1(11)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTab1(12)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTab1(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTab1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblTab1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTab1(44)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmChempump"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBilNo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtShpNo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtModelNo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDesignFlow"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDesignTDH"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtRemarks"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdEnterPumpData"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtSalesOrderNumber"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdDeletePump"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdApprovePump"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "frmMfr"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdClearPumpData"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtImpellerDia"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "frmMiscPumpData"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtLineNumber"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "CommonDialog2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "frmTEMC"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Test Setup"
      TabPicture(1)   =   "Main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRMA"
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(2)=   "frmTAndI"
      Tab(1).Control(3)=   "cmdApproveTestDate"
      Tab(1).Control(4)=   "cmdDeleteTestDate"
      Tab(1).Control(5)=   "CommonDialog1"
      Tab(1).Control(6)=   "frmOtherFiles"
      Tab(1).Control(7)=   "frmPerfMods"
      Tab(1).Control(8)=   "frmThrustBalMods"
      Tab(1).Control(9)=   "frmElecData"
      Tab(1).Control(10)=   "frmLoopAndXducer"
      Tab(1).Control(11)=   "frmInstrumentTags"
      Tab(1).Control(12)=   "txtTestSetupRemarks"
      Tab(1).Control(13)=   "cmdAddNewTestDate"
      Tab(1).Control(14)=   "txtWho"
      Tab(1).Control(15)=   "cmdEnterTestSetupData"
      Tab(1).Control(16)=   "cmbTestSpec"
      Tab(1).Control(17)=   "lbltab2(88)"
      Tab(1).Control(18)=   "lbltab2(65)"
      Tab(1).Control(19)=   "lbltab2(1)"
      Tab(1).Control(20)=   "lbltab2(0)"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Test Data"
      TabPicture(2)   =   "Main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSChart2"
      Tab(2).Control(1)=   "frmReport"
      Tab(2).Control(2)=   "MSChart1"
      Tab(2).Control(3)=   "txtUpDn2"
      Tab(2).Control(4)=   "txtUpDn1"
      Tab(2).Control(5)=   "UpDown2"
      Tab(2).Control(6)=   "UpDown1"
      Tab(2).Control(7)=   "frmMagtrol"
      Tab(2).Control(8)=   "frmPLCMisc"
      Tab(2).Control(9)=   "Command2"
      Tab(2).Control(10)=   "txtTDH"
      Tab(2).Control(11)=   "cmdReport"
      Tab(2).Control(12)=   "txtNPSHa"
      Tab(2).Control(13)=   "frmPumpData"
      Tab(2).Control(14)=   "DataGrid2"
      Tab(2).Control(15)=   "cmdEnterTestData"
      Tab(2).Control(16)=   "fmrMiscTestData"
      Tab(2).Control(17)=   "frmThermocouples"
      Tab(2).Control(18)=   "frmAI"
      Tab(2).Control(19)=   "cmbPLCLoop"
      Tab(2).Control(20)=   "DataGrid1"
      Tab(2).Control(21)=   "shpGetPLCData"
      Tab(2).Control(22)=   "lbltab2(54)"
      Tab(2).Control(23)=   "lbltab2(53)"
      Tab(2).Control(24)=   "lbltab2(64)"
      Tab(2).Control(25)=   "lbltab2(63)"
      Tab(2).Control(26)=   "Line2"
      Tab(2).Control(27)=   "Line1"
      Tab(2).Control(28)=   "lbltab2(59)"
      Tab(2).Control(29)=   "lbltab2(58)"
      Tab(2).Control(30)=   "lbltab2(55)"
      Tab(2).ControlCount=   31
      TabCaption(3)   =   "Charts"
      TabPicture(3)   =   "Main.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSChart6"
      Tab(3).Control(1)=   "MSChart5"
      Tab(3).Control(2)=   "MSChart4"
      Tab(3).Control(3)=   "MSChart3"
      Tab(3).ControlCount=   4
      Begin VB.Frame frmTEMC 
         Caption         =   "TEMC Pump Data"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         TabIndex        =   209
         Top             =   4680
         Visible         =   0   'False
         Width           =   14535
         Begin VB.ComboBox cmbTEMCNominalSuctionSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   261
            Top             =   600
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCNominalDischargeSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   260
            Top             =   240
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCVoltage 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   259
            Top             =   2400
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCDesignPressure 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   256
            Top             =   1320
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCCirculation 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   255
            Top             =   3120
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCModel 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   253
            Top             =   240
            Width           =   5445
         End
         Begin VB.TextBox txtTEMCFrameNumber 
            Height          =   315
            Left            =   1680
            TabIndex        =   233
            Top             =   1680
            Width           =   855
         End
         Begin VB.ComboBox cmbTEMCTRG 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   230
            Top             =   3120
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCPumpStages 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   228
            Top             =   2040
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCOtherMotor 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   226
            Top             =   2760
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCNominalImpSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   224
            Top             =   960
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCMaterials 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   222
            Top             =   960
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCJacketGasket 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   220
            Top             =   2400
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCInsulation 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   218
            Top             =   2040
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCImpellerType 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   216
            Top             =   1320
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCDivisionType 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   214
            Top             =   1680
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCAdditions 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   212
            Top             =   2760
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCAdapter 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   210
            Top             =   600
            Width           =   5445
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Suction Size:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   7200
            TabIndex        =   264
            Top             =   626
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   263
            Top             =   2433
            Width           =   1095
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Discharge Size:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   6840
            TabIndex        =   262
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Design Pressure:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   0
            TabIndex        =   258
            Top             =   1359
            Width           =   1575
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Circulation:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   7200
            TabIndex        =   257
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Type:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   360
            TabIndex        =   254
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Frame No:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   480
            TabIndex        =   232
            Top             =   1717
            Width           =   1095
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TRG:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   480
            TabIndex        =   231
            Top             =   3150
            Width           =   1095
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Stages:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   7200
            TabIndex        =   229
            Top             =   2050
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   480
            TabIndex        =   227
            Top             =   2791
            Width           =   1095
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nom Imp Size:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   7200
            TabIndex        =   225
            Top             =   982
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Materials:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   600
            TabIndex        =   223
            Top             =   1001
            Width           =   975
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jacket/Gasket:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   7200
            TabIndex        =   221
            Top             =   2406
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Insulation:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   219
            Top             =   2075
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impeller Type:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   7080
            TabIndex        =   217
            Top             =   1338
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Division Type:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   7200
            TabIndex        =   215
            Top             =   1694
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Additions:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   7560
            TabIndex        =   213
            Top             =   2762
            Width           =   855
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Adapter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   720
            TabIndex        =   211
            Top             =   643
            Width           =   855
         End
      End
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   1455
         Left            =   -62280
         OleObjectBlob   =   "Main.frx":0070
         TabIndex        =   421
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame frmReport 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Select Report"
         Height          =   4695
         Left            =   -66720
         TabIndex        =   186
         Top             =   2760
         Visible         =   0   'False
         Width           =   5175
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "TEMC Inspection Report"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   318
            Top             =   3600
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Customer Report with Vibration without Axial Position"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   298
            Top             =   1800
            Visible         =   0   'False
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Customer Report without Circulation Flow without Axial Position"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   297
            Top             =   720
            Visible         =   0   'False
            Width           =   4815
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Export Data to Excel"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   204
            Top             =   3240
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Unapproved Data"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   201
            Top             =   3960
            Visible         =   0   'False
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Charts"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   194
            Top             =   2880
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Cancel Report Request"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   192
            Top             =   4320
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Balance Holes"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   191
            Top             =   2520
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Internal"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   190
            Top             =   2160
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Customer Report with Vibration (No Circ Flow)"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   189
            Top             =   1440
            Visible         =   0   'False
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Customer Report with Options"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   188
            Top             =   1080
            Width           =   4575
         End
         Begin VB.OptionButton optReport 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Customer Report without Circulation Flow"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   187
            Top             =   360
            Visible         =   0   'False
            Width           =   4575
         End
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3012
         Left            =   -68040
         OleObjectBlob   =   "Main.frx":1D02
         TabIndex        =   420
         Top             =   2880
         Width           =   5892
      End
      Begin VB.TextBox txtUpDn2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -60840
         TabIndex        =   419
         Text            =   "8"
         Top             =   5520
         Width           =   372
      End
      Begin VB.TextBox txtUpDn1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   -74280
         TabIndex        =   418
         Text            =   "1"
         Top             =   8760
         Width           =   480
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   432
         Left            =   -61080
         TabIndex        =   417
         Top             =   5520
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   767
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUpDn2"
         BuddyDispid     =   196620
         OrigLeft        =   14040
         OrigTop         =   7560
         OrigRight       =   14292
         OrigBottom      =   7932
         Max             =   8
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   492
         Left            =   -74520
         TabIndex        =   416
         Top             =   8760
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUpDn1"
         BuddyDispid     =   196621
         OrigLeft        =   480
         OrigTop         =   8760
         OrigRight       =   732
         OrigBottom      =   9252
         Max             =   8
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Frame frmMagtrol 
         BackColor       =   &H8000000A&
         Caption         =   "Magtrol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   71
         Top             =   3960
         Width           =   6735
         Begin VB.OptionButton optKW 
            Caption         =   "Use Ana In 4"
            Height          =   195
            Index           =   2
            Left            =   5280
            TabIndex        =   296
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton optKW 
            Caption         =   "Enter KW"
            Height          =   195
            Index           =   1
            Left            =   5280
            TabIndex        =   295
            Top             =   1440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton optKW 
            Caption         =   "Add 3 powers"
            Height          =   195
            Index           =   0
            Left            =   5280
            TabIndex        =   294
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdFindMagtrols 
            Caption         =   "Find Magtrols"
            Height          =   255
            Left            =   5040
            TabIndex        =   207
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cmbMagtrol 
            Height          =   315
            Left            =   2520
            TabIndex        =   205
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtV1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   82
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtV2 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   81
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtV3 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   80
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtI1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   79
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtI2 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   78
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtI3 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   77
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtP1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   76
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtP2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   75
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtP3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   74
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtPF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4320
            TabIndex        =   73
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtKW 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5520
            TabIndex        =   72
            Top             =   840
            Width           =   855
         End
         Begin VB.Shape shpGetMagtrolData 
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   350
            Left            =   120
            Shape           =   3  'Circle
            Top             =   240
            Width           =   492
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Magtrol Select"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   600
            TabIndex        =   206
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Voltage"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   1080
            TabIndex        =   90
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   2160
            TabIndex        =   89
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Power"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   3240
            TabIndex        =   88
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Phase 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   240
            TabIndex        =   87
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Phase 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   240
            TabIndex        =   86
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Phase 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   240
            TabIndex        =   85
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "PF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   4440
            TabIndex        =   84
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Total KW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   5520
            TabIndex        =   83
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame frmPLCMisc 
         Caption         =   "PLC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -68040
         TabIndex        =   120
         Top             =   1020
         Width           =   4095
         Begin VB.TextBox txtManualLamp 
            Height          =   285
            Left            =   2520
            TabIndex        =   133
            Text            =   "Text1"
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtWriteSP 
            Height          =   375
            Left            =   2880
            TabIndex        =   132
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtWriteSPData 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   131
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSetPointDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   130
            Top             =   330
            Width           =   615
         End
         Begin VB.TextBox txtValvePositionDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   129
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtValvePosition 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3120
            TabIndex        =   128
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDCoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2400
            TabIndex        =   127
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtICoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3000
            TabIndex        =   126
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtPCoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2160
            TabIndex        =   125
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSetPoint 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   124
            Text            =   "0"
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton cmdWriteSP 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Write SP"
            Height          =   405
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtInHgDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   122
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtInHg 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            TabIndex        =   121
            Top             =   1200
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Valve Position"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   44
            Left            =   2280
            TabIndex        =   137
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Set Point"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   240
            TabIndex        =   136
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "SP to Write"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   43
            Left            =   120
            TabIndex        =   135
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "In Hg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   45
            Left            =   2160
            TabIndex        =   134
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Send Hydraulic Test Report to Excel"
         Height          =   855
         Left            =   -63720
         TabIndex        =   411
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtRMA 
         Height          =   315
         Left            =   -69960
         TabIndex        =   407
         Top             =   540
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   720
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtLineNumber 
         Height          =   315
         Left            =   1920
         TabIndex        =   399
         Top             =   840
         Width           =   615
      End
      Begin VB.Frame frmMiscPumpData 
         Caption         =   "Pump Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   373
         Top             =   2160
         Width           =   14535
         Begin VB.Frame Frame1 
            Caption         =   "NPSH Data File Directory"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8040
            TabIndex        =   405
            Top             =   1200
            Width           =   6375
            Begin VB.TextBox txtNPSHFileLocation 
               Height          =   315
               Left            =   120
               TabIndex        =   406
               Top             =   240
               Width           =   5895
            End
         End
         Begin VB.TextBox txtLiquid 
            Height          =   315
            Left            =   2400
            TabIndex        =   393
            Top             =   1440
            Width           =   5415
         End
         Begin VB.TextBox txtJobNum 
            Height          =   315
            Left            =   12840
            TabIndex        =   392
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtSpGr 
            Height          =   315
            Left            =   5880
            TabIndex        =   388
            Top             =   990
            Width           =   1335
         End
         Begin VB.TextBox txtRatedInputPower 
            Height          =   315
            Left            =   2400
            TabIndex        =   387
            Top             =   1020
            Width           =   1335
         End
         Begin VB.TextBox txtLiquidTemperature 
            Height          =   315
            Left            =   12840
            TabIndex        =   386
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNPSHr 
            Height          =   315
            Left            =   2400
            TabIndex        =   382
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox txtThermalClass 
            Height          =   315
            Left            =   5880
            TabIndex        =   381
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox txtExpClass 
            Height          =   315
            Left            =   9120
            TabIndex        =   380
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtAmps 
            Height          =   315
            Left            =   5880
            TabIndex        =   377
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtViscosity 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   9120
            TabIndex        =   376
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNoPhases 
            Height          =   315
            Left            =   2400
            TabIndex        =   374
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   960
            TabIndex        =   395
            Top             =   1470
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Job Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   43
            Left            =   11400
            TabIndex        =   394
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Specific Gravity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   4200
            TabIndex        =   391
            Top             =   1050
            Width           =   1575
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rated Input Power:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   36
            Left            =   360
            TabIndex        =   390
            Top             =   1056
            Width           =   1932
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Temperature:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   10800
            TabIndex        =   389
            Top             =   330
            Width           =   1935
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "NPSHr:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   960
            TabIndex        =   385
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Thermal Class:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   4440
            TabIndex        =   384
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "EXP Class:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   7680
            TabIndex        =   383
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Amps:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   4440
            TabIndex        =   379
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Viscosity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   7680
            TabIndex        =   378
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Phases:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   1440
            TabIndex        =   375
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -67320
         TabIndex        =   372
         Top             =   6600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame frmTAndI 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Test and Inspection Report Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -66480
         TabIndex        =   319
         Top             =   3600
         Width           =   6135
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   3720
            TabIndex        =   371
            Top             =   5160
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   5640
            TabIndex        =   367
            Top             =   4800
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   5640
            TabIndex        =   366
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   5640
            TabIndex        =   365
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   364
            Top             =   3720
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   5640
            TabIndex        =   363
            Top             =   3360
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2760
            TabIndex        =   362
            Top             =   4800
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2760
            TabIndex        =   361
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   2760
            TabIndex        =   360
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2760
            TabIndex        =   359
            Top             =   3720
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2760
            TabIndex        =   358
            Top             =   3360
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   5580
            TabIndex        =   346
            Top             =   2490
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   5580
            TabIndex        =   345
            Top             =   1890
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   5580
            TabIndex        =   344
            Top             =   1290
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   5580
            TabIndex        =   343
            Top             =   690
            Width           =   255
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   840
            TabIndex        =   338
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   3120
            TabIndex        =   337
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   840
            TabIndex        =   336
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   3120
            TabIndex        =   335
            Top             =   1920
            Width           =   735
         End
         Begin VB.ComboBox cmbTestAndInspection 
            Height          =   315
            Index           =   1
            ItemData        =   "Main.frx":406E
            Left            =   1680
            List            =   "Main.frx":4078
            TabIndex        =   334
            Top             =   2520
            Width           =   975
         End
         Begin VB.ComboBox cmbTestAndInspection 
            Height          =   315
            Index           =   0
            ItemData        =   "Main.frx":4088
            Left            =   1680
            List            =   "Main.frx":4092
            TabIndex        =   333
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   3120
            TabIndex        =   330
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   840
            TabIndex        =   328
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   326
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   320
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Supervisor Approval?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   85
            Left            =   1680
            TabIndex        =   370
            Top             =   5220
            Width           =   1935
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   6000
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Good?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   84
            Left            =   5400
            TabIndex        =   369
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Good?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   83
            Left            =   2580
            TabIndex        =   368
            Top             =   3120
            Width           =   612
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Nameplate Check?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   82
            Left            =   3600
            TabIndex        =   357
            Top             =   4860
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Paint Check?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   81
            Left            =   3360
            TabIndex        =   356
            Top             =   4500
            Width           =   2175
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Clean, Purge and Seal?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   80
            Left            =   3360
            TabIndex        =   355
            Top             =   4140
            Width           =   2172
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "NPSH Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   79
            Left            =   3600
            TabIndex        =   354
            Top             =   3780
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Hydraulic Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   78
            Left            =   3600
            TabIndex        =   353
            Top             =   3420
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Hydrostatic Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   77
            Left            =   837
            TabIndex        =   352
            Top             =   4860
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Motor Locked Rotor Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   76
            Left            =   240
            TabIndex        =   351
            Top             =   4500
            Width           =   2532
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Motor No-load Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   75
            Left            =   837
            TabIndex        =   350
            Top             =   4140
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Outline Dimensions?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   74
            Left            =   837
            TabIndex        =   349
            Top             =   3780
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "General Appearance?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   73
            Left            =   840
            TabIndex        =   348
            Top             =   3420
            Width           =   1932
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Good?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   72
            Left            =   5400
            TabIndex        =   347
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   71
            Left            =   2640
            TabIndex        =   342
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   70
            Left            =   2640
            TabIndex        =   341
            Top             =   1950
            Width           =   495
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   69
            Left            =   3960
            TabIndex        =   340
            Top             =   2550
            Width           =   735
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   3960
            TabIndex        =   339
            Top             =   1950
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "AC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   68
            Left            =   240
            TabIndex        =   332
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   67
            Left            =   3960
            TabIndex        =   331
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "V       X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   1680
            TabIndex        =   329
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "MOhms Above"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   3960
            TabIndex        =   327
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pneumatic Test for N2 Gas:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   325
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Hydrostatic Test:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   324
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Dielectric Strength:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   323
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Insulation Resistance:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   322
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "V Megger"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   1680
            TabIndex        =   321
            Top             =   750
            Width           =   975
         End
      End
      Begin VB.TextBox txtImpellerDia 
         Height          =   315
         Left            =   10560
         TabIndex        =   284
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearPumpData 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Clear Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   282
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame frmMfr 
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   234
         Top             =   600
         Width           =   2655
         Begin VB.OptionButton optMfr 
            Caption         =   "TEMC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   236
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optMfr 
            Caption         =   "Chempump"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   235
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdApproveTestDate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Approve/Unapprove This Test Date"
         Height          =   615
         Left            =   -62760
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdApprovePump 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Approve/Unapprove This Pump"
         Height          =   615
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteTestDate 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Delete This Test Date"
         Height          =   615
         Left            =   -64440
         Style           =   1  'Graphical
         TabIndex        =   196
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeletePump 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Delete This Pump"
         Height          =   615
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -60960
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Browse for File"
      End
      Begin VB.TextBox txtTDH 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -61920
         TabIndex        =   182
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Frame frmOtherFiles 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Other Files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -69840
         TabIndex        =   181
         Top             =   5040
         Width           =   3255
         Begin VB.TextBox txtNPSHFile 
            Height          =   285
            Left            =   2040
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkNPSH 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "NPSH Data:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   39
            Top             =   195
            Width           =   1455
         End
         Begin VB.CheckBox chkPictures 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pictures:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   555
            Width           =   1455
         End
         Begin VB.CheckBox chkVibration 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "FFT Vibration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   915
            Width           =   1455
         End
         Begin VB.TextBox txtPicturesFile 
            Height          =   285
            Left            =   2040
            TabIndex        =   42
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtVibrationFile 
            Height          =   285
            Left            =   2040
            TabIndex        =   44
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame frmPerfMods 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Performance Modifications"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -69840
         TabIndex        =   178
         Top             =   3000
         Width           =   3255
         Begin VB.TextBox txtOrifice 
            Height          =   285
            Left            =   1680
            TabIndex        =   31
            Top             =   1365
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtImpTrim 
            Height          =   285
            Left            =   1680
            TabIndex        =   29
            Top             =   825
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkOrifice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pump Discharge Orifice:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox chkTrimmed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Trimmed:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkFeathered 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Feathered:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblOrifice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orifice Diameter"
            Height          =   255
            Left            =   1560
            TabIndex        =   180
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblImpTrim 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Diameter"
            Height          =   255
            Left            =   1620
            TabIndex        =   179
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame frmThrustBalMods 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Thrust Balance Modifications"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3252
         Left            =   -74880
         TabIndex        =   175
         Top             =   6480
         Width           =   7215
         Begin VB.CheckBox chkAddedDiodes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Added Diodes to TRG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   414
            Top             =   2040
            Width           =   2652
         End
         Begin VB.TextBox txtNoOfDiodes 
            Height          =   288
            Left            =   5520
            TabIndex        =   412
            Top             =   2160
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtGGap 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            TabIndex        =   409
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdModifyBalanceHoleData 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Modify Balance Hole Data"
            Height          =   495
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   317
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtCircOrifice 
            Height          =   288
            Left            =   5520
            TabIndex        =   22
            Top             =   1800
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkCircOrifice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flow Orifice:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   21
            Top             =   1680
            Width           =   2772
         End
         Begin VB.CommandButton cmdAddNewBalanceHoles 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Add New Balance Hole Data"
            Height          =   495
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   185
            Top             =   120
            Width           =   1575
         End
         Begin VB.CheckBox chkBalanceHoles 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Balance Holes Modified:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   20
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtOtherMods 
            Height          =   555
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   2640
            Width           =   5055
         End
         Begin VB.TextBox txtEndPlay 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   330
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid dgBalanceHoles 
            Height          =   975
            Left            =   2400
            TabIndex        =   184
            ToolTipText     =   "Click left column (where arrow is) to select to modify or delete. Choose date in Test Data above to add new data."
            Top             =   600
            Visible         =   0   'False
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1720
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            Enabled         =   -1  'True
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label lblNoOfDiodes 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Number of Diodes:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3720
            TabIndex        =   413
            Top             =   2160
            Visible         =   0   'False
            Width           =   1812
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "G-Gap:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   89
            Left            =   240
            TabIndex        =   410
            Top             =   750
            Width           =   855
         End
         Begin VB.Label lblCircOrifice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orifice Diameter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3960
            TabIndex        =   193
            Top             =   1800
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Other Mods:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   12
            Left            =   360
            TabIndex        =   177
            Top             =   2760
            Width           =   1332
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "End Play:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   176
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame frmElecData 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Electrical Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69840
         TabIndex        =   169
         Top             =   1440
         Width           =   3255
         Begin VB.TextBox txtVFDFreq 
            Height          =   315
            Left            =   2040
            TabIndex        =   396
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtKWMult 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            TabIndex        =   25
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox cmbFrequency 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   600
            Width           =   1175
         End
         Begin VB.ComboBox cmbVoltage 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   1175
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "VFD Freq:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   86
            Left            =   1920
            TabIndex        =   397
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "KW Multiplier:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   172
            Top             =   1110
            Width           =   1215
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Freq:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   171
            Top             =   630
            Width           =   495
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Volt:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   170
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.Frame frmLoopAndXducer 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Loop and Transducer (Gauge) Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74880
         TabIndex        =   164
         Top             =   1440
         Width           =   4935
         Begin VB.ComboBox cmbMounting 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   299
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox cmbLoopNumber 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbOrificeNumber 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtHDCor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2520
            TabIndex        =   18
            Top             =   3360
            Width           =   1335
         End
         Begin VB.ComboBox cmbSuctDia 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2460
            Width           =   1335
         End
         Begin VB.ComboBox cmbDischDia 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2940
            Width           =   1335
         End
         Begin VB.TextBox txtDischHeight 
            Height          =   375
            Left            =   3600
            TabIndex        =   17
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtSuctHeight 
            Height          =   375
            Left            =   3600
            TabIndex        =   16
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Mounting:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   66
            Left            =   720
            TabIndex        =   300
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Click here for a diagram"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   840
            TabIndex        =   202
            Top             =   4080
            Width           =   3612
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orifice Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   174
            Top             =   750
            Width           =   1575
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Loop Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   173
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "HD Cor:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   168
            Top             =   3390
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suction Diameter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   167
            Top             =   2490
            Width           =   1575
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Dischge Diameter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   166
            Top             =   2970
            Width           =   1695
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Transducer Height (in Inches)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   3360
            TabIndex        =   165
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.Frame frmInstrumentTags 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Instrument Identification (Tags)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -66480
         TabIndex        =   156
         Top             =   1440
         Width           =   6135
         Begin VB.ComboBox cmbFlowMeter 
            Height          =   288
            Left            =   120
            TabIndex        =   415
            Top             =   240
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.ComboBox cmbPLCNo 
            Height          =   315
            ItemData        =   "Main.frx":40A2
            Left            =   4560
            List            =   "Main.frx":40A4
            Style           =   2  'Dropdown List
            TabIndex        =   400
            Top             =   1620
            Width           =   1455
         End
         Begin VB.ComboBox cmbTachID 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   780
            Width           =   1455
         End
         Begin VB.ComboBox cmbAnalyzerNo 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtFlowmeterID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtSuctionID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   33
            Top             =   780
            Width           =   1335
         End
         Begin VB.TextBox txtDischargeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   34
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtTemperatureID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1800
            TabIndex        =   35
            Top             =   1620
            Width           =   1335
         End
         Begin VB.TextBox txtMagflowID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            TabIndex        =   36
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "PLC:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   87
            Left            =   3480
            TabIndex        =   401
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Flowmeter:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   720
            TabIndex        =   163
            Top             =   435
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suction:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   14
            Left            =   960
            TabIndex        =   162
            Top             =   817
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discharge:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   720
            TabIndex        =   161
            Top             =   1190
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Temperature:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   160
            Top             =   1650
            Width           =   1215
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flowmeter:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   61
            Left            =   3360
            TabIndex        =   159
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Vibration/ Tach:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   62
            Left            =   3360
            TabIndex        =   158
            Top             =   720
            Width           =   972
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Analyzer:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   3480
            TabIndex        =   157
            Top             =   1230
            Width           =   975
         End
      End
      Begin VB.TextBox txtTestSetupRemarks 
         Height          =   375
         Left            =   -72360
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   9840
         Width           =   8655
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Show / Print / Export Report"
         Height          =   615
         Left            =   -61920
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox txtNPSHa 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -61920
         TabIndex        =   151
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Frame frmPumpData 
         Caption         =   "Transducers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   138
         Top             =   720
         Width           =   6735
         Begin VB.TextBox txtTemperatureDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   146
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtFlowDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   145
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtDischargeDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3675
            TabIndex        =   144
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtSuctionDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2085
            TabIndex        =   143
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTemperature 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   142
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtDischarge 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   141
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtSuction 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   140
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtFlow 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   139
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   3
            Left            =   4920
            TabIndex        =   289
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   2
            Left            =   3320
            TabIndex        =   288
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   1
            Left            =   1720
            TabIndex        =   287
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   0
            Left            =   120
            TabIndex        =   286
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Temperature"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   5280
            TabIndex        =   150
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Discharge Pressure"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   18
            Left            =   3675
            TabIndex        =   149
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTab3 
            Alignment       =   2  'Center
            Caption         =   "Suction Pressure"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2085
            TabIndex        =   148
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTab3 
            Alignment       =   2  'Center
            Caption         =   "Flow"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   147
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2415
         Left            =   -66840
         TabIndex        =   61
         Top             =   7620
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4260
         _Version        =   393216
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdAddNewTestDate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add New Test Date"
         Height          =   615
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   660
         Width           =   2055
      End
      Begin VB.TextBox txtWho 
         Height          =   315
         Left            =   -73080
         TabIndex        =   11
         Top             =   900
         Width           =   1695
      End
      Begin VB.TextBox txtSalesOrderNumber 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   420
         Width           =   1575
      End
      Begin VB.CommandButton cmdEnterTestSetupData 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Test Setup Data"
         Height          =   615
         Left            =   -66000
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnterPumpData 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Pump Data"
         Height          =   615
         Left            =   7200
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdEnterTestData 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Test Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   7800
         Width           =   1575
      End
      Begin VB.Frame fmrMiscTestData 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Miscellaneous"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   105
         Top             =   6000
         Width           =   14535
         Begin VB.TextBox txtTEMCTRGReading 
            Height          =   285
            Left            =   7080
            TabIndex        =   404
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtVibAx 
            Height          =   285
            Left            =   6000
            TabIndex        =   403
            Top             =   480
            Width           =   855
         End
         Begin VB.Frame frmTEMCData 
            BackColor       =   &H00FFFFC0&
            Caption         =   "TEMC Data"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   8400
            TabIndex        =   265
            Top             =   120
            Visible         =   0   'False
            Width           =   6015
            Begin VB.TextBox txtTEMCPVValue 
               Height          =   285
               Left            =   4080
               TabIndex        =   278
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtTEMCCalcForce 
               Height          =   285
               Left            =   4080
               TabIndex        =   276
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCViscosity 
               Height          =   285
               Left            =   2760
               TabIndex        =   274
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCThrustRigPressure 
               Height          =   285
               Left            =   1560
               TabIndex        =   272
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtTEMCMomentArm 
               Height          =   285
               Left            =   1560
               TabIndex        =   269
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCRearThrust 
               Height          =   285
               Left            =   240
               TabIndex        =   268
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtTEMCFrontThrust 
               Height          =   285
               Left            =   240
               TabIndex        =   267
               Top             =   390
               Width           =   855
            End
            Begin VB.Label lblTEMCFrontRear 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5040
               TabIndex        =   281
               Top             =   390
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblTEMCPassFail 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5040
               TabIndex        =   280
               Top             =   750
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label txtTEMCPV 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "PV"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   53
               Left            =   4080
               TabIndex        =   279
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lbltab2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Calculated Force"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   47
               Left            =   3780
               TabIndex        =   277
               Top             =   180
               Width           =   1455
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Viscosity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   52
               Left            =   2760
               TabIndex        =   275
               Top             =   180
               Width           =   855
            End
            Begin VB.Label lbltab2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Th Rig Pressure"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   51
               Left            =   1320
               TabIndex        =   273
               Top             =   750
               Width           =   1335
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Front Thrust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   50
               Left            =   120
               TabIndex        =   271
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label lbltab2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Moment Arm"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   49
               Left            =   1320
               TabIndex        =   270
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lbltab2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Rear Thrust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   48
               Left            =   180
               TabIndex        =   266
               Top             =   750
               Width           =   975
            End
         End
         Begin VB.TextBox txtRPM 
            Height          =   285
            Left            =   4680
            TabIndex        =   93
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtVibRad 
            Height          =   285
            Left            =   6000
            TabIndex        =   110
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtThrustBal 
            Height          =   285
            Left            =   4680
            TabIndex        =   108
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtTestRemarks 
            Height          =   855
            Left            =   360
            MaxLength       =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   106
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "TRG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   46
            Left            =   7200
            TabIndex        =   402
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "RPM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   4860
            TabIndex        =   113
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Y Vibration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   5880
            TabIndex        =   112
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "X Vibration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   5880
            TabIndex        =   111
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Thrust Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   4440
            TabIndex        =   109
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   360
            TabIndex        =   107
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbTestSpec 
         Height          =   315
         Left            =   -73080
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox txtRemarks 
         Height          =   555
         Left            =   3360
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   8700
         Width           =   7695
      End
      Begin VB.TextBox txtDesignTDH 
         Height          =   315
         Left            =   9360
         TabIndex        =   6
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox txtDesignFlow 
         Height          =   315
         Left            =   9360
         TabIndex        =   5
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox txtModelNo 
         Height          =   315
         Left            =   4800
         TabIndex        =   4
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox txtShpNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1410
         Width           =   4935
      End
      Begin VB.TextBox txtBilNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   1770
         Width           =   4935
      End
      Begin VB.Frame frmThermocouples 
         Caption         =   "Thermocouples"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   62
         Top             =   1800
         Width           =   6735
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   3552
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   308
            Text            =   "TC 3"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   3552
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   307
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   306
            Text            =   "TC 4"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   305
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1952
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   304
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1952
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   303
            Text            =   "TC 2"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   302
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   301
            Text            =   "TC 1"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTC4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   70
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   69
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   68
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   67
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC1Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   66
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTC2Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2080
            TabIndex        =   65
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTC3Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3680
            TabIndex        =   64
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTC4Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   63
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame frmAI 
         Caption         =   "Analog Inputs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   52
         Top             =   2880
         Width           =   6735
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   3554
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   316
            Text            =   "RBH Press"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   3554
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   315
            Text            =   "(psig)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   314
            Text            =   "AI 4"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   313
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   312
            Text            =   "Circ Flow"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   311
            Text            =   "(GPM)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   1957
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   310
            Text            =   "RBH Temp"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   1957
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   309
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtAI1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   60
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   59
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   58
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   57
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI4Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   56
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAI3Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3682
            TabIndex        =   55
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAI2Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2085
            TabIndex        =   54
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAI1Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   53
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   7
            Left            =   4920
            TabIndex        =   293
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   6
            Left            =   3320
            TabIndex        =   292
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   5
            Left            =   1720
            TabIndex        =   291
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   4
            Left            =   120
            TabIndex        =   290
            Top             =   720
            Width           =   300
         End
      End
      Begin VB.ComboBox cmbPLCLoop 
         Height          =   315
         ItemData        =   "Main.frx":40A6
         Left            =   -72480
         List            =   "Main.frx":40A8
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   420
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   -72960
         TabIndex        =   91
         Top             =   7620
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame frmChempump 
         Caption         =   "Chempump Pump Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         TabIndex        =   237
         Top             =   5040
         Width           =   14415
         Begin VB.ComboBox cmbCirculationPath 
            Height          =   315
            ItemData        =   "Main.frx":40AA
            Left            =   1800
            List            =   "Main.frx":40AC
            Style           =   2  'Dropdown List
            TabIndex        =   247
            Top             =   1860
            Width           =   3615
         End
         Begin VB.ComboBox cmbStatorFill 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   246
            Top             =   1140
            Width           =   3615
         End
         Begin VB.ComboBox cmbDesignPressure 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   245
            Top             =   1500
            Width           =   3615
         End
         Begin VB.ComboBox cmbRPM 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   244
            Top             =   780
            Width           =   3615
         End
         Begin VB.ComboBox cmbMotor 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   243
            Top             =   420
            Width           =   3615
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "User Entry - Model and Group"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   5520
            TabIndex        =   238
            Top             =   360
            Width           =   5175
            Begin VB.ComboBox cmbModelGroup 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   240
               Top             =   960
               Width           =   3615
            End
            Begin VB.ComboBox cmbModel 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   239
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lblTab1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Model:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   480
               TabIndex        =   242
               Top             =   540
               Width           =   855
            End
            Begin VB.Label lblTab1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Model Group:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   241
               Top             =   1020
               Width           =   1215
            End
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Motor:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   252
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stator Fill:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   251
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "DesignPressure:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   250
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Design RPM:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   249
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblTab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Circulation Path:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   248
            Top             =   1920
            Width           =   1575
         End
      End
      Begin MSChart20Lib.MSChart MSChart3 
         Height          =   3012
         Left            =   -73920
         OleObjectBlob   =   "Main.frx":40AE
         TabIndex        =   422
         Top             =   1080
         Width           =   5892
      End
      Begin MSChart20Lib.MSChart MSChart4 
         Height          =   3012
         Left            =   -67920
         OleObjectBlob   =   "Main.frx":6436
         TabIndex        =   423
         Top             =   1080
         Width           =   5892
      End
      Begin MSChart20Lib.MSChart MSChart5 
         Height          =   3012
         Left            =   -73920
         OleObjectBlob   =   "Main.frx":87A2
         TabIndex        =   424
         Top             =   4440
         Width           =   5892
      End
      Begin MSChart20Lib.MSChart MSChart6 
         Height          =   3012
         Left            =   -67920
         OleObjectBlob   =   "Main.frx":AB0E
         TabIndex        =   425
         Top             =   4440
         Width           =   5892
      End
      Begin VB.Shape shpGetPLCData 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   350
         Left            =   -74880
         Shape           =   3  'Circle
         Top             =   360
         Width           =   492
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "RMA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   88
         Left            =   -70920
         TabIndex        =   408
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Line Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   600
         TabIndex        =   398
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Original Impeller Dia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   8280
         TabIndex        =   285
         Top             =   4380
         Width           =   2055
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "Number of Points to Plot"
         Height          =   375
         Index           =   54
         Left            =   -62160
         TabIndex        =   203
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "TDH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   -61920
         TabIndex        =   183
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "Test Setup Remarks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   65
         Left            =   -74400
         TabIndex        =   155
         Top             =   9900
         Width           =   1812
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "NPSHa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   -61920
         TabIndex        =   152
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "Operator:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   118
         Top             =   930
         Width           =   1575
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Order:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   117
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "Test Number"
         Height          =   252
         Index           =   63
         Left            =   -74520
         TabIndex        =   115
         Top             =   9360
         Width           =   972
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "Test Specification:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   103
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2400
         TabIndex        =   102
         Top             =   8850
         Width           =   855
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Design TDH:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   7920
         TabIndex        =   100
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Design Flow:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   8040
         TabIndex        =   99
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Model No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   98
         Top             =   4380
         Width           =   1335
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bill to:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   97
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblTab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ship to:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   96
         Top             =   1440
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   -60960
         X2              =   -60360
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         X1              =   -62040
         X2              =   -61440
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label lbltab2 
         Caption         =   "Flow"
         Height          =   255
         Index           =   59
         Left            =   -60960
         TabIndex        =   95
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lbltab2 
         Caption         =   "Set Point"
         Height          =   255
         Index           =   58
         Left            =   -62040
         TabIndex        =   94
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "PLC Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   -74160
         TabIndex        =   92
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.Label lblPumpApproved 
      Caption         =   "Pump Data Approved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   198
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblTestDateApproved 
      Caption         =   "Test Setup Data Approved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   197
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version 1.10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   154
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lbltab2 
      Alignment       =   1  'Right Justify
      Caption         =   "Test Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   57
      Left            =   5640
      TabIndex        =   104
      Top             =   150
      Width           =   975
   End
   Begin VB.Label lbltab2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   60
      Left            =   960
      TabIndex        =   101
      Top             =   6900
      Width           =   855
   End
   Begin VB.Label lbltab2 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   56
      Left            =   480
      TabIndex        =   49
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPLCData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v1.1.20 - MHR - 7/12/05
'   Added Export to Excel PrintChart option

'v1.1.21 - MHR - 11/11/05
'   Modified reports to fit test remarks on the page

'v1.1.22 - MHR - 12/14/05
'   Added multiple magtrols
'   Removed tool tips

'v1.1.23 - MHR - 12/16/05
'   updated pingsilent to find magtrols
'   added excel9.olb to references so we could use  older versions of excel

'v1.1.24 - MHR - 1/16/06
'   added calibration

'v1.1.25 - MHR - 2/1/06
'   modified calibration sheet to fit on one page

'v1.1.26 - MHR - 2/21/06
'   removed barometric pressure from reports

'v1.1.27 - MHR - 3/1/06
'   modified for TEMC part numbers and parameters
'   modified data reports
'   added TEMC force data to internal form

'v1.1.28 - MHR - 4/12/06
'   added TEMC Force data to Excel sheet
'   prevent overwriting data when entering
'   added search for pumps

'v1.1.29 - MHR - 4/20/06
'   left remarks, thrust balance, etc data when selecting new test

'v1.1.30 - MHR - 4/23/06
'   fixed test remarks and test setup impeller trim read/write to file

'v1.1.31 - MHR - 4/23/06
 '  Fixed cmbtestdate to order the output by test date

'v1.1.32 - MHR - 4/24/06
'   added motor efficiency for TEMC Motors

'v1.1.33 - MHR - 4/26/06
'   Added Shaft power to excel
'   Fixed and updated Search form
'   Added percent full scale to excel export

'v1.1.34 - MHR - 5/1/06
'   Added Model Number and TEMC Frame Number Search
'   auto add a new test date when adding a new pump

'v1.1.35 - MHR - 5/8/06
'   Fixed export to excel to show negative sign for front force

'v1.1.36 - MHR - 6/1/06
'   Reloaded Search for Model Combo Box in cmdSearchForPumps_click
'   Fixed Impeller value on printout

'v1.1.37 - MHR - 7/11/06
'   changed tdh graph to go from 0 to 10 above highest reading

'v1.1.38 - MHR - 8/8/06
'   fixed impeller diameter printout

'v1.1.39 - MHR 9/20/06
'   fixed impeller diameter on excel sheet

'v1.1.40 - MHR 9/27/06
'   added radio buttons to select manually or automatically entered flow

'v1.1.41 - MHR - 11/21/06
'   when enter pump data is pressed, only add new test date if this is a new pump entry

'v1.1.42 - MHR - 6/22/07
'   limited test setup remarks field to 150 characters
'   test data remarks to 80 characters
'   pump data remarks to 255 characters
'   balance holes does not fail on cancel button
'   allowed original impeller diameter to show on temc and chempump pumps
'   added provision for 3rd magtrol at 192.0.0.118

'v1.1.43 - MHR - 6/29/07
'   Added auto/manual selection for 4 transducer inputs and 4 AI inputs
'   Carry viscosity from test to test

'v1.1.44 - MHR - 8/7/07
'   fixed new magtrol address and name to 192.0.0.118 and gpib3

'v1.1.45 - MHR - 8/8/07
'   changed gpib addresses to use database table and use ibfind to use names like gpib0

'v1.1.46 - MHR - 8/23/07
'   saved display text to data file for manual entries

'v1.1.47 - MHR - 12/17/07
'   changed export to excel to write RBHPress for AI3 - was writing RBHTemp
'   added (in/sec) to customer report with vibration

'v1.1.48 - MHR - 12/20/07
'   Added Bill To Customer to search
'   Added Ship To Customer to search
'   Added Date range to search
'   Modified Search results to show more data
'   Changed DataLists to MSHFlexGrids
'   Removed Customer from Customer Reports
'   Show Bill To and Ship To on Internal Reports

'v1.4.49 - MHR - 1/28/08
'   changed manual magtrol -
'       when entering v or i, repeats phase 1 to other 2 phases
'       allow total kw to come from entry, sum of 3 powers, or ai4
'   added 2 new reports
'       vibration w/o axial position
'       no circ flow w/o axial position
'   modified balance hole
'       only show the current date's or last date's balance hole info
'       allow addition of multiple balance holes for each test setup
'   fixed frmreport to set all options to false

'v1.1.50 - MHR - 2/12/08
'   Balance holes - 99 in diameter = slots
'                 - 99 in BC = unknown
'                 - Allow row delete
'   Reports - Made remarks Can Grow Enabled
'           - Made footer bigger
'   Search - Added Impeller size to model no search and allowed sort

'v1.1.51 - MHR - 3/5/08
'   Manual Magtrol - divide total KW by 3 and put into each Power text because we only save these values
 '   Added Mounting data: Horizontal, Suction up, Suction down, Other

'v1.1.52 - MHR - 3/19/08
'   Fixed balance holes to show older balance holes if different bolt circle

'v1.1.53 - MHR - 3/31/08
'   Fixed Balance Hole data to show correct data

'v1.1.54 - MHR - 4/2/08
'   Allow 0 for balance hole parameters to be entered

'v1.1.55 - MHR - 4/24/08
'   Fix search for pump by SN to order query by serialnumber
'   Fix printouts to show TEMC frame number instead of Chempump motor for TEMC pumps
'   Start dropdowns for transducer tags

'v1.1.56 - MHR - 10/27/08
'   Allow for new Magtrol 6350

'v1.1.57 - MHR - 11/05/08
'   Fix KW by only using total power
'   Magtrol 5300 = P1 + P2 + P3
'   Magtrol 6530 = P1 + P3
'   Also, allow for 5 significant digits in Total KW
'            0 <= Power < 1kW            Show 0.xxxxx
'            1kW <= Power < 10kW         Show x.xxxx
'            10kW <= Power < 100kW       Show xx.xxx
'            100kW <= Power              Show xxx.xx

'v1.1.58 - MHR - 11/10/08
'   Fixed DoEfficiencyCalcs if totalpower is null (like when we enter a new test date)

'v1.1.59 - MHR - 1/16/09
'   Fixed motor efficiency calculation by sending itemdata instead of listindex
'   Added new table for efficiencies, MotorEfficiencies using input power and efficiency instead of output power

'v1.1.60 - MHR - 2/9/09
'   Fixed calculation of total power in the event an older version is used to take data

'v1.1.61 - MHR - 2/12/09
'   Added Speed and SG correction to Export to Excel Report Option

'v1.1.62 - MHR - 4/16/09
'   Fixed export to excel -- failed for TEMC motors

'v1.1.63 - MHR - 4/20/09
'   Fixed DLookupA to allow searching array to end
'   Was ubound(ArrayName) -1.  Made it ubound(ArrayName)

'v1.1.64 - MHR - 4/24/09
'   Added Speed/SG correction for TEMC pumps
'   Changed TEMC Motor Efficiency calcs and Database Table

'v1.1.65 - MHR - 4/30/09
'   Allow entry of 2 lines of labels for TCs and AIs
'   Modified Pumpdata to include new table AITitles
'   Modified TDH and Vel Hd calculations to use itemdata instead of listindex

'v1.1.66 - MHR - 5/18/09
'   Fill Suction and Discharge Diameter dropdowns in size order by adding ORDER BY val(Description)
'   Allow editing, deleting and adding of Balance Hole Data
'   Fixed TDH calculation reported in txtTDH

'v.1.1.67 - MHR - 6/15/09
'   Only add default titles to AI when adding a new pump

'v1.1.68 - MHR - 6/17/09
'   Fixed velocity head calcs if no disch or suct dia
'   Restore titles when finding a new pump

'v1.1.69 - MHR - 9/28/09
'   Fixed power on Excel and Printouts from Magtrol 6530

'v1.1.70 - MHR - 11/2/09
'   Fixed Speed/SG calculations to use speed squared for head calculations

'v1.1.71 - MHR - 1/31/11
'   Removed hard coding to \\checpsa\ due to server name being changed.  Added reading / writing
 '       to server.txt in c:\Program Files\PumpData\ directory, which holds server name.

'v1.1.72 - MHR - 2/1/11
 '   Changed server.txt from c:\program files\pumpdata\ directory to c:\

'v1.1.73 - MHR - 2/6/11
'   Forced entry of all dropdowns for TEMC pump

'v1.1.74 - MHR - 2/15/11
'   Corrected Enter TEMC Frame Number so only asks when a TEMC pump

'v1.1.75 - MHR - 5/25/11
 '   Changed Eff database path from c:\Program Files\PumpData\ to app.path to accomodate Win 7 Program Files (x86) and other windows Program Files directories.

'v1.1.76 - MHR - 11/29/11
'   Changed connections strings in report data environments to reflect win 7

'v1.1.77 - MHR - 1/16/12
'   Added pingsilent when looking from PLCs

'v1.1.78 - MHR - 1/17/12
'   Fixed errors in pinging for PLCs and reduced timeout in GPIB ibdev function

'v1.2.0 - MHR - 1/19/12
'   Incorporated searching Epicor
'   Added TEMC Inspection Report
'   Added chart Flow vs Amps

'v1.2.1 - MHR - 6/21/12
'   Fixed PLC dropdown to use item data in the event that there is a plc not operating

'v1.2.2 - MHR - 8/6/12
'   Modified Reports - one customer report with options for vib, circ flow, rpm, ax pos that are
'       selected from frmReportOptions

'v1.2.3 - MHR - 8/22/12
'   Fixed controls so all are disabled correctly when approved

'v1.2.4 - MHR - 12/19/12
'   Changed Epicor904 to Epicor905 in ODBC open routine to accomodate new version

'v1.2.5 - MHR - 2/11/13
 '   Added Development.mdb at f:\Groups\Dev\3393567 as repository for database names and locations

'v1.2.6 - MHR - 3/19/13
'   Added / Correct some fields from Epicor records

'v1.2.7 - MHR - 4/30/13
'   Fixed Motor Name on Calibration Spreadsheet

'v1.2.8 - MHR - 5/3/13
'   Fixed CalibrateData.mdb and only allow calibrate if Admin login

'v1.2.9 - MHR - 10/29/13
'   Added PLCNo to Test Setup Screen
'   Fixed Search to show model no in TEMC Frame Search
'   Fixed Epicor to get Bill to and Ship to
'   Added Wildcard search
'   Fixed reports

'v1.2.10 - MHR - 11/6/13
'   Fixed Report problem when using Chempump Pump
'   Fixed plcno not storing correctly

'v1.2.11 - MHR - 11/14/13
'   Fixed so TRG is available on all pumps
'   Fixed report to allow print TRG also
'   Changes report position calculations
'   Fixed so plc only shows when one is stored

'v1.2.12 - MHR - 5/14/14
'   Changed Epicor Search to return Customer in Ship to if OrderHed.ShipToNum = ""
'   Add A-28845 to Test Specifications
'   Add Class N to TEMC Insulation
'   Add RMA to Test Setup Tab
'   Add G-Gap to Test Setup Tab
'   Added send Hydraulic Test Report to Excel

'v1.2.13 - MHR - 11/21/14
'   Swapped VibrationX and VibrationY for reports

'v1.2.14 - MHR - 6/28/17
'   Modified Epicor routine to return proper data for line number
'   Added NoOfDiodes added for TRG

'v1.2.15 - MHR - 9/21/17
 '   removed ni cw controls: cwnumedit and cwgraph and replaced with standard ms up/down and mschart
'   ni keeps complaining that we need to update controls

'v1.2.17 - MHR - 10/27/17
'   fixed dropdowns that were locked at first entry
'   autoscaled MSChart2

'v1.2.18 - MHR - 11/29/17
'   fixed Efficiency data grid and used new recordset for data source

'v1.2.19 - MHR - 3/15/18
'   requeried rsEffDisp at the end of eff calcs to update grid

'v1.2.20 - MHR - 3/19/18
'   was failing on pump data and test setup data writes
'   allowed for 0 length text fields in database tables

'v1.2.21 - MHR - 3/26/18
'   fixed vsplit routine on Magtrol data for incorrect/empty string data
'   fixed chart so it plots when writing test data

'v1.3.0 - MHR - 4/12/18
'   changed to excel template PumpData Excel Template.xls and copied exporttoexcel from polar

    Option Explicit

    Dim debugging As Integer        'debugging 1=true 0=false
    Dim sDataBaseName As String
    Dim ParentDirectoryName As String

    Dim vResponse As Variant        'Parsed response from Magtrol
    Dim sData As String             'string response from Magtrol
    Dim iUD As Integer              'GPIB address of Magtrol
    Dim vPlot(20, 1) As Variant    'arrays for mini graph

    Dim boUsingHP As Boolean            'We're using the HP database
    Dim boFoundPump As Boolean          'found the pump in database
    Dim boPumpIsApproved As Boolean     'pump data is approved
    Dim boTestDateIsApproved As Boolean 'data for this date is approved
    Dim boFoundTestSetup As Boolean     'found setup data
    Dim boFoundTestData As Boolean      'found test data
    Dim boUsingEpicor As Boolean        'search epicor for pump

    Dim boPLCOperating As Boolean       'is the PLC working?
    Dim boMagtrolOperating As Boolean   'is Magtrol working?
    Dim boGotBalanceHoles               'do we have any balance hole data?

    'recordsets
    Dim rsPumpData As New ADODB.Recordset       'PumpData recordset
    Dim rsTestSetup As New ADODB.Recordset      'TestSetup recordset
    Dim rsTestData As New ADODB.Recordset       'Test Data recordset
    Dim rsEff As New ADODB.Recordset            'Efficiency Calcs
    Dim rsEffDisp As New ADODB.Recordset        'for displaying the Efficiency calcs
    Dim rsBalanceHoles As New ADODB.Recordset   'Balance holes
    Dim rsPumpParameters As New ADODB.Recordset 'Other parameters

    'commands
    Dim qyPumpData As New ADODB.Command         'Query for PumpData
    Dim qyTestSetup As New ADODB.Command        'Query for TestSetup
    Dim qyBalanceHoles As New ADODB.Command     'query for Balance Holes

    'changing dropdown from stored data
    Dim FromStoredData As Boolean

    'array for head/flow chart
    Dim HeadFlow(1, 7) As Single            'x and y
    Dim EffFlow(1, 7) As Single
    Dim KWFlow(1, 7) As Single
    Dim AmpsFlow(1, 7) As Single
    Dim FlowHead(7, 1) As Single


    Dim RatedKW As Single               'TEMC Motor rated output

    Dim blnEnabled As Boolean           'auto enabled

    Dim EpicorConnectionString As String

    'Efficiency Database Name
    Const sEffDataBaseName As String = "\eff.mdb"

    'Server Name Text File
    Const sServerNameTextFile = "C:\Server.txt"
    Const sSaveFileMacroFile = "\savefile.bas"

    'HP Database Path
    Const sHPDataBaseName As String = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=HP-3000/32"

     'mdb at f:\groups\dev\3393567 where database names and locations reside
     ' we're using f:\ instead of a fully qualified unc since the names of the servers change

    Const sDevelopmentDatabase = "\\tei-main-01\F\Groups\DEV\3393567\Development.mdb"

    Private xlApp As Excel.Application  ' Excel Application Object
    Private xlBook As Excel.Workbook    ' Excel Workbook Object

' <VB WATCH>
Const VBWMODULE = "frmPLCData"
' </VB WATCH>

Private Sub chkAddedDiodes_Click()
           'if the AddedDiodes box is checked, show the number of diodes
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmPLCData.chkAddedDiodes_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>
10         If chkAddedDiodes.value = 1 Then
11             lblNoOfDiodes.Visible = True
12             txtNoOfDiodes.Visible = True
13         Else
14             lblNoOfDiodes.Visible = False
15             txtNoOfDiodes.Visible = False
16         End If
' <VB WATCH>
17         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
18         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkAddedDiodes_Click"

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

Private Sub chkBalanceHoles_Click()
           'if the balance holes box is checked, show the datagrid
' <VB WATCH>
19         On Error GoTo vbwErrHandler
20         Const VBWPROCNAME = "frmPLCData.chkBalanceHoles_Click"
21         If vbwProtector.vbwTraceProc Then
22             Dim vbwProtectorParameterString As String
23             If vbwProtector.vbwTraceParameters Then
24                 vbwProtectorParameterString = "()"
25             End If
26             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
27         End If
' </VB WATCH>
28         If chkBalanceHoles.value = 1 Then
29             dgBalanceHoles.Visible = True
30         Else
31             dgBalanceHoles.Visible = False
32         End If
33         If LenB(frmPLCData.txtSN.Text) = 0 Or LenB(cmbTestDate.Text) = 0 Then
34             dgBalanceHoles.Visible = False
35         End If
' <VB WATCH>
36         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
37         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkBalanceHoles_Click"

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

Private Sub chkCircOrifice_Click()
               'if the CircOrifice box is checked, show the size
' <VB WATCH>
38         On Error GoTo vbwErrHandler
39         Const VBWPROCNAME = "frmPLCData.chkCircOrifice_Click"
40         If vbwProtector.vbwTraceProc Then
41             Dim vbwProtectorParameterString As String
42             If vbwProtector.vbwTraceParameters Then
43                 vbwProtectorParameterString = "()"
44             End If
45             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
46         End If
' </VB WATCH>
47         If chkCircOrifice.value = 1 Then
48             lblCircOrifice.Visible = True
49             txtCircOrifice.Visible = True
50         Else
51             lblCircOrifice.Visible = False
52             txtCircOrifice.Visible = False
53         End If
' <VB WATCH>
54         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
55         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkCircOrifice_Click"

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


Private Sub chkNPSH_Click()
           'if the NPSH file box is checked, show the file name
' <VB WATCH>
56         On Error GoTo vbwErrHandler
57         Const VBWPROCNAME = "frmPLCData.chkNPSH_Click"
58         If vbwProtector.vbwTraceProc Then
59             Dim vbwProtectorParameterString As String
60             If vbwProtector.vbwTraceParameters Then
61                 vbwProtectorParameterString = "()"
62             End If
63             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
64         End If
' </VB WATCH>
65         If chkNPSH.value = 1 Then
66             txtNPSHFile.Visible = True
67         Else
68             txtNPSHFile.Visible = False
69         End If
' <VB WATCH>
70         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
71         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkNPSH_Click"

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

Private Sub chkOrifice_Click()
           'if the orifice box is checked, show the size
' <VB WATCH>
72         On Error GoTo vbwErrHandler
73         Const VBWPROCNAME = "frmPLCData.chkOrifice_Click"
74         If vbwProtector.vbwTraceProc Then
75             Dim vbwProtectorParameterString As String
76             If vbwProtector.vbwTraceParameters Then
77                 vbwProtectorParameterString = "()"
78             End If
79             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
80         End If
' </VB WATCH>
81         If chkOrifice.value = 1 Then
82             lblOrifice.Visible = True
83             txtOrifice.Visible = True
84         Else
85             lblOrifice.Visible = False
86             txtOrifice.Visible = False
87         End If
' <VB WATCH>
88         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
89         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkOrifice_Click"

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

Private Sub chkPictures_Click()
           'if the pictures box is checked, show the file name
' <VB WATCH>
90         On Error GoTo vbwErrHandler
91         Const VBWPROCNAME = "frmPLCData.chkPictures_Click"
92         If vbwProtector.vbwTraceProc Then
93             Dim vbwProtectorParameterString As String
94             If vbwProtector.vbwTraceParameters Then
95                 vbwProtectorParameterString = "()"
96             End If
97             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
98         End If
' </VB WATCH>
99         If chkPictures.value = 1 Then
100            txtPicturesFile.Visible = True
101        Else
102            txtPicturesFile.Visible = False
103        End If
' <VB WATCH>
104        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
105        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkPictures_Click"

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

Private Sub chkTrimmed_Click()
           'if the trimmed box is checked, show the impeller size
' <VB WATCH>
106        On Error GoTo vbwErrHandler
107        Const VBWPROCNAME = "frmPLCData.chkTrimmed_Click"
108        If vbwProtector.vbwTraceProc Then
109            Dim vbwProtectorParameterString As String
110            If vbwProtector.vbwTraceParameters Then
111                vbwProtectorParameterString = "()"
112            End If
113            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
114        End If
' </VB WATCH>
115        If chkTrimmed.value = 1 Then
116            lblImpTrim.Visible = True
117            txtImpTrim.Visible = True
118        Else
119            lblImpTrim.Visible = False
120            txtImpTrim.Visible = False
121        End If
' <VB WATCH>
122        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
123        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkTrimmed_Click"

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

Private Sub chkVibration_Click()
           'if the vibration box is checked, show the file name
' <VB WATCH>
124        On Error GoTo vbwErrHandler
125        Const VBWPROCNAME = "frmPLCData.chkVibration_Click"
126        If vbwProtector.vbwTraceProc Then
127            Dim vbwProtectorParameterString As String
128            If vbwProtector.vbwTraceParameters Then
129                vbwProtectorParameterString = "()"
130            End If
131            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
132        End If
' </VB WATCH>
133        If chkVibration.value = 1 Then
134            txtVibrationFile.Visible = True
135        Else
136            txtVibrationFile.Visible = False
137        End If
' <VB WATCH>
138        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
139        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkVibration_Click"

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

Private Sub cmbAnalyzerNo_Click()
' <VB WATCH>
140        On Error GoTo vbwErrHandler
141        Const VBWPROCNAME = "frmPLCData.cmbAnalyzerNo_Click"
142        If vbwProtector.vbwTraceProc Then
143            Dim vbwProtectorParameterString As String
144            If vbwProtector.vbwTraceParameters Then
145                vbwProtectorParameterString = "()"
146            End If
147            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
148        End If
' </VB WATCH>
' <VB WATCH>
149        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
150        Exit Sub
151        Dim LI As Integer
152        LI = cmbAnalyzerNo.ListIndex

153        Dim I As Integer
154        Dim SepNo As Integer
155        For I = 0 To cmbAnalyzerNo.ListCount - 1
156            If Left$(cmbAnalyzerNo.List(I), 4) = "----" Then
157                SepNo = I
158                Exit For
159            End If
160        Next
161        If FromStoredData = False Then
162            If LI >= SepNo Then
163                cmbAnalyzerNo.ListIndex = 0
164            End If
165        End If
' <VB WATCH>
166        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
167        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbAnalyzerNo_Click"

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
            vbwReportVariable "LI", LI
            vbwReportVariable "I", I
            vbwReportVariable "SepNo", SepNo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbFlowMeter_Click()
' <VB WATCH>
168        On Error GoTo vbwErrHandler
169        Const VBWPROCNAME = "frmPLCData.cmbFlowMeter_Click"
170        If vbwProtector.vbwTraceProc Then
171            Dim vbwProtectorParameterString As String
172            If vbwProtector.vbwTraceParameters Then
173                vbwProtectorParameterString = "()"
174            End If
175            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
176        End If
' </VB WATCH>
' <VB WATCH>
177        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
178        Exit Sub
179        Dim LI As Integer
180        LI = cmbFlowMeter.ListIndex

181        Dim I As Integer
182        Dim SepNo As Integer
183        For I = 0 To cmbFlowMeter.ListCount - 1
184            If Left$(cmbFlowMeter.List(I), 4) = "----" Then
185                SepNo = I
186                Exit For
187            End If
188        Next
189        If FromStoredData = False Then
190            If LI >= SepNo Then
191                cmbFlowMeter.ListIndex = 0
192            End If
193        End If

' <VB WATCH>
194        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
195        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbFlowMeter_Click"

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
            vbwReportVariable "LI", LI
            vbwReportVariable "I", I
            vbwReportVariable "SepNo", SepNo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbFrequency_Click()
' <VB WATCH>
196        On Error GoTo vbwErrHandler
197        Const VBWPROCNAME = "frmPLCData.cmbFrequency_Click"
198        If vbwProtector.vbwTraceProc Then
199            Dim vbwProtectorParameterString As String
200            If vbwProtector.vbwTraceParameters Then
201                vbwProtectorParameterString = "()"
202            End If
203            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
204        End If
' </VB WATCH>
205        If cmbFrequency.Text = "VFD" Then
206            txtVFDFreq.Visible = True
207            lbltab2(86).Visible = True
208        Else
209            txtVFDFreq.Visible = False
210            lbltab2(86).Visible = False
211        End If
' <VB WATCH>
212        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
213        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbFrequency_Click"

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



Private Sub cmbMagtrol_Click()
' <VB WATCH>
214        On Error GoTo vbwErrHandler
215        Const VBWPROCNAME = "frmPLCData.cmbMagtrol_Click"
216        If vbwProtector.vbwTraceProc Then
217            Dim vbwProtectorParameterString As String
218            If vbwProtector.vbwTraceParameters Then
219                vbwProtectorParameterString = "()"
220            End If
221            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
222        End If
' </VB WATCH>
223        Dim I As Integer
224        Dim sSendStr As String
225        Dim sGPIBName As String

226        I = cmbMagtrol.ItemData(cmbMagtrol.ListIndex)
227        sGPIBName = "GPIB" & I

228        If I = 99 Then      'manual entry
229            boMagtrolOperating = False
230            EnableMagtrolFields
' <VB WATCH>
231        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
232            Exit Sub
233        Else
234            boMagtrolOperating = True
235        End If


       '    If iUD <> 0 Then
       '        ibonl iUD, 0
       '    End If
       '
       '    ibdev i, 14, 0, 10, 1, 0, iUD
       '    If iberr Then
       '        i = 0
       '        boMagtrolOperating = False
       '    Else
       '        'tell the magtrol that we want full data
       '        sSendStr = "FULL" & vbCrLf
       '        ibwrt iUD, sSendStr
       '        'tell the magtrol that we don't want to wait for data
       '        sSendStr = "OPEN" & vbCrLf
       '        ibwrt iUD, sSendStr
       '        boMagtrolOperating = True
       '        DisableMagtrolFields
       '    End If

           'if we are already talking to a magtrol, close the connection
236        If iUD <> 0 Then
237            ibonl iUD, 0
238        End If

           'open a new connection to the magtrol:
               'primary address = 14
               'secondary address = 0
               'timeout = 1 second
               'eoi mode = 1
               'stop reading when line feed character is received - 0x10
               'and return iUD

239        ibdev I, 14, 0, 11, 1, &H140A, iUD

240        If iberr Then   'if we have an error
241            I = 0
242        Else
243            If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                   'tell the magtrol that we want full data
244                sSendStr = "FULL" & vbCrLf
245                ibwrt iUD, sSendStr
                   'tell the magtrol that we don't want to wait for data
246                sSendStr = "OPEN" & vbCrLf
247                ibwrt iUD, sSendStr
248            Else
249            End If
250        End If

' <VB WATCH>
251        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
252        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbMagtrol_Click"

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
            vbwReportVariable "sSendStr", sSendStr
            vbwReportVariable "sGPIBName", sGPIBName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Private Sub cmbPLCLoop_Click()
           'Change the PLC that we're looking at
' <VB WATCH>
253        On Error GoTo vbwErrHandler
254        Const VBWPROCNAME = "frmPLCData.cmbPLCLoop_Click"
255        If vbwProtector.vbwTraceProc Then
256            Dim vbwProtectorParameterString As String
257            If vbwProtector.vbwTraceParameters Then
258                vbwProtectorParameterString = "()"
259            End If
260            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
261        End If
' </VB WATCH>

262        Dim RetVal As String

           'manual data entry selection
263        If cmbPLCLoop.ListIndex = cmbPLCLoop.ListCount - 1 Then 'no plc
264            boPLCOperating = False
265            EnablePLCFields
266            If DeviceOpen = True Then
267                RetVal = DisconnectPLC()
268            End If
' <VB WATCH>
269        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
270            Exit Sub
271        End If

272        If DeviceOpen = True Then
273            RetVal = DisconnectPLC()
274        End If

275        RetVal = ConnectToPLC(cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex))
276        If RetVal <> 0 Then
277            MsgBox ("Can't connect to PLC - " & Description(cmbPLCLoop.ListIndex))
278            boPLCOperating = False
279            EnablePLCFields
280        Else
281            boPLCOperating = True
282            tDevice = cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex)
283            DisablePLCFields
284        End If
' <VB WATCH>
285        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
286        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbPLCLoop_Click"

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
            vbwReportVariable "RetVal", RetVal
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbPLCNo_Click()
' <VB WATCH>
287        On Error GoTo vbwErrHandler
288        Const VBWPROCNAME = "frmPLCData.cmbPLCNo_Click"
289        If vbwProtector.vbwTraceProc Then
290            Dim vbwProtectorParameterString As String
291            If vbwProtector.vbwTraceParameters Then
292                vbwProtectorParameterString = "()"
293            End If
294            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
295        End If
' </VB WATCH>
' <VB WATCH>
296        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
297        Exit Sub
298       Dim LI As Integer
299        LI = cmbPLCNo.ListIndex

300        Dim I As Integer
301        Dim SepNo As Integer
302        For I = 0 To cmbPLCNo.ListCount - 1
303            If Left$(cmbPLCNo.List(I), 4) = "----" Then
304                SepNo = I
305                Exit For
306            End If
307        Next
308        If FromStoredData = False Then
309            If LI >= SepNo Then
310                cmbPLCNo.ListIndex = 0
311            End If
312        End If
' <VB WATCH>
313        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
314        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbPLCNo_Click"

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
            vbwReportVariable "LI", LI
            vbwReportVariable "I", I
            vbwReportVariable "SepNo", SepNo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbTachID_Change()
' <VB WATCH>
315        On Error GoTo vbwErrHandler
316        Const VBWPROCNAME = "frmPLCData.cmbTachID_Change"
317        If vbwProtector.vbwTraceProc Then
318            Dim vbwProtectorParameterString As String
319            If vbwProtector.vbwTraceParameters Then
320                vbwProtectorParameterString = "()"
321            End If
322            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
323        End If
' </VB WATCH>
' <VB WATCH>
324        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
325            Exit Sub
326        Dim LI As Integer
327        LI = cmbTachID.ListIndex

328        Dim I As Integer
329        Dim SepNo As Integer
330        For I = 0 To cmbTachID.ListCount - 1
331            If Left$(cmbTachID.List(I), 4) = "----" Then
332                SepNo = I
333                Exit For
334            End If
335        Next
336        If FromStoredData = False Then
337            If LI >= SepNo Then
338                cmbTachID.ListIndex = 0
339            End If
340        End If
' <VB WATCH>
341        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
342        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbTachID_Change"

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
            vbwReportVariable "LI", LI
            vbwReportVariable "I", I
            vbwReportVariable "SepNo", SepNo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbTestDate_Click()
           'select a test date to show
' <VB WATCH>
343        On Error GoTo vbwErrHandler
344        Const VBWPROCNAME = "frmPLCData.cmbTestDate_Click"
345        If vbwProtector.vbwTraceProc Then
346            Dim vbwProtectorParameterString As String
347            If vbwProtector.vbwTraceParameters Then
348                vbwProtectorParameterString = "()"
349            End If
350            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
351        End If
' </VB WATCH>

352        Dim sName As String
353        Dim sParam As String
354        Dim I As Integer
355        Dim j As Integer
356        Dim k As Integer
357        Dim bSk As Boolean
358        Dim sBC As Single
359        Dim NOK() As Long

360        cmdModifyBalanceHoleData.Visible = False


361        If Not boFoundTestSetup Then    'if we don't have any TestSetup data written
362            boFoundTestData = False
' <VB WATCH>
363        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
364            Exit Sub
365        End If


           'select the testsetup data for the serial number
366        qyTestSetup.ActiveConnection = cnPumpData
367        qyTestSetup.CommandText = "SELECT * " & _
           "From TempTestSetupData " & _
           "Where (((TempTestSetupData.SerialNumber) = '" & txtSN.Text & "') AND TempTestSetupData.Date = #" & cmbTestDate.List(cmbTestDate.ListIndex) & "#) " & _
           "ORDER BY TempTestSetupData.Date;"

           '"SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
           txtSN.Text & "'))ORDER BY TempTestSetupData.Date;"

368        If rsTestSetup.State = adStateOpen Then
369            rsTestSetup.Close
370        End If

371        With rsTestSetup     'open the recordset for the query
       '        .Index = "FindData"
372            .CursorLocation = adUseClient
373            .CursorType = adOpenStatic
374            .Open qyTestSetup
375        End With

           'move to the selected date
376        rsTestSetup.MoveFirst
       '
           'show the correct combo box entries for this record
377        SetComboTestSetup cmbOrificeNumber, "OrificeNumber", "OrificeNumber", rsTestSetup
378        SetComboTestSetup cmbTestSpec, "TestSpec", "TestSpecification", rsTestSetup
379        SetComboTestSetup cmbLoopNumber, "LoopNumber", "LoopNumber", rsTestSetup
380        SetComboTestSetup cmbSuctDia, "SuctDiam", "SuctionDiameter", rsTestSetup
381        SetComboTestSetup cmbDischDia, "DischDiam", "DischargeDiameter", rsTestSetup
382        SetComboTestSetup cmbTachID, "TachID", "TachID", rsTestSetup
383        SetComboTestSetup cmbAnalyzerNo, "AnalyzerNo", "AnalyzerNo", rsTestSetup
384        SetComboTestSetup cmbVoltage, "Voltage", "Voltage", rsTestSetup
385        SetComboTestSetup cmbFrequency, "Frequency", "Frequency", rsTestSetup
386        SetComboTestSetup cmbMounting, "Mounting", "Mounting", rsTestSetup
387        SetComboTestSetup cmbPLCNo, "PLCNo", "PLCNo", rsTestSetup

       ' use this for flowmeter dropdown
       '    SetComboTestSetup cmbFlowMeter, "FlowmeterID", "Flowmeter", rsTestSetup


           'show the correct data in the text boxes
388        sName = "FlowmeterID"
389        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
390            sParam = rsTestSetup.Fields(sName)
391        Else
392            sParam = vbNullString   '""
393        End If
394        txtFlowmeterID.Text = sParam

395        sName = "SuctionID"
396        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
397            sParam = rsTestSetup.Fields(sName)
398        Else
399            sParam = vbNullString
400        End If
401        txtSuctionID.Text = sParam

402        sName = "DischID"
403        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
404            sParam = rsTestSetup.Fields(sName)
405        Else
406            sParam = vbNullString
407        End If
408        txtDischargeID.Text = sParam

409        sName = "TemperatureID"
410        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
411            sParam = rsTestSetup.Fields(sName)
412        Else
413            sParam = vbNullString
414        End If
415        txtTemperatureID.Text = sParam

416        sName = "MagflowID"
417        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
418            sParam = rsTestSetup.Fields(sName)
419        Else
420            sParam = vbNullString
421        End If
422        txtMagflowID.Text = sParam

423        sName = "HDCor"
424        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
425            sParam = rsTestSetup.Fields(sName)
426        Else
427            sParam = vbNullString
428        End If
429        txtHDCor.Text = sParam

430        sName = "KWMult"
431        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
432            sParam = rsTestSetup.Fields(sName)
433        Else
434            sParam = vbNullString
435        End If
436        txtKWMult.Text = sParam

437        sName = "Who"
438        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
439            sParam = rsTestSetup.Fields(sName)
440        Else
441            sParam = vbNullString
442        End If
443        txtWho.Text = sParam

444        sName = "RMA"
445        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
446            sParam = rsTestSetup.Fields(sName)
447        Else
448            sParam = vbNullString
449        End If
450        txtRMA.Text = sParam

451        sName = "Remarks"
452        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
453            sParam = rsTestSetup.Fields(sName)
454        Else
455            sParam = vbNullString
456        End If
457        txtTestSetupRemarks.Text = sParam

458        sName = "VFDFrequency"
459        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
460            sParam = rsTestSetup.Fields(sName)
461        Else
462            sParam = vbNullString
463        End If
464        txtVFDFreq.Text = sParam

465        sName = "SuctionGageHeight"
466        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
467            sParam = rsTestSetup.Fields(sName)
468        Else
469            sParam = 0
470        End If
471        txtSuctHeight.Text = sParam

472        sName = "DischargeGageHeight"
473        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
474            sParam = rsTestSetup.Fields(sName)
475        Else
476            sParam = 0
477        End If
478        txtDischHeight.Text = sParam

479        sName = "EndPlay"
480        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
481            sParam = rsTestSetup.Fields(sName)
482        Else
483            sParam = vbNullString
484        End If
485        txtEndPlay.Text = sParam

486        sName = "GGAP"
487        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
488            sParam = rsTestSetup.Fields(sName)
489        Else
490            sParam = vbNullString
491        End If
492        txtGGap.Text = sParam

493        sName = "OtherMods"
494        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
495            sParam = rsTestSetup.Fields(sName)
496        Else
497            sParam = vbNullString
498        End If
499        txtOtherMods.Text = sParam

500        If rsTestSetup.Fields("ImpFeathered") Then
501            chkFeathered.value = 1
502        Else
503            chkFeathered.value = 0
504        End If

505        If Val(rsTestSetup.Fields("ImpTrimmed")) = 0 Then
506            chkTrimmed.value = 0
507            txtImpTrim.Visible = False
508            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
509        Else
510            chkTrimmed.value = 1
511            txtImpTrim.Visible = True
512            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
513        End If

514        If Val(rsTestSetup.Fields("PumpDischOrifice")) = 0 Then
515            chkOrifice.value = 0
516            txtOrifice.Visible = False
517        Else
518            chkOrifice.value = 1
519            txtOrifice.Visible = True
520            txtOrifice.Text = rsTestSetup.Fields("PumpDischOrifice")
521        End If

522        If Val(rsTestSetup.Fields("CircFlowOrifice")) = 0 Then
523            chkCircOrifice.value = 0
524            txtCircOrifice.Visible = False
525        Else
526            chkCircOrifice.value = 1
527            txtCircOrifice.Visible = True
528            txtCircOrifice.Text = rsTestSetup.Fields("CircFlowOrifice")
529        End If

530        If IsNull(rsTestSetup.Fields("NoOfTRGDiodes")) Then
531            Me.chkAddedDiodes.value = 0
532            Me.txtNoOfDiodes.Visible = False
533        Else
534            chkAddedDiodes.value = 1
535            txtNoOfDiodes.Visible = True
536            txtNoOfDiodes.Text = rsTestSetup.Fields("NoOfTRGDiodes")
537        End If

538         If Not IsNull(rsTestSetup.Fields("NoOfTRGDiodes")) Then
539            If Val(rsTestSetup.Fields("NoOfTRGDiodes")) = 0 Then
540                Me.chkAddedDiodes.value = 0
541                Me.txtNoOfDiodes.Visible = False
542            Else
543                chkAddedDiodes.value = 1
544                txtNoOfDiodes.Visible = True
545                txtNoOfDiodes.Text = rsTestSetup.Fields("NoOfTRGDiodes")
546            End If
547        End If

548       If (IsNull(rsTestSetup.Fields("NPSHFile"))) Or (LenB(rsTestSetup.Fields("NPSHFile")) = 0) Then
549            chkNPSH.value = 0
550            txtNPSHFile.Visible = False
551        Else
552            chkNPSH.value = 1
553            txtNPSHFile.Visible = True
554            txtNPSHFile.Text = rsTestSetup.Fields("NPSHFile")
555        End If

556        If (IsNull(rsTestSetup.Fields("PictureFile"))) Or (LenB(rsTestSetup.Fields("PictureFile")) = 0) Then
557            chkPictures.value = 0
558            txtPicturesFile.Visible = False
559        Else
560            chkPictures.value = 1
561            txtPicturesFile.Visible = True
562            txtPicturesFile.Text = rsTestSetup.Fields("PictureFile")
563        End If

564        If (IsNull(rsTestSetup.Fields("VibrationFile"))) Or (LenB(rsTestSetup.Fields("VibrationFile")) = 0) Then
565            chkVibration.value = 0
566            txtVibrationFile.Visible = False
567        Else
568            chkVibration.value = 1
569            txtVibrationFile.Visible = True
570            txtVibrationFile.Text = rsTestSetup.Fields("VibrationFile")
571        End If


           'for TEMC Inspection Report
572        sName = "InsulationMeggerVolts"
573        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
574            sParam = rsTestSetup.Fields(sName)
575        Else
576            sParam = 0
577        End If
578        txtTestAndInspection(0).Text = sParam

579        sName = "InsulationMegOhms"
580        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
581            sParam = rsTestSetup.Fields(sName)
582        Else
583            sParam = 0
584        End If
585        txtTestAndInspection(1).Text = sParam

586        sName = "DielectricVolts"
587        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
588            sParam = rsTestSetup.Fields(sName)
589        Else
590            sParam = 0
591        End If
592        txtTestAndInspection(2).Text = sParam

593        sName = "DielectricTime"
594        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
595            sParam = rsTestSetup.Fields(sName)
596        Else
597            sParam = 0
598        End If
599        txtTestAndInspection(3).Text = sParam

600        sName = "HydrostaticValue"
601        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
602            sParam = rsTestSetup.Fields(sName)
603        Else
604            sParam = 0
605        End If
606        txtTestAndInspection(4).Text = sParam

607        sName = "HydrostaticTime"
608        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
609            sParam = rsTestSetup.Fields(sName)
610        Else
611            sParam = 0
612        End If
613        txtTestAndInspection(5).Text = sParam

614        sName = "PneumaticValue"
615        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
616            sParam = rsTestSetup.Fields(sName)
617        Else
618            sParam = 0
619        End If
620        txtTestAndInspection(6).Text = sParam

621        sName = "PneumaticTime"
622        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
623            sParam = rsTestSetup.Fields(sName)
624        Else
625            sParam = 0
626        End If
627        txtTestAndInspection(7).Text = sParam

628        For I = 0 To cmbTestAndInspection(0).ListCount - 1
629            If cmbTestAndInspection(0).Text = rsTestSetup.Fields("HydrostaticUnits") Then
630                    cmbTestAndInspection(0).ListIndex = I
631                    Exit For
632            End If
633            cmbTestAndInspection(0).ListIndex = -1
634        Next I


635        For I = 0 To cmbTestAndInspection(1).ListCount - 1
636            If cmbTestAndInspection(1).Text = rsTestSetup.Fields("PneumaticUnits") Then
637                    cmbTestAndInspection(1).ListIndex = I
638                    Exit For
639            End If
640            cmbTestAndInspection(1).ListIndex = -1
641        Next I

642        TestAndInspectionGood(0).value = Abs(rsTestSetup!insulationgood)
643        TestAndInspectionGood(1).value = Abs(rsTestSetup!DielectricGood)
644        TestAndInspectionGood(2).value = Abs(rsTestSetup!HydrostaticGood)
645        TestAndInspectionGood(3).value = Abs(rsTestSetup!PneumaticGood)
646        TestAndInspectionGood(4).value = Abs(rsTestSetup!GeneralAppearanceGood)
647        TestAndInspectionGood(5).value = Abs(rsTestSetup!OutlineDimensionsGood)
648        TestAndInspectionGood(6).value = Abs(rsTestSetup!MotorNoLoadTestGood)
649        TestAndInspectionGood(7).value = Abs(rsTestSetup!MotorLockedRotorTestGood)
650        TestAndInspectionGood(8).value = Abs(rsTestSetup!HydrostaticTestGood)
651        TestAndInspectionGood(9).value = Abs(rsTestSetup!HydraulicTestGood)
652        TestAndInspectionGood(10).value = Abs(rsTestSetup!NPSHTestGood)
653        TestAndInspectionGood(11).value = Abs(rsTestSetup!CleanPurgeSealGood)
654        TestAndInspectionGood(12).value = Abs(rsTestSetup!PaintCheckGood)
655        TestAndInspectionGood(13).value = Abs(rsTestSetup!NameplateGood)
656        TestAndInspectionGood(14).value = Abs(rsTestSetup!SupervisorApproval)

       '    rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
       '    If rsBalanceHoles.RecordCount = 0 Then
       '        rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "' AND Date <= #" & cmbTestDate.Text & "#"
       '        If rsBalanceHoles.RecordCount = 0 Then
       '            chkBalanceHoles.value = 0
       '            dgBalanceHoles.Visible = False
       '        Else
       '            rsBalanceHoles.MoveLast
       '            rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "' AND Date = #" & rsBalanceHoles.Fields("Date") & "#"
       '            chkBalanceHoles.value = 1
       '            dgBalanceHoles.Visible = True
       '        End If
       '    Else
       '        chkBalanceHoles.value = 1
       '        dgBalanceHoles.Visible = True
       '    End If
       '

       '    rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "' AND Date <= #" & cmbTestDate.Text & "#"
657        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

658        If rsBalanceHoles.RecordCount = 0 Then
659            chkBalanceHoles.value = 0
660            dgBalanceHoles.Visible = False
661            boGotBalanceHoles = False
662        Else
663            boGotBalanceHoles = True
664            ReDim NOK(rsBalanceHoles.RecordCount)
665            rsBalanceHoles.MoveLast
666            For I = 1 To rsBalanceHoles.RecordCount
667                NOK(I) = 0
668            Next I

669            For j = 1 To rsBalanceHoles.RecordCount - 1
670                rsBalanceHoles.MoveFirst
671                rsBalanceHoles.Move rsBalanceHoles.RecordCount - j
672                sBC = rsBalanceHoles.Fields("BoltCircle")
673                bSk = False
674                For k = 1 To rsBalanceHoles.RecordCount
675                    If NOK(k) = rsBalanceHoles.Fields(0) Then
676                        bSk = True
677                    End If
678                Next k
679                If Not bSk Then
680                    For I = rsBalanceHoles.RecordCount - j To 1 Step -1
681                        rsBalanceHoles.MovePrevious
682                        If rsBalanceHoles.Fields("BoltCircle") = sBC Then
683                            NOK(I) = rsBalanceHoles.Fields(0)
684                        End If
685                    Next I
686                End If
687            Next j

688            Dim sFilt
689            sFilt = ""
690            For I = 1 To rsBalanceHoles.RecordCount
691                If NOK(I) <> 0 Then
692                    sFilt = sFilt & "(BalanceHoleID <> " & NOK(I) & ") AND "
       '                sFilt = sFilt & "(" & rsBalanceHoles.Filter & " AND BalanceHoleID <> " & NOK(I) & ") AND "
693                End If
694            Next I

695            If Len(sFilt) > 4 Then
696                sFilt = Left(sFilt, Len(sFilt) - 4)
697                rsBalanceHoles.Filter = sFilt
698            End If

699            chkBalanceHoles.value = 1
700            dgBalanceHoles.Visible = True
       '        Set dgBalanceHoles.DataSource = rsBalanceHoles
701        End If
       '
           'set the test date filter for the test data
702        rsTestData.Filter = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"

703        If rsTestData.RecordCount = 0 Then
704            boFoundTestData = False
705            AddTestData
706            EnableTestDataControls
707            MsgBox "No Test Data Exists for this Serial Number"
708        Else
709            boFoundTestData = True
710            DisableTestDataControls                         'if it's in the real database, don't allow changes here
711        End If

712        If Not boTestDateIsApproved Then    'data approved?
713            EnableTestDataControls
714        End If

715        If rsTestSetup.Fields("Approved") = True Then
716            DisableTestDataControls                         'if it's in the real database, don't allow changes here
717            lblTestDateApproved.Visible = True
718            MsgBox ("Found pump.  Data cannot be modified.")
719            If boCanApprove Then
720                cmdApproveTestDate.Caption = "Unapprove this Test Date"
721            End If
722        Else
723            EnableTestDataControls                          'it's in the temp database, allow changes
724            lblTestDateApproved.Visible = False
725            If boPumpIsApproved = True Then
726                MsgBox ("Found pump.  Pump data cannot be modified, but test setup data and test data can be modified.")
727            Else
728                MsgBox ("Found pump.  Pump data, test setup data, and test data can be modified.")
729            End If
730            If boCanApprove Then
731                If rsPumpData.Fields("Approved") = True Then
732                    cmdApproveTestDate.Enabled = True
733                    cmdApproveTestDate.Caption = "Approve this Test Date"
734                Else
735                    cmdApproveTestDate.Caption = "You Must Approve Pump First"
736                    cmdApproveTestDate.Enabled = False
737                End If
738            End If
739        End If

740        rsEff.MoveFirst
741        rsTestData.MoveFirst

742        For I = 1 To rsTestData.RecordCount
743            DoEfficiencyCalcs
744            rsEff.MoveNext
745            rsTestData.MoveNext
746        Next I

           'get a recordset to display
747        If rsEffDisp.State = adStateOpen Then
748            rsEffDisp.Close
749        End If

750        Dim qyEffDisp As New ADODB.Command
751        qyEffDisp.ActiveConnection = cnEffData
752        qyEffDisp.CommandText = "SELECT Flow, TDH, KW, Volts, Amps, OverallEfficiency FROM Efficiency;"

753        With rsEffDisp     'open the recordset for the query
754            .CursorLocation = adUseClient
755            .CursorType = adOpenStatic
756            .LockType = adLockOptimistic
757            .Open qyEffDisp
758        End With


          ' fix the datagrid
759       Set DataGrid1.DataSource = rsTestData
760       Set DataGrid2.DataSource = rsEffDisp

761       Dim c As Column
762       For Each c In DataGrid1.Columns
763          Select Case c.DataField
             Case "TestDataID"     'Hide some columns
764             c.Visible = False
765          Case "SerialNumber"
766             c.Visible = False
767          Case "Date"
768             c.Visible = False
769          Case Else             ' Show all other columns.
770             c.Visible = True
771             c.Alignment = dbgRight
772          End Select
773        Next c

       'DataGrid2.Columns(0).NumberFormat = "###0.00"
       'DataGrid2.Columns(1).NumberFormat = "N2"
       'DataGrid2.Columns(2).NumberFormat = "N2"
       'DataGrid2.Columns(3).NumberFormat = "N2"
       'DataGrid2.Columns(4).NumberFormat = "N2"
       'DataGrid2.Columns(5).NumberFormat = "N2"

774        For Each c In DataGrid2.Columns
775            c.Alignment = dbgCenter
776            c.Width = 750
777            c.NumberFormat = "###0.00"
778            Select Case c.ColIndex
                   Case 0
779                    c.Caption = "Flow"
780                    c.NumberFormat = "###0.00"
781                Case 1
782                    c.Caption = "TDH"
783                    c.NumberFormat = "##0.00"
784                Case 2
785                    c.Caption = "Input Pwr"
786                    c.NumberFormat = "##0.00"
787                    c.Width = 850
788                    c.Visible = True
789                Case 3
790                    c.Caption = "Voltage"
791                    c.NumberFormat = "##0.00"
792                Case 4
793                    c.Caption = "Current"
794                    c.NumberFormat = "##0.00"
795                Case 5
796                    c.Caption = "Overall Eff"
797                    c.NumberFormat = "##0.00"
798                    c.Width = 850
       '            Case 7
       '                c.Caption = "NPSHr"
       '                c.NumberFormat = "#0.00"
799                Case Else
                       'c.Visible = False
800            End Select
801        Next c


802        txtUpDn1.Text = 1
803        UpDown2.value = rsTestData.RecordCount

       'unlock the text boxes
804        For I = 0 To 7
805            txtTitle(I).Locked = False
806        Next I

807        For I = 20 To 27
808            txtTitle(I).Locked = False
809        Next I

       ''set TC and AI labels with default values
       '    txtTitle(0).Text = "TC 1"
       '    txtTitle(1).Text = "(F)"
       '    txtTitle(2).Text = "TC 2"
       '    txtTitle(3).Text = "(F)"
       '    txtTitle(4).Text = "TC 3"
       '    txtTitle(5).Text = "(F)"
       '    txtTitle(6).Text = "TC 4"
       '    txtTitle(7).Text = "(F)"
       '    txtTitle(20).Text = "Circ Flow"
       '    txtTitle(21).Text = "(GPM)"
       '    txtTitle(22).Text = "RBH Temp"
       '    txtTitle(23).Text = "(F)"
       '    txtTitle(24).Text = "RBH Press"
       '    txtTitle(25).Text = "(psig)"
       '    txtTitle(26).Text = "AI 4"
       '    txtTitle(27).Text = ""

       'look for titles for TCs and AIs
810        Dim qy As New ADODB.Command
811        Dim rs As New ADODB.Recordset

812        qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
813        qy.CommandText = "SELECT * FROM AITitles " & _
               "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
               "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#)); "

814        With rs     'open the recordset for the query
815            .CursorLocation = adUseClient
816            .CursorType = adOpenStatic
817            .LockType = adLockOptimistic
818            .Open qy
819        End With

820        If Not (rs.BOF = True And rs.EOF = True) Then   'update titles
821            rs.MoveFirst
822            Do While Not rs.EOF
823                txtTitle(rs.Fields("Channel")).Text = rs.Fields("Title")
824                rs.MoveNext
825            Loop
826        End If

827        rs.Close
828        Set rs = Nothing
829        Set qy = Nothing
' <VB WATCH>
830        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
831        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbTestDate_Click"

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
            vbwReportVariable "sName", sName
            vbwReportVariable "sParam", sParam
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "k", k
            vbwReportVariable "bSk", bSk
            vbwReportVariable "sBC", sBC
            vbwReportVariable "NOK", NOK
            vbwReportVariable "sFilt", sFilt
            vbwReportVariable "qyEffDisp", qyEffDisp
            vbwReportVariable "c", c
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdAddNewBalanceHoles_Click()
' <VB WATCH>
832        On Error GoTo vbwErrHandler
833        Const VBWPROCNAME = "frmPLCData.cmdAddNewBalanceHoles_Click"
834        If vbwProtector.vbwTraceProc Then
835            Dim vbwProtectorParameterString As String
836            If vbwProtector.vbwTraceParameters Then
837                vbwProtectorParameterString = "()"
838            End If
839            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
840        End If
' </VB WATCH>
841        Dim strInput As String
842        Dim I As Integer
843        Dim sNumber As Integer
844        Dim sDia As Single
845        Dim sBC As Single

           'get the data for the balance holes
846        strInput = InputBox("Enter Number of Holes")
847        If strInput <> "" Then
848            sNumber = CInt(strInput)
849        Else
850            GoTo CancelPressed
851        End If

852        strInput = InputBox("Enter Decimal Value of Hole Diameter or Slot (For Example, 0.675) ")
853        If strInput <> "" Then
854            If UCase(strInput) = "SLOT" Then
855                strInput = 99
856            End If
857            sDia = CSng(strInput)
858        Else
859            GoTo CancelPressed
860        End If

861        strInput = InputBox("Enter Decimal Value of Bolt Circle or Unknown (For Example, 4.525)")
862        If strInput <> "" Then
863            If UCase(strInput) = "UNKNOWN" Then
864                strInput = 99
865            End If
866            sBC = CSng(strInput)
867        Else
868            GoTo CancelPressed
869        End If

870        If rsBalanceHoles.State <> adStateOpen Then
871            rsBalanceHoles.Open
872        End If

873        rsBalanceHoles.AddNew
874        rsBalanceHoles!SerialNo = txtSN.Text
875        rsBalanceHoles!Date = cmbTestDate.Text
876        rsBalanceHoles!Number = sNumber
877        rsBalanceHoles!diameter = sDia
878        rsBalanceHoles!boltcircle = sBC

879        rsBalanceHoles.Update
       '    rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"
       '    rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "' AND Date <= #" & cmbTestDate.Text & "#"

880        GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '    rsBalanceHoles.Requery
881        rsBalanceHoles.MoveLast
882        dgBalanceHoles.Refresh
883        chkBalanceHoles.value = 1


       '    Set dgBalanceHoles.DataSource = rsBalanceHoles

       '    Dim c As Column
       '    For Each c In dgBalanceHoles.Columns
       '        Select Case c.DataField
       '        Case "BalanceHoleID"
       '            c.Visible = False
       '        Case "SerialNo"
       '            c.Visible = False
       '        Case "Date"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 2000
       '        Case "Number"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 700
       '        Case "Diameter"
       '            c.Visible = False
       '        Case "Diameter1"
       '            c.Caption = "Diameter"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 700
       '        Case "BoltCircle1"
       '            c.Caption = "Bolt Circle"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 800
       '        Case "BoltCircle"
       '            c.Visible = False
       '        Case Else
       '        End Select
       '    Next c
' <VB WATCH>
884        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
885        Exit Sub

886    CancelPressed:
887        MsgBox "No New Balance Hole Data Entered", vbOKOnly
' <VB WATCH>
888        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
889        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdAddNewBalanceHoles_Click"

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
            vbwReportVariable "strInput", strInput
            vbwReportVariable "I", I
            vbwReportVariable "sNumber", sNumber
            vbwReportVariable "sDia", sDia
            vbwReportVariable "sBC", sBC
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdAddNewTestDate_Click()
           'add a new test date/time
' <VB WATCH>
890        On Error GoTo vbwErrHandler
891        Const VBWPROCNAME = "frmPLCData.cmdAddNewTestDate_Click"
892        If vbwProtector.vbwTraceProc Then
893            Dim vbwProtectorParameterString As String
894            If vbwProtector.vbwTraceParameters Then
895                vbwProtectorParameterString = "()"
896            End If
897            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
898        End If
' </VB WATCH>
899        Dim I As Integer

900        For I = 1 To cmbTestDate.ListCount      'see if we already have today's date entered
901            If cmbTestDate.List(I) = Date Then
902                MsgBox "There is already an entry for today.  You can only have one entry for each Serial Number and a given date.  You may want to modify the Serial Number.", vbOKOnly
' <VB WATCH>
903        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
904                Exit Sub
905            End If
906        Next I

           'we didn't find today's date entered, allow data entry
907        boFoundTestSetup = False

908        EnableTestSetupDataControls
909        cmdEnterTestSetupData_Click
910        cmdAddNewBalanceHoles.Visible = True
911        txtWho.Text = LogInInitials
912        MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
' <VB WATCH>
913        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
914        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdAddNewTestDate_Click"

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
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdApprovePump_Click()
           'allow the pump data to be approved
' <VB WATCH>
915        On Error GoTo vbwErrHandler
916        Const VBWPROCNAME = "frmPLCData.cmdApprovePump_Click"
917        If vbwProtector.vbwTraceProc Then
918            Dim vbwProtectorParameterString As String
919            If vbwProtector.vbwTraceParameters Then
920                vbwProtectorParameterString = "()"
921            End If
922            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
923        End If
' </VB WATCH>
924        rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
925        rsPumpData.Update
926        rsPumpData.Requery
927        lblPumpApproved.Visible = rsPumpData.Fields("Approved")
928        If rsPumpData.Fields("Approved") = True Then
929            cmdApprovePump.Caption = "Unapprove This Pump"
930            cmdApproveTestDate.Enabled = True
931            If rsTestSetup.Fields("Approved") = True Then
932                cmdApproveTestDate.Caption = "Unapprove This Test Date"
933            Else
934                cmdApproveTestDate.Caption = "Approve This Test Date"
935            End If
936        Else
937            cmdApprovePump.Caption = "Approve This Pump"
938            cmdApproveTestDate.Caption = "You Must Approve Pump First"
939            cmdApproveTestDate.Enabled = False
940        End If
' <VB WATCH>
941        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
942        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdApprovePump_Click"

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

Private Sub cmdApproveTestDate_Click()
           'allow the test setup data to be approved
' <VB WATCH>
943        On Error GoTo vbwErrHandler
944        Const VBWPROCNAME = "frmPLCData.cmdApproveTestDate_Click"
945        If vbwProtector.vbwTraceProc Then
946            Dim vbwProtectorParameterString As String
947            If vbwProtector.vbwTraceParameters Then
948                vbwProtectorParameterString = "()"
949            End If
950            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
951        End If
' </VB WATCH>
952        rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
953        rsTestSetup.Update
954        rsTestSetup.Requery
955        lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
956        If rsTestSetup.Fields("Approved") = True Then
957            cmdApproveTestDate.Caption = "Unapprove This Test Date"
958        Else
959            cmdApproveTestDate.Caption = "Approve This Test Date"
960        End If
' <VB WATCH>
961        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
962        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdApproveTestDate_Click"

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

Private Sub cmdCalibrate_Click()
' <VB WATCH>
963        On Error GoTo vbwErrHandler
964        Const VBWPROCNAME = "frmPLCData.cmdCalibrate_Click"
965        If vbwProtector.vbwTraceProc Then
966            Dim vbwProtectorParameterString As String
967            If vbwProtector.vbwTraceParameters Then
968                vbwProtectorParameterString = "()"
969            End If
970            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
971        End If
' </VB WATCH>
972        Dim ans As Integer
973        Dim I As Integer

974        ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
975        If ans = vbNo Then
976            Calibrating = False
' <VB WATCH>
977        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
978            Exit Sub
979        Else
980            CalibrateSoftware
981        End If
' <VB WATCH>
982        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
983        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdCalibrate_Click"

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
            vbwReportVariable "ans", ans
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

Private Sub cmdClearPumpData_Click()
' <VB WATCH>
984        On Error GoTo vbwErrHandler
985        Const VBWPROCNAME = "frmPLCData.cmdClearPumpData_Click"
986        If vbwProtector.vbwTraceProc Then
987            Dim vbwProtectorParameterString As String
988            If vbwProtector.vbwTraceParameters Then
989                vbwProtectorParameterString = "()"
990            End If
991            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
992        End If
' </VB WATCH>
993        BlankData
' <VB WATCH>
994        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
995        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdClearPumpData_Click"

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

Private Sub cmdDeletePump_Click()
           'delete this pump
' <VB WATCH>
996        On Error GoTo vbwErrHandler
997        Const VBWPROCNAME = "frmPLCData.cmdDeletePump_Click"
998        If vbwProtector.vbwTraceProc Then
999            Dim vbwProtectorParameterString As String
1000           If vbwProtector.vbwTraceParameters Then
1001               vbwProtectorParameterString = "()"
1002           End If
1003           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1004       End If
' </VB WATCH>
1005       Dim Answer As Integer
1006       Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
1007       If Answer = vbYes Then
1008           rsPumpData.Delete
1009           rsPumpData.Update
1010           cmdFindPump_Click
1011       End If
' <VB WATCH>
1012       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1013       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdDeletePump_Click"

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
            vbwReportVariable "Answer", Answer
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdDeleteTestDate_Click()
           'delete this test date
' <VB WATCH>
1014       On Error GoTo vbwErrHandler
1015       Const VBWPROCNAME = "frmPLCData.cmdDeleteTestDate_Click"
1016       If vbwProtector.vbwTraceProc Then
1017           Dim vbwProtectorParameterString As String
1018           If vbwProtector.vbwTraceParameters Then
1019               vbwProtectorParameterString = "()"
1020           End If
1021           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1022       End If
' </VB WATCH>
1023       Dim Answer As Integer
1024       Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
1025       If Answer = vbYes Then
1026           rsTestSetup.Delete
1027           rsTestSetup.Update
1028           cmdFindPump_Click
1029       End If
' <VB WATCH>
1030       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1031       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdDeleteTestDate_Click"

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
            vbwReportVariable "Answer", Answer
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdEnterPumpData_Click()
           'store the data on the screen to the pump (pumpdata)
' <VB WATCH>
1032       On Error GoTo vbwErrHandler
1033       Const VBWPROCNAME = "frmPLCData.cmdEnterPumpData_Click"
1034       If vbwProtector.vbwTraceProc Then
1035           Dim vbwProtectorParameterString As String
1036           If vbwProtector.vbwTraceParameters Then
1037               vbwProtectorParameterString = "()"
1038           End If
1039           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1040       End If
' </VB WATCH>
1041       Dim d As Integer
1042       Dim sSearch As String
1043       Dim ans As Integer
1044       Dim boWriteDataWritten As Boolean


           'check for a serial number
1045       If LenB(txtSN.Text) = 0 Then
1046           MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
' <VB WATCH>
1047       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1048           Exit Sub
1049       End If

           'check to make sure most entries are filled in
1050       If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
1051           MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1052       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1053           Exit Sub
1054       End If
1055       If LenB(txtSalesOrderNumber.Text) = 0 Then
1056           If InStr(1, txtSN.Text, "-") <> 0 Then
1057               txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
1058           End If
1059       End If
1060       If LenB(txtSalesOrderNumber.Text) = 0 Then
1061           MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1062       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1063           Exit Sub
1064       End If

1065       If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
1066           MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1067       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1068           Exit Sub
1069       End If

1070       If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
1071           cmbStatorFill.ListIndex = 0
1072       End If

1073       If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
1074           MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1075       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1076           Exit Sub
1077       End If

1078       If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
1079           MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1080       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1081           Exit Sub
1082       End If


1083       If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
1084           MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1085       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1086           Exit Sub
1087       End If

1088       If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
1089           MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1090       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1091           Exit Sub
1092       End If

1093       If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
1094           MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1095       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1096           Exit Sub
1097       End If

       'check TEMC dropdowns

1098       If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
1099           MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1100       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1101           Exit Sub
1102       End If

1103       If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
1104           MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1105       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1106           Exit Sub
1107       End If

1108       If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
1109           MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1110       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1111           Exit Sub
1112       End If

1113       If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
1114           MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1115       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1116           Exit Sub
1117       End If

1118       If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
1119           MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1120       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1121           Exit Sub
1122       End If

1123       If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
1124           MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1125       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1126           Exit Sub
1127       End If

1128       If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
1129           MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1130       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1131           Exit Sub
1132       End If

1133       If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
1134           MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1135       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1136           Exit Sub
1137       End If

1138       If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
1139           MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1140       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1141           Exit Sub
1142       End If

1143       If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
1144           MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1145       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1146           Exit Sub
1147       End If

1148       If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
1149           MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1150       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1151           Exit Sub
1152       End If

1153       If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
1154           MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1155       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1156           Exit Sub
1157       End If

1158       If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
1159           MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1160       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1161           Exit Sub
1162       End If

1163       If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
1164           MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1165       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1166           Exit Sub
1167       End If

1168       If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
1169           MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1170       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1171           Exit Sub
1172       End If

1173       If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
1174           MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1175       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1176           Exit Sub
1177       End If

1178       If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
1179           MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1180       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1181           Exit Sub
1182       End If

1183       If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
1184           MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1185       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1186           Exit Sub
1187       End If


1188       If Not boFoundPump Then     'if we havent found a pump in the database, add it
1189           rsPumpData.AddNew
1190           boWriteDataWritten = False
1191       Else    'else, find the entry
1192           sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
1193           rsPumpData.MoveFirst
1194           rsPumpData.Find sSearch, , adSearchForward, 1
1195           boWriteDataWritten = True
1196       End If

1197       If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
1198           ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
1199           If ans = vbNo Then
1200               rsPumpData!DataWritten = True
1201               rsPumpData.Update   'update datawritten
' <VB WATCH>
1202       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1203               Exit Sub
1204           End If
1205       End If

1206       rsPumpData!SerialNumber = frmPLCData.txtSN.Text
1207       If LenB(frmPLCData.txtModelNo.Text) <> 0 Then
1208           rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
1209       End If
1210       rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text

1211       If LenB(frmPLCData.txtShpNo.Text) <> 0 Then
1212           rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
1213       End If

1214       If LenB(frmPLCData.txtBilNo.Text) <> 0 Then
1215           rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
1216       End If

1217       rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
1218       If Len(frmPLCData.txtViscosity) <> 0 Then
1219           rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
1220       End If

1221       If LenB(txtSpGr.Text) <> 0 Then
1222           If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
1223               MsgBox "Specific Gravity must be a number."
' <VB WATCH>
1224       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1225               Exit Sub
1226           End If
1227           rsPumpData!SpGr = frmPLCData.txtSpGr.Text
1228       End If
1229       If LenB(txtImpellerDia.Text) <> 0 Then
1230           If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
1231               MsgBox "Impeller Diameter must be a number."
' <VB WATCH>
1232       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1233               Exit Sub
1234           End If
1235           rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
1236       End If

1237       If LenB(txtLiquid.Text) <> 0 Then
1238           rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
1239       End If
1240       If LenB(txtDesignFlow.Text) <> 0 Then
1241           rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
1242       End If
1243       If LenB(txtDesignTDH.Text) <> 0 Then
1244           rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
1245       End If
1246       If LenB(txtRemarks.Text) <> 0 Then
1247           rsPumpData!Remarks = txtRemarks.Text
1248       End If

1249       If optMfr(0).value = True Then
1250           d = cmbMotor.ItemData(cmbMotor.ListIndex)
1251           rsPumpData!Motor = d
1252           d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
1253           rsPumpData!StatorFill = d
1254            d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
1255           rsPumpData!DesignPressure = d
1256           d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
1257           rsPumpData!CirculationPath = d
1258           d = cmbRPM.ItemData(cmbRPM.ListIndex)
1259           rsPumpData!RPM = d
1260           d = cmbModel.ItemData(cmbModel.ListIndex)
1261           rsPumpData!Model = d
1262           d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
1263           rsPumpData!ModelGroup = d
1264       End If
       '   TEMC fields
1265       If optMfr(0).value = False Then
1266           d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
1267           rsPumpData!TEMCAdapter = d

1268           d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
1269           rsPumpData!TEMCAdditions = d

1270           d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
1271           rsPumpData!TEMCcirculation = d

1272           d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
1273           rsPumpData!TEMCDesignpressure = d

1274           d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
1275           rsPumpData!TEMCDivisionType = d

1276           d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
1277           rsPumpData!TEMCImpellerType = d

1278           d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
1279           rsPumpData!TEMCInsulation = d

1280           d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
1281           rsPumpData!TEMCJacketGasket = d

1282           d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
1283           rsPumpData!TEMCMaterials = d

1284           d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
1285           rsPumpData!TEMCModel = d

1286           d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
1287           rsPumpData!TEMCNominalImpSize = d

1288           d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
1289           rsPumpData!TEMCNominalDischargeSize = d

1290           d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
1291           rsPumpData!TEMCNominalSuctionSize = d

1292           d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
1293           rsPumpData!TEMCOtherMotor = d

1294           d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
1295           rsPumpData!TEMCPumpStages = d

1296           d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
1297           rsPumpData!TEMCTRG = d

1298           d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
1299           rsPumpData!TEMCVoltage = d

1300           If LenB(txtTEMCFrameNumber.Text) <> 0 Then
1301               rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
1302           End If
1303       End If

1304       rsPumpData!ChempumpPump = optMfr(0).value

1305       rsPumpData!Approved = False

       'added from TEMC Inspection Report
1306       If Len(txtJobNum.Text) <> 0 Then
1307           rsPumpData!JobNumber = txtJobNum.Text
1308       End If

1309       If Len(txtNoPhases.Text) <> 0 Then
1310           rsPumpData!Phases = txtNoPhases.Text
1311       End If

1312       If Len(txtExpClass.Text) <> 0 Then
1313           rsPumpData!ExpClass = txtExpClass.Text
1314       End If

1315       If Len(txtThermalClass.Text) <> 0 Then
1316           rsPumpData!ThermalClass = txtThermalClass.Text
1317       End If

1318       If LenB(txtNPSHr.Text) <> 0 Then
1319           rsPumpData!NPSHr = Val(txtNPSHr.Text)
1320       End If

1321       If LenB(txtLiquidTemperature.Text) <> 0 Then
1322           rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
1323       End If

1324       If LenB(txtRatedInputPower.Text) <> 0 Then
1325           rsPumpData!RatedOutput = Val(txtRatedInputPower.Text)
1326       End If

1327       If LenB(txtAmps.Text) <> 0 Then
1328           rsPumpData!FLCurrent = Val(txtAmps.Text)
1329       End If



1330       If boWriteDataWritten Then
1331           rsPumpData!DataWritten = True
1332       Else
1333           rsPumpData!DataWritten = False
1334       End If

           'write the data into the database
1335       rsPumpData.Update
1336       boFoundPump = True

           'enter a new test date if it's a new entry
1337       If Not boWriteDataWritten Then


1338           cmdAddNewTestDate_Click
1339       End If
' <VB WATCH>
1340       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1341       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdEnterPumpData_Click"

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
            vbwReportVariable "d", d
            vbwReportVariable "sSearch", sSearch
            vbwReportVariable "ans", ans
            vbwReportVariable "boWriteDataWritten", boWriteDataWritten
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmdEnterTestData_Click()
           ' save the data on the screen to test data at the selected run
' <VB WATCH>
1342       On Error GoTo vbwErrHandler
1343       Const VBWPROCNAME = "frmPLCData.cmdEnterTestData_Click"
1344       If vbwProtector.vbwTraceProc Then
1345           Dim vbwProtectorParameterString As String
1346           If vbwProtector.vbwTraceParameters Then
1347               vbwProtectorParameterString = "()"
1348           End If
1349           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1350       End If
' </VB WATCH>
1351       Dim sSearch As String
1352       Dim ans As Integer

           'if we didn't find the test setup, can't enter test data
1353       If Not boFoundTestSetup Then
1354           MsgBox "You must enter Test Setup Data before entering the Test Data"
' <VB WATCH>
1355       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1356           Exit Sub
1357       End If

           'if we don't find data in the test database, add records
1358       If boFoundTestData = False Then     'add 8 records for 8 tests
1359           AddTestData
1360           rsTestData.MoveFirst
1361       Else        'find the data in the database
1362           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1363           rsTestData.MoveFirst
1364           rsTestData.Filter = sSearch
1365       End If

           'find the desired record from the form
1366       rsTestData.MoveFirst
1367       rsTestData.Move UpDown1.value - 1

1368       If rsTestData!DataWritten = True Then
1369           ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
1370           If ans = vbNo Then
' <VB WATCH>
1371       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1372               Exit Sub
1373           End If
1374       End If

1375       rsEff.MoveFirst
1376       rsEff.Move UpDown1.value - 1

1377       If LenB(txtV1.Text) <> 0 Then
1378           rsTestData!VoltageA = Val(txtV1.Text)
1379       End If

1380       If LenB(txtV2.Text) <> 0 Then
1381           rsTestData!VoltageB = Val(txtV2.Text)
1382       End If

1383       If LenB(txtV3.Text) <> 0 Then
1384           rsTestData!VoltageC = Val(txtV3.Text)
1385       End If

1386       If LenB(txtI1.Text) <> 0 Then
1387           rsTestData!CurrentA = Val(txtI1.Text)
1388       End If

1389       If LenB(txtI2.Text) <> 0 Then
1390           rsTestData!CurrentB = Val(txtI2.Text)
1391       End If

1392       If LenB(txtI3.Text) <> 0 Then
1393           rsTestData!CurrentC = Val(txtI3.Text)
1394       End If

1395       If LenB(txtP1.Text) <> 0 Then
1396           rsTestData!PowerA = Val(txtP1.Text)
1397       End If

1398       If LenB(txtP2.Text) <> 0 Then
1399           rsTestData!PowerB = Val(txtP2.Text)
1400       End If

1401       If LenB(txtP3.Text) <> 0 Then
1402           rsTestData!PowerC = Val(txtP3.Text)
1403       End If

1404       If LenB(txtKW.Text) <> 0 Then
1405           rsTestData!TotalPower = Val(txtKW.Text)
1406       End If

1407       rsTestData!Flow = Val(txtFlowDisplay.Text)
1408       rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
1409       rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
1410       rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

1411       rsTestData!TC1 = Val(txtTC1Display.Text)
1412       rsTestData!TC2 = Val(txtTC2Display.Text)
1413       rsTestData!TC3 = Val(txtTC3Display.Text)
1414       rsTestData!TC4 = Val(txtTC4Display.Text)

1415       rsTestData!CircFlow = Val(txtAI1Display.Text)
1416       rsTestData!RBHTemp = Val(txtAI2Display.Text)
1417       rsTestData!RBHPress = Val(txtAI3Display.Text)
1418       rsTestData!AI4 = Val(txtAI4Display.Text)

1419       rsTestData!ValvePosition = Val(txtValvePosition.Text)
1420       rsTestData!SetPoint = Val(txtSetPoint.Text)

1421       If LenB(txtThrustBal.Text) <> 0 Then
1422           rsTestData!ThrustBalance = txtThrustBal.Text
1423       End If

1424       If LenB(txtVibAx.Text) <> 0 Then
1425           rsTestData!VibrationX = txtVibAx.Text
1426       End If

1427       If LenB(txtVibRad.Text) <> 0 Then
1428           rsTestData!VibrationY = txtVibRad.Text
1429       End If

1430       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1431           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1432       Else
1433           rsTestData!TEMCTRG = 0
1434       End If

1435       If LenB(txtRPM.Text) <> 0 Then
1436           rsTestData!RPM = txtRPM.Text
1437       End If

1438       If LenB(txtTestRemarks.Text) <> 0 Then
1439           rsTestData!Remarks = txtTestRemarks.Text
1440       Else
1441           rsTestData!Remarks = " "
1442       End If

1443       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1444           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1445       End If

1446       If LenB(txtTEMCFrontThrust.Text) <> 0 Then
1447           rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
1448       End If

1449       If LenB(txtTEMCRearThrust.Text) <> 0 Then
1450           rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
1451       End If

1452       If LenB(txtTEMCMomentArm.Text) <> 0 Then
1453           rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
1454       End If

1455       If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
1456           rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
1457       End If

1458       If LenB(txtTEMCViscosity.Text) <> 0 Then
1459           rsTestData!TEMCViscosity = txtTEMCViscosity.Text
1460       End If

1461       If LenB(txtNPSHa.Text) <> 0 Then
1462           rsTestData!NPSHa = txtNPSHa.Text
1463       End If

1464       rsTestData!Approved = False

1465       rsTestData!DataWritten = True

           'update the database
1466       rsTestData.Update

1467       DoEfficiencyCalcs
1468       rsEff.Update

           'update the form
1469       DataGrid1.Refresh
1470       DataGrid2.Refresh
' <VB WATCH>
1471       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1472       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdEnterTestData_Click"

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
            vbwReportVariable "sSearch", sSearch
            vbwReportVariable "ans", ans
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmdEnterTestSetupData_Click()
           'save the data on the screen to testsetupdata
' <VB WATCH>
1473       On Error GoTo vbwErrHandler
1474       Const VBWPROCNAME = "frmPLCData.cmdEnterTestSetupData_Click"
1475       If vbwProtector.vbwTraceProc Then
1476           Dim vbwProtectorParameterString As String
1477           If vbwProtector.vbwTraceParameters Then
1478               vbwProtectorParameterString = "()"
1479           End If
1480           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1481       End If
' </VB WATCH>
1482       Dim I As Integer
1483       Dim d As Integer
1484       Dim sSearch As String
1485       Dim ans As Integer
1486       Dim boWriteDataWritten As Boolean

           'check for a serial number
1487       If LenB(txtSN.Text) = 0 Then
1488           MsgBox "You must have a Serial Number to enter data."
' <VB WATCH>
1489       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1490           Exit Sub
1491       End If

1492       If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
1493           rsTestSetup.AddNew
1494           cmbTestDate.AddItem Now
1495           cmbTestDate.ListIndex = cmbTestDate.NewIndex
1496           cmdAddNewBalanceHoles.Visible = True
1497           boFoundTestSetup = True
1498           boWriteDataWritten = False
1499           rsTestSetup!DataWritten = False
1500       Else    'find the record and display
1501           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1502           rsTestSetup.MoveFirst
1503           rsTestSetup.Filter = sSearch
1504           If Not boCanApprove Then
       '            cmdAddNewBalanceHoles.Visible = False
1505           End If
1506           boWriteDataWritten = True
1507       End If

1508       If rsTestSetup!DataWritten = True Then
1509           ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
1510           If ans = vbNo Then
' <VB WATCH>
1511       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1512               Exit Sub
1513           End If
1514       End If

1515       rsTestSetup!SerialNumber = txtSN
1516       rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

1517       If LenB(txtFlowmeterID.Text) <> 0 Then
1518           rsTestSetup!flowmeterid = txtFlowmeterID
1519       Else
1520           rsTestSetup!flowmeterid = vbNullString
1521       End If

       '    I = cmbFlowMeter.ListIndex
       '    If I = -1 Then
       '        d = -1
       '    Else
       '        d = cmbFlowMeter.ItemData(I)
       '        rsTestSetup!flowmeterid = d
       '    End If



1522       If LenB(txtSuctionID.Text) <> 0 Then
1523           rsTestSetup!suctionid = txtSuctionID
1524       Else
1525           rsTestSetup!suctionid = vbNullString
1526       End If
1527       If LenB(txtDischargeID.Text) <> 0 Then
1528           rsTestSetup!dischid = txtDischargeID
1529       Else
1530           rsTestSetup!dischid = vbNullString
1531       End If
1532       If LenB(txtTemperatureID.Text) <> 0 Then
1533           rsTestSetup!temperatureid = txtTemperatureID
1534       Else
1535           rsTestSetup!temperatureid = vbNullString
1536       End If
1537       If LenB(txtMagflowID.Text) <> 0 Then
1538           rsTestSetup!magflowid = txtMagflowID
1539       Else
1540           rsTestSetup!magflowid = vbNullString
1541       End If
1542       If LenB(txtHDCor.Text) <> 0 Then
1543           rsTestSetup!HDCor = txtHDCor
1544       Else
1545           rsTestSetup!HDCor = 0
1546       End If
1547       If LenB(txtKWMult.Text) <> 0 Then
1548           rsTestSetup!kwmult = txtKWMult
1549       Else
1550           rsTestSetup!kwmult = 1
1551       End If
1552       If LenB(txtWho.Text) <> 0 Then
1553           rsTestSetup!who = txtWho
1554       Else
1555           rsTestSetup!who = vbNullString
1556       End If
1557       If LenB(txtRMA.Text) <> 0 Then
1558           rsTestSetup!RMA = txtRMA
1559       Else
1560           rsTestSetup!RMA = vbNullString
1561       End If
1562       If LenB(frmPLCData.txtDischHeight) <> 0 Then
1563           rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
1564       Else
1565           rsTestSetup!DischargeGageHeight = 0
1566       End If
1567       If LenB(frmPLCData.txtSuctHeight) <> 0 Then
1568           rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
1569       Else
1570           rsTestSetup!SuctionGageHeight = 0
1571       End If
1572       If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
1573           rsTestSetup!Remarks = txtTestSetupRemarks.Text
1574       Else
1575           rsTestSetup!Remarks = vbNullString
1576       End If
1577       If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
1578           rsTestSetup!VFDFrequency = txtVFDFreq.Text
1579       Else
1580           rsTestSetup!VFDFrequency = 0
1581       End If

1582       I = cmbOrificeNumber.ListIndex
1583       If I = -1 Then
1584           d = 18      'entry for None
1585       Else
1586           d = cmbOrificeNumber.ItemData(I)
1587       End If
1588       rsTestSetup!orificenumber = d

1589       If LenB(txtEndPlay.Text) <> 0 Then
1590           rsTestSetup!EndPlay = Val(frmPLCData.txtEndPlay.Text)
1591       Else
1592           rsTestSetup!EndPlay = 0
1593       End If

1594       If LenB(txtGGap.Text) <> 0 Then
1595           rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
1596       Else
1597           rsTestSetup!GGAP = 0
1598       End If

1599       If LenB(txtOtherMods.Text) <> 0 Then
1600           rsTestSetup!OtherMods = txtOtherMods.Text
1601       Else
1602           rsTestSetup!OtherMods = vbNullString
1603       End If

1604       rsTestSetup!Approved = False

1605       I = cmbLoopNumber.ListIndex
1606       If I = -1 Then
1607           d = -1
1608       Else
1609           d = cmbLoopNumber.ItemData(I)
1610           rsTestSetup!loopnumber = d
1611       End If

1612       I = cmbSuctDia.ListIndex
1613       If I = -1 Then
1614           d = -1
1615       Else
1616           d = cmbSuctDia.ItemData(I)
1617           rsTestSetup!SuctDiam = d
1618       End If

1619       I = cmbDischDia.ListIndex
1620       If I = -1 Then
1621           d = -1
1622       Else
1623           d = cmbDischDia.ItemData(I)
1624           rsTestSetup!DischDiam = d
1625       End If

1626       I = cmbTachID.ListIndex
1627       If I = -1 Then
1628           d = -1
1629       Else
1630           d = cmbTachID.ItemData(I)
1631           rsTestSetup!tachid = d
1632       End If

1633       I = cmbAnalyzerNo.ListIndex
1634       If I = -1 Then
1635           d = -1
1636       Else
1637           d = cmbAnalyzerNo.ItemData(I)
1638           rsTestSetup!analyzerno = d
1639       End If

1640       I = cmbTestSpec.ListIndex
1641       If I = -1 Then
1642           d = 0
1643       Else
1644           d = cmbTestSpec.ItemData(I)
1645       End If
1646       rsTestSetup!testspec = d

1647       I = cmbVoltage.ListIndex
1648       If I = -1 Then
1649           d = -1
1650       Else
1651           d = cmbVoltage.ItemData(I)
1652           rsTestSetup!Voltage = d
1653       End If

1654       I = cmbFrequency.ListIndex
1655       If I = -1 Then
1656           d = -1
1657       Else
1658           d = cmbFrequency.ItemData(I)
1659           rsTestSetup!Frequency = d
1660       End If

1661       I = cmbMounting.ListIndex
1662       If I = -1 Then
1663           d = -1
1664       Else
1665           d = cmbMounting.ItemData(I)
1666           rsTestSetup!Mounting = d
1667       End If

1668       I = cmbPLCNo.ListIndex
1669       If I = -1 Then
1670           d = -1
1671       Else
1672           d = cmbPLCNo.ItemData(I)
1673           rsTestSetup!PLCNo = d
1674       End If

1675       rsTestSetup!ImpFeathered = chkFeathered.value

1676       If chkTrimmed.value = 1 Then
1677           rsTestSetup!ImpTrimmed = Val(txtImpTrim)
1678       Else
1679           rsTestSetup!ImpTrimmed = 0
1680       End If
1681       chkTrimmed_Click

1682       If chkOrifice.value = 1 Then
1683           rsTestSetup!PumpDischOrifice = Val(txtOrifice)
1684       Else
1685           rsTestSetup!PumpDischOrifice = 0
1686       End If
1687       chkOrifice_Click

1688       If chkCircOrifice.value = 1 Then
1689           rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
1690       Else
1691           rsTestSetup!CircFlowOrifice = 0
1692       End If
1693       chkCircOrifice_Click

1694       If Me.chkAddedDiodes.value = 1 Then
1695           rsTestSetup!NoOfTRGDiodes = Val(Me.txtNoOfDiodes.Text)
1696       Else
1697           rsTestSetup!NoOfTRGDiodes = 0
1698       End If
1699       chkAddedDiodes_Click

1700       chkBalanceHoles_Click

1701       If chkNPSH.value = 1 Then
1702           txtNPSHFile.Visible = True
1703           rsTestSetup!NPSHFile = txtNPSHFile
1704       Else
1705           rsTestSetup!NPSHFile = vbNullString
1706           txtNPSHFile.Visible = False
1707       End If

1708       If chkPictures.value = 1 Then
1709           txtPicturesFile.Visible = True
1710           rsTestSetup!PictureFile = txtPicturesFile
1711       Else
1712           rsTestSetup!PictureFile = vbNullString
1713           txtPicturesFile.Visible = False
1714       End If

1715       If chkVibration.value = 1 Then
1716           txtVibrationFile.Visible = True
1717           rsTestSetup!VibrationFile = txtVibrationFile
1718       Else
1719           rsTestSetup!VibrationFile = vbNullString
1720           txtVibrationFile.Visible = False
1721       End If

1722       If boWriteDataWritten Then
1723           rsTestSetup!DataWritten = True
1724       Else
1725           rsTestSetup!DataWritten = False
1726       End If

           'for TEMC Inspection Report
1727       If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
1728           rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
1729       Else
1730           rsTestSetup!InsulationMeggerVolts = ""
1731       End If

1732       If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
1733           rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
1734       Else
1735           rsTestSetup!InsulationMegOhms = ""
1736       End If

1737       If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
1738           rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
1739       Else
1740           rsTestSetup!DielectricVolts = ""
1741       End If

1742       If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
1743           rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
1744       Else
1745           rsTestSetup!DielectricTime = ""
1746       End If

1747       If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
1748           rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
1749       Else
1750           rsTestSetup!HydrostaticValue = ""
1751       End If

1752       If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
1753           rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
1754       Else
1755           rsTestSetup!HydrostaticTime = ""
1756       End If

1757       If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
1758           rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
1759       Else
1760           rsTestSetup!PneumaticValue = ""
1761       End If

1762       If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
1763           rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
1764       Else
1765           rsTestSetup!PneumaticTime = ""
1766       End If

1767       I = cmbTestAndInspection(0).ListIndex
1768       If I = -1 Then
1769           rsTestSetup!HydrostaticUnits = ""
1770       Else
1771           rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
1772       End If


1773       I = cmbTestAndInspection(1).ListIndex
1774       If I = -1 Then
1775           rsTestSetup!PneumaticUnits = ""
1776       Else
1777           rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
1778       End If

           'use abs to convert from 1 and 0 to boolean
1779       rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
1780       rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
1781       rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
1782       rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
1783       rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
1784       rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
1785       rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
1786       rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
1787       rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
1788       rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
1789       rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
1790       rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
1791       rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
1792       rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
1793       rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)



           'update the database
1794       rsTestSetup.Update

1795       If boFoundTestData = False Then     'add 8 records for 8 tests
1796           AddTestData
1797       End If

1798       rsTestSetup.Filter = vbNullString
' <VB WATCH>
1799       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1800       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdEnterTestSetupData_Click"

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
            vbwReportVariable "d", d
            vbwReportVariable "sSearch", sSearch
            vbwReportVariable "ans", ans
            vbwReportVariable "boWriteDataWritten", boWriteDataWritten
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmdExit_Click()
' <VB WATCH>
1801       On Error GoTo vbwErrHandler
1802       Const VBWPROCNAME = "frmPLCData.cmdExit_Click"
1803       If vbwProtector.vbwTraceProc Then
1804           Dim vbwProtectorParameterString As String
1805           If vbwProtector.vbwTraceParameters Then
1806               vbwProtectorParameterString = "()"
1807           End If
1808           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1809       End If
' </VB WATCH>
1810       End
' <VB WATCH>
1811       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1812       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdExit_Click"

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

Private Sub cmdFindMagtrols_Click()
' <VB WATCH>
1813       On Error GoTo vbwErrHandler
1814       Const VBWPROCNAME = "frmPLCData.cmdFindMagtrols_Click"
1815       If vbwProtector.vbwTraceProc Then
1816           Dim vbwProtectorParameterString As String
1817           If vbwProtector.vbwTraceParameters Then
1818               vbwProtectorParameterString = "()"
1819           End If
1820           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1821       End If
' </VB WATCH>
1822       FindMagtrols
' <VB WATCH>
1823       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1824       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdFindMagtrols_Click"

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

Private Sub cmdFindPump_Click()
           ' find the pump whose sn is shown
' <VB WATCH>
1825       On Error GoTo vbwErrHandler
1826       Const VBWPROCNAME = "frmPLCData.cmdFindPump_Click"
1827       If vbwProtector.vbwTraceProc Then
1828           Dim vbwProtectorParameterString As String
1829           If vbwProtector.vbwTraceParameters Then
1830               vbwProtectorParameterString = "()"
1831           End If
1832           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1833       End If
' </VB WATCH>

       '    Dim m As SNRecord
       '    m = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)

1834       Dim sAns As String
1835       Dim sSO As String
1836       Dim sParam As String
1837       Dim sName As String

1838       Dim I As Integer

           'set TC and AI labels with default values
1839       txtTitle(0).Text = "TC 1"
1840       txtTitle(1).Text = "(F)"
1841       txtTitle(2).Text = "TC 2"
1842       txtTitle(3).Text = "(F)"
1843       txtTitle(4).Text = "TC 3"
1844       txtTitle(5).Text = "(F)"
1845       txtTitle(6).Text = "TC 4"
1846       txtTitle(7).Text = "(F)"
1847       txtTitle(20).Text = "Circ Flow"
1848       txtTitle(21).Text = "(GPM)"
1849       txtTitle(22).Text = "RBH Temp"
1850       txtTitle(23).Text = "(F)"
1851       txtTitle(24).Text = "RBH Press"
1852       txtTitle(25).Text = "(psig)"
1853       txtTitle(26).Text = "AI 4"
1854       txtTitle(27).Text = ""


1855       For I = 0 To 7
1856           lblAutoMan(I).Caption = "Auto"
1857       Next I

1858       txtFlowDisplay.Enabled = False
1859       txtSuctionDisplay.Enabled = False
1860       txtDischargeDisplay.Enabled = False
1861       txtTemperatureDisplay.Enabled = False
1862       txtAI1Display.Enabled = False
1863       txtAI2Display.Enabled = False
1864       txtAI3Display.Enabled = False
1865       txtAI4Display.Enabled = False


1866       cmdFindPump.Default = False

           'set all found booleans to false
1867       boUsingHP = False
1868       boFoundPump = False
1869       boPumpIsApproved = False
1870       boFoundTestSetup = False
1871       boFoundTestData = False


           'get rid of all test dates in combo box
1872       For I = cmbTestDate.ListCount - 1 To 0 Step -1
1873           cmbTestDate.RemoveItem 0
1874       Next I

1875       rsTestData.Filter = "SerialNumber = ''"

1876       DataGrid2.ClearFields
1877       ClearEff

1878       If rsPumpData.State = adStateOpen Then
1879           rsPumpData.Close
1880       End If

           'find the pump listed in the Serial Number text box
1881       qyPumpData.ActiveConnection = cnPumpData
1882       qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                      txtSN.Text & "'))"
1883       rsPumpData.CursorType = adOpenStatic
1884       rsPumpData.CursorLocation = adUseClient
1885       rsPumpData.Index = "SerialNumber"
1886       rsPumpData.Open qyPumpData

1887       If rsPumpData.BOF = True And rsPumpData.EOF = True Then
               'if the bof=eof, we have an empty recordset
1888           boFoundPump = False
1889       Else
               'we found it
1890           boFoundPump = True
1891       End If

       '    If InStr(1, txtSN.Text, "-") = 0 Then
       '        sAns = MsgBox("There is no dash in the Serial Number.  Please add a dash and try again.", vbOKOnly, "No dash in Serial Number")
       '        Exit Sub
       '    End If


1892       If boFoundPump = False Then
               'not found in either database, try HP?
1893           sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
1894           If sAns = vbNo Then     'new pump - don't get data from HP
1895               boUsingEpicor = False
1896           Else
1897               boUsingEpicor = True
1898               boUsingHP = False
1899           End If
       '        If boUsingEpicor = False Then
       '            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
       '            If sAns = vbNo Then     'new pump - don't get data from HP
1900                    boUsingHP = False
       '            Else
       '                boUsingHP = True
       '            End If
       '        End If
1901           EnablePumpDataControls
1902           EnableTestSetupDataControls
1903           EnableTestDataControls
       '        BlankData               'clear any data on the screen
1904           cmdAddNewBalanceHoles.Visible = True

1905       End If

1906       If boFoundPump = True Then    'found the pump
1907           If rsPumpData.Fields("Approved") = True Then
1908               DisablePumpDataControls                         'if it's in the real database, don't allow changes here
1909               boPumpIsApproved = True
1910               lblPumpApproved.Visible = True
1911               If boCanApprove Then
1912                   cmdApprovePump.Caption = "Unapprove this pump"
1913               End If
1914               frmPLCData.cmdApproveTestDate.Enabled = True
1915           Else
1916               EnablePumpDataControls                          'it's in the temp database, allow changes
1917               boPumpIsApproved = False
1918               boTestDateIsApproved = False
1919               lblPumpApproved.Visible = False
1920               If boCanApprove Then
1921                   cmdApprovePump.Caption = "Approve this pump"
1922               End If
1923               cmdApproveTestDate.Caption = "You Must Approve Pump First"
1924               frmPLCData.cmdApproveTestDate.Enabled = False
1925           End If

               'found the pump, show the data
1926           txtModelNo.Text = rsPumpData.Fields("ModelNumber")
1927           frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

1928           If rsPumpData.Fields("ChempumpPump") = True Then
1929               SetCombo cmbMotor, "Motor", rsPumpData
1930               SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
1931               SetCombo cmbRPM, "RPM", rsPumpData
1932               SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
1933               SetCombo cmbStatorFill, "StatorFill", rsPumpData
1934               SetCombo cmbModel, "Model", rsPumpData
1935               SetCombo cmbModelGroup, "ModelGroup", rsPumpData
1936               RatedKW = 999
1937           End If

               'set the TEMC data
1938           If rsPumpData.Fields("ChempumpPump") = False Then
1939               SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
1940               SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
1941               SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
1942               SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
1943               SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
1944               SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
1945               SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
1946               SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
1947               SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
1948               SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
1949               SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
1950               SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
1951               SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
1952               SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
1953               SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
1954               SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
1955               SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
1956           End If

               'write ship to and bill to info
1957           If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
1958               txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
1959           Else
1960               txtShpNo.Text = vbNullString
1961           End If

1962           If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
1963               txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
1964           Else
1965               txtBilNo.Text = vbNullString
1966           End If

1967           sName = "ImpellerDia"
1968           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1969               sParam = rsPumpData.Fields(sName)
1970           Else
1971               sParam = vbNullString
1972           End If
1973           txtImpellerDia.Text = sParam

1974           sName = "DesignFlow"
1975           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1976               sParam = rsPumpData.Fields(sName)
1977           Else
1978               sParam = vbNullString
1979           End If
1980           txtDesignFlow.Text = sParam

1981           sName = "DesignTDH"
1982           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1983               sParam = rsPumpData.Fields(sName)
1984           Else
1985               sParam = vbNullString
1986           End If
1987           txtDesignTDH.Text = sParam

1988           sName = "SpGr"
1989           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1990               sParam = rsPumpData.Fields(sName)
1991           Else
1992               sParam = vbNullString
1993           End If
1994           txtSpGr.Text = sParam

1995           sName = "Remarks"
1996           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1997               sParam = rsPumpData.Fields(sName)
1998           Else
1999               sParam = vbNullString
2000           End If
2001           txtRemarks.Text = sParam

2002           sName = "SalesOrderNumber"
2003           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2004               sParam = rsPumpData.Fields(sName)
2005           Else
2006               sParam = vbNullString
2007           End If
2008           txtSalesOrderNumber.Text = sParam

2009           sName = "ApplicationFluid"
2010           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2011               sParam = rsPumpData.Fields(sName)
2012           Else
2013               sParam = vbNullString
2014           End If
2015           txtLiquid.Text = sParam

2016           sName = "NPSHFile"
2017           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2018               sParam = rsPumpData.Fields(sName)
2019           Else
2020               sParam = vbNullString
2021           End If
2022           txtNPSHFileLocation.Text = sParam

2023           sName = "ApplicationViscosity"
2024           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2025               sParam = Format(rsPumpData.Fields(sName), "#0.00")
2026           Else
2027               sParam = vbNullString
2028           End If
2029           txtViscosity.Text = sParam

       'added from TEMC Inspection Report
2030           sName = "JobNumber"
2031           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2032               sParam = rsPumpData.Fields(sName)
2033           Else
2034               sParam = ""
2035           End If
2036           txtJobNum.Text = sParam

2037           sName = "Phases"
2038           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2039               sParam = rsPumpData.Fields(sName)
2040           Else
2041               sParam = vbNullString
2042           End If
2043           txtNoPhases.Text = sParam

2044           sName = "ThermalClass"
2045           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2046               sParam = rsPumpData.Fields(sName)
2047           Else
2048               sParam = vbNullString
2049           End If
2050           txtThermalClass.Text = sParam

2051           sName = "ExpClass"
2052           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2053               sParam = rsPumpData.Fields(sName)
2054           Else
2055               sParam = vbNullString
2056           End If
2057           txtExpClass.Text = sParam

2058           sName = "NPSHr"
2059           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2060               sParam = rsPumpData.Fields(sName)
2061           Else
2062               sParam = vbNullString
2063           End If
2064           txtNPSHr.Text = sParam

2065           sName = "LiquidTemperature"
2066           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2067               sParam = rsPumpData.Fields(sName)
2068           Else
2069               sParam = vbNullString
2070           End If
2071           txtLiquidTemperature.Text = sParam

2072           sName = "RatedOutput"
2073           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2074               sParam = rsPumpData.Fields(sName)
2075           Else
2076               sParam = vbNullString
2077           End If
2078           txtRatedInputPower.Text = sParam

2079           sName = "FLCurrent"
2080           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2081               sParam = rsPumpData.Fields(sName)
2082           Else
2083               sParam = vbNullString
2084           End If
2085           txtAmps.Text = sParam

2086           sName = "TEMCFrameNumber"
2087           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2088               sParam = rsPumpData.Fields(sName)
2089           Else
2090               sParam = vbNullString
2091           End If
2092           txtTEMCFrameNumber.Text = sParam

2093           optMfr(0).value = rsPumpData.Fields("ChempumpPump")
2094           optMfr(1).value = Not optMfr(0).value

               'select the testsetup data
2095           qyTestSetup.ActiveConnection = cnPumpData
2096           qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                      txtSN.Text & "')) ORDER BY Date"
       '        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

2097           With rsTestSetup
2098               If .State = adStateOpen Then
2099                   .Close
2100               End If
2101               .CursorLocation = adUseClient
2102               .CursorType = adOpenStatic
2103               .Index = "FindData"
2104               .Open qyTestSetup
2105           End With


               'add the selection of dates to the Test Date combo box
2106           If rsTestSetup.RecordCount <> 0 Then
2107               For I = 0 To cmbTestDate.ListCount - 1
2108                   cmbTestDate.RemoveItem 0
2109               Next I
2110               rsTestSetup.MoveFirst
2111               For I = 1 To rsTestSetup.RecordCount
2112                   cmbTestDate.AddItem rsTestSetup.Fields("Date")
2113                   rsTestSetup.MoveNext
2114               Next I
2115               rsTestSetup.MoveFirst
2116               boFoundTestSetup = True

2117               If rsTestSetup.Fields("Approved") = True Then
2118                   DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
2119                   boTestDateIsApproved = True
2120                   lblTestDateApproved.Visible = True
2121                   If boCanApprove Then
2122                       cmdApproveTestDate.Caption = "Unapprove this Test Date"
2123                   End If
2124               Else
2125                   EnableTestSetupDataControls                          'it's in the temp database, allow changes
2126                   lblTestDateApproved.Visible = False
2127                   If boCanApprove Then
2128                       cmdApproveTestDate.Caption = "Approve this Test Date"
2129                   End If
2130               End If
2131               cmbTestDate.ListIndex = 0
2132           Else
2133               MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
2134               boFoundTestSetup = False        'didn't find any data
2135               boFoundTestData = False
2136               cmbTestDate.AddItem Date        'load with today
2137               cmbTestDate.ListIndex = 0       'show the entry
2138               EnableTestSetupDataControls
2139               txtTestRemarks.Text = ""
2140               txtVibAx.Text = ""
2141               txtVibRad.Text = ""
2142               txtThrustBal.Text = ""
2143               txtTEMCTRGReading.Text = ""
2144               txtTEMCFrontThrust.Text = ""
2145               txtTEMCRearThrust.Text = ""
' <VB WATCH>
2146       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2147               Exit Sub
2148           End If

2149           If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
2150           End If
' <VB WATCH>
2151       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2152           Exit Sub
2153       End If

           'hp interface stuff
2154       If boUsingHP = True Then
2155           If InStr(1, txtSN.Text, "-") = 0 Then
2156               MsgBox "Please check the Serial Number.  There doesn't seem to be a -", vbOKOnly
' <VB WATCH>
2157       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2158               Exit Sub
2159           Else
2160               sSO = Left$(txtSN.Text, 7)       'look for the sales order
2161               If Len(sSO) <> 7 Then
2162                   MsgBox "Please check the Serial Number.  There doesn't seem to be 7 digits before the -", vbOKOnly
' <VB WATCH>
2163       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2164                   Exit Sub
2165               End If
2166           End If

2167           If Not cnHPOpen Then
2168               frmConnectingToPLC.Show
2169               DoEvents
2170               With cnHP
2171                   .ConnectionString = sHPDataBaseName
2172                   .CommandTimeout = 10
2173                   .Open
2174               End With
2175               cnHPOpen = True
2176               frmConnectingToPLC.Hide
2177               DoEvents
2178           End If

2179           txtSalesOrderNumber.Text = sSO

2180           GetDetail sSO

2181           SearchSalesOrder txtSalesOrderNumber.Text

2182           frmPLCData.txtShpNo.Text = strShipTo
2183           frmPLCData.txtBilNo.Text = strBillTo

2184           rsHPDetail.Filter = 0

2185           If rsHPDetail.BOF = True And rsHPDetail.EOF = True Then
2186               MsgBox "Not found on HP.  You may enter data manually."
2187           Else
2188               frmPLCData.txtModelNo = strModelNo(1, intLineNo)
2189               frmPLCData.txtDesignTDH.Text = strTDH(1, intLineNo)
2190               frmPLCData.txtSpGr.Text = strSpGr(1, intLineNo)
2191               frmPLCData.txtImpellerDia.Text = strImpellers(1, intLineNo)
2192               frmPLCData.txtDesignFlow.Text = strCapacity(1, intLineNo)

2193               For I = 0 To cmbStatorFill.ListCount - 1
2194                   If InStr(1, UCase$(strStatorFill(1, intLineNo)), UCase$(cmbStatorFill.List(I))) <> 0 Then
2195                       cmbStatorFill.ListIndex = I
2196                       Exit For
2197                   End If
2198               Next I

2199               For I = 0 To cmbDesignPressure.ListCount - 1
2200                   If InStr(1, strDesignPress(1, intLineNo), cmbDesignPressure.List(I)) <> 0 Then
2201                       cmbDesignPressure.ListIndex = I
2202                       Exit For
2203                   End If
2204               Next I

2205               I = InStr(strVoltage(1, intLineNo), "VOLT")
2206               sName = strFindTheNumber(strVoltage(1, intLineNo), I)

2207               For I = 0 To cmbVoltage.ListCount - 1
2208                   If InStr(1, sName, cmbVoltage.List(I)) <> 0 Then
2209                       cmbVoltage.ListIndex = I
2210                       Exit For
2211                   End If
2212               Next I

2213               I = InStr(strVoltage(1, intLineNo), "CY")
2214               sName = strFindTheNumber(strVoltage(1, intLineNo), I)

2215               For I = 0 To cmbFrequency.ListCount - 1
2216                   If InStr(1, cmbFrequency.List(I), sName) <> 0 Then
2217                       cmbFrequency.ListIndex = I
2218                       Exit For
2219                   End If
2220               Next I

2221               For I = 0 To cmbRPM.ListCount - 1
2222                   If InStr(1, strRPM(1, intLineNo), cmbRPM.List(I)) <> 0 Then
2223                       cmbRPM.ListIndex = I
2224                       Exit For
2225                   End If
2226               Next I

2227               For I = 0 To cmbSuctDia.ListCount - 1
2228                   If InStr(1, strSuctFlg(1, intLineNo), cmbSuctDia.List(I)) <> 0 Then
2229                       cmbSuctDia.ListIndex = I
2230                       Exit For
2231                   End If
2232               Next I

2233               For I = 0 To cmbDischDia.ListCount - 1
2234                   If InStr(1, strDischFlg(1, intLineNo), cmbDischDia.List(I)) <> 0 Then
2235                       cmbDischDia.ListIndex = I
2236                       Exit For
2237                   End If
2238               Next I

2239               For I = 0 To cmbTestSpec.ListCount - 1
2240                   If InStr(1, strTestProcedure(1, intLineNo), cmbTestSpec.List(I)) <> 0 Then
2241                       cmbTestSpec.ListIndex = I
2242                       Exit For
2243                   End If
2244               Next I

2245               rsHPDetail.MoveFirst
2246               Load FrmSODetails
2247               FrmSODetails.Show
2248               FrmSODetails.txtSOData.Text = vbNullString
2249           End If

2250           Dim intLastLineNo As Integer
2251           Dim vFilter As Variant

2252           intLastLineNo = 0

2253           If rsHPLineNo.State = adStateOpen Then
2254               Do While Not rsHPDetail.EOF
2255                   If Int(Val(rsHPDetail.Fields(1))) <> intLastLineNo Then
2256                       intLastLineNo = Val(rsHPDetail.Fields(1))
2257                       FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & vbCrLf & "Line No. = " & intLastLineNo & " Quan = "
2258                       vFilter = "LINE = '" & str$(intLastLineNo) & "'"
2259                       rsHPLineNo.Filter = vFilter
2260                       If rsHPLineNo.BOF = True And rsHPLineNo.EOF = True Then
2261                           FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & vbCrLf
2262                       Else
2263                           FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & rsHPLineNo.Fields(2) & vbCrLf
2264                       End If
2265                       rsHPLineNo.Filter = 0
2266                   End If
2267                   FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & "   " & rsHPDetail.Fields(2) & vbCrLf
2268                   rsHPDetail.MoveNext
2269               Loop
2270           End If
2271       End If

2272       If boUsingEpicor = True Then
2273           Dim MyRecord As SNRecord
       '            I = InStr(1, txtSN.Text, "-")
       '            If I > 0 Then
2274               MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
       '            End If
2275           If MyRecord.SONumber = "" Then
2276               MsgBox ("Not found in Epicor")
' <VB WATCH>
2277       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2278               Exit Sub
2279           End If
2280           txtSalesOrderNumber.Text = MyRecord.SONumber
2281           txtLineNumber.Text = MyRecord.SOLine
2282           txtBilNo.Text = MyRecord.Customer
2283           If MyRecord.ShipTo = "" Then
2284               txtShpNo.Text = MyRecord.Customer
2285           Else
2286               txtShpNo.Text = MyRecord.ShipTo
2287           End If
2288           txtModelNo.Text = MyRecord.ModelNo
2289           txtModelNo_Change
2290           txtDesignTDH.Text = MyRecord.TDH
2291           txtSpGr.Text = MyRecord.SpGr
2292           txtImpellerDia.Text = MyRecord.ImpellerDiameter
2293           txtDesignFlow.Text = MyRecord.Flow
2294           txtNoPhases.Text = MyRecord.Phases
2295           txtNPSHr.Text = MyRecord.NPSHr
2296           txtRatedInputPower.Text = MyRecord.RatedInputPower
2297           txtAmps.Text = MyRecord.FLCurrent
2298           txtThermalClass.Text = MyRecord.ThermalClass
2299           txtViscosity.Text = MyRecord.Viscosity
2300           txtExpClass.Text = MyRecord.ExpClass
2301           txtLiquidTemperature.Text = MyRecord.LiquidTemp
2302           txtLiquid.Text = MyRecord.Fluid
2303           txtJobNum.Text = MyRecord.JobNumber

2304           For I = 0 To cmbStatorFill.ListCount - 1
2305               If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
2306                   cmbStatorFill.ListIndex = I
2307                   Exit For
2308               End If
2309           Next I

2310           For I = 0 To cmbCirculationPath.ListCount - 1
2311               If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
2312                   cmbCirculationPath.ListIndex = I
2313                   Exit For
2314               End If
2315           Next I

2316           For I = 0 To cmbDesignPressure.ListCount - 1
2317               If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
2318                   cmbDesignPressure.ListIndex = I
2319                   Exit For
2320               End If
2321           Next I

2322           For I = 0 To cmbVoltage.ListCount - 1
2323               If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
2324                   cmbVoltage.ListIndex = I
2325                   Exit For
2326               End If
2327           Next I

2328           For I = 0 To cmbFrequency.ListCount - 1
2329               If InStr(1, MyRecord.Frequency, sName) <> 0 Then
2330                   cmbFrequency.ListIndex = I
2331                   Exit For
2332               End If
2333           Next I

2334           For I = 0 To cmbRPM.ListCount - 1
2335               If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
2336                   cmbRPM.ListIndex = I
2337                   Exit For
2338               End If
2339           Next I

2340           For I = 0 To cmbSuctDia.ListCount - 1
2341               If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
2342                   cmbSuctDia.ListIndex = I
2343                   Exit For
2344               End If
2345           Next I

2346           For I = 0 To cmbDischDia.ListCount - 1
2347               If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
2348                   cmbDischDia.ListIndex = I
2349                   Exit For
2350               End If
2351           Next I

2352           For I = 0 To cmbTestSpec.ListCount - 1
2353               If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
2354                   cmbTestSpec.ListIndex = I
2355                   Exit For
2356               End If
2357           Next I

2358           For I = 0 To cmbMotor.ListCount - 1
2359               If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
2360                   cmbMotor.ListIndex = I
2361                   Exit For
2362               End If
2363           Next I


2364       End If

' <VB WATCH>
2365       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2366       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdFindPump_Click"

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
            vbwReportVariable "sAns", sAns
            vbwReportVariable "sSO", sSO
            vbwReportVariable "sParam", sParam
            vbwReportVariable "sName", sName
            vbwReportVariable "I", I
            vbwReportVariable "intLastLineNo", intLastLineNo
            vbwReportVariable "vFilter", vFilter
            vbwReport_EpicorRoutines_SNRecord "MyRecord", MyRecord
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdModifyBalanceHoleData_Click()
' <VB WATCH>
2367       On Error GoTo vbwErrHandler
2368       Const VBWPROCNAME = "frmPLCData.cmdModifyBalanceHoleData_Click"
2369       If vbwProtector.vbwTraceProc Then
2370           Dim vbwProtectorParameterString As String
2371           If vbwProtector.vbwTraceParameters Then
2372               vbwProtectorParameterString = "()"
2373           End If
2374           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2375       End If
' </VB WATCH>
2376       Dim strInput As String
2377       Dim I As Integer
2378       Dim sNumber As Integer
2379       Dim sDia As String
2380       Dim sBC As String

2381       cmdModifyBalanceHoleData.Visible = False

2382       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2383           cmdModifyBalanceHoleData.Visible = False
' <VB WATCH>
2384       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2385           Exit Sub
2386       End If

2387       rsBalanceHoles.MoveFirst
2388       rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

2389       sNumber = rsBalanceHoles!Number
2390       If rsBalanceHoles!diameter = 99 Then
2391           sDia = "Slot"
2392       Else
2393           sDia = str(rsBalanceHoles!diameter)
2394       End If
2395       If rsBalanceHoles!boltcircle = 99 Then
2396           sBC = "Unknown"
2397       Else
2398           sBC = str(rsBalanceHoles!boltcircle)
2399       End If


           'get the data for the balance holes
2400       strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
2401       If strInput = "" Then
2402           GoTo DeleteIt
2403       End If
2404       sNumber = CInt(strInput)
2405       If Val(sNumber) = 0 Then
2406           GoTo DeleteIt
2407       End If

2408       strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
2409       If strInput <> "" Then
2410           If UCase(strInput) = "SLOT" Then
2411               strInput = 99
2412           End If
2413           sDia = CSng(strInput)
2414       Else
2415           GoTo CancelPressed
2416       End If

2417       strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
2418       If strInput <> "" Then
2419           If UCase(strInput) = "UNKNOWN" Then
2420               strInput = 99
2421           End If
2422           sBC = CSng(strInput)
2423       Else
2424           GoTo CancelPressed
2425       End If

2426       rsBalanceHoles!Number = sNumber
2427       rsBalanceHoles!diameter = sDia
2428       rsBalanceHoles!boltcircle = sBC

2429       rsBalanceHoles.Update
           'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

2430       GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '    rsBalanceHoles.Requery
2431       rsBalanceHoles.MoveLast
2432       dgBalanceHoles.Refresh
2433       chkBalanceHoles.value = 1
2434       rsBalanceHoles.MoveFirst

       '    Set dgBalanceHoles.DataSource = rsBalanceHoles
       '
       '    Dim c As Column
       '    For Each c In dgBalanceHoles.Columns
       '        Select Case c.DataField
       '        Case "BalanceHoleID"
       '            c.Visible = False
       '        Case "SerialNo"
       '            c.Visible = False
       '        Case "Date"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 2000
       '        Case "Number"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 700
       '        Case "Diameter"
       '            c.Visible = False
       '        Case "Diameter1"
       '            c.Caption = "Diameter"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 700
       '        Case "BoltCircle1"
       '            c.Caption = "Bolt Circle"
       '            c.Visible = True
       '            c.Alignment = dbgCenter
       '            c.Width = 800
       '        Case "BoltCircle"
       '            c.Visible = False
       '        Case Else
       '        End Select
       '    Next c
' <VB WATCH>
2435       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2436       Exit Sub

2437   CancelPressed:
2438       MsgBox "No New Balance Hole Data Entered", vbOKOnly

2439   DeleteIt:
2440       If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
2441           rsBalanceHoles.Delete
2442           rsBalanceHoles.Update
2443           GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '        rsBalanceHoles.Requery
2444           If rsBalanceHoles.RecordCount > 0 Then
2445               rsBalanceHoles.MoveLast
2446           End If
2447           dgBalanceHoles.Refresh
2448           chkBalanceHoles.value = 1
2449           rsBalanceHoles.MoveFirst
2450       End If


' <VB WATCH>
2451       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2452       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdModifyBalanceHoleData_Click"

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
            vbwReportVariable "strInput", strInput
            vbwReportVariable "I", I
            vbwReportVariable "sNumber", sNumber
            vbwReportVariable "sDia", sDia
            vbwReportVariable "sBC", sBC
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdReport_Click()
           'view/print a report
' <VB WATCH>
2453       On Error GoTo vbwErrHandler
2454       Const VBWPROCNAME = "frmPLCData.cmdReport_Click"
2455       If vbwProtector.vbwTraceProc Then
2456           Dim vbwProtectorParameterString As String
2457           If vbwProtector.vbwTraceParameters Then
2458               vbwProtectorParameterString = "()"
2459           End If
2460           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2461       End If
' </VB WATCH>
2462       Dim I As Integer

2463       frmReport.Visible = True
2464       For I = 0 To optReport.Count - 1
2465           optReport(I).value = False
2466       Next I

' <VB WATCH>
2467       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2468       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdReport_Click"

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
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdSearchForPump_Click()
' <VB WATCH>
2469       On Error GoTo vbwErrHandler
2470       Const VBWPROCNAME = "frmPLCData.cmdSearchForPump_Click"
2471       If vbwProtector.vbwTraceProc Then
2472           Dim vbwProtectorParameterString As String
2473           If vbwProtector.vbwTraceParameters Then
2474               vbwProtectorParameterString = "()"
2475           End If
2476           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2477       End If
' </VB WATCH>
2478       LoadCombo frmSearch.cmbSearchModel, "Model"
2479       frmSearch.Show
' <VB WATCH>
2480       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2481       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdSearchForPump_Click"

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

Private Sub cmdWriteSP_Click()
           'write the sp to the plc
' <VB WATCH>
2482       On Error GoTo vbwErrHandler
2483       Const VBWPROCNAME = "frmPLCData.cmdWriteSP_Click"
2484       If vbwProtector.vbwTraceProc Then
2485           Dim vbwProtectorParameterString As String
2486           If vbwProtector.vbwTraceParameters Then
2487               vbwProtectorParameterString = "()"
2488           End If
2489           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2490       End If
' </VB WATCH>
2491       Dim rc As String
2492       Dim S As String

           'write the set point data to the PLC
2493           bWrite = True
2494           S = Right$("0000" & txtWriteSPData, 4)
2495           S = Right$(S, 2) & Left$(S, 2)
2496           rc = StringToByteArray(S, ByteBuffer)

2497           DataLength = HexConvert(ByteBuffer, 2)
2498           DataAddress = StringToHexInt("2005")

2499           rc = GetData

2500           bWrite = False
' <VB WATCH>
2501       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2502       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdWriteSP_Click"

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
            vbwReportVariable "S", S
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Private Sub Command1_Click()
       '    Dim frmem As New InteropDBWithButtons.Form1
       '    frmem.ConString = cnPumpData.ConnectionString
       '    frmem.Caption = "Email Database Maintenance"
       '    frmem.Show 1
' <VB WATCH>
2503       On Error GoTo vbwErrHandler
2504       Const VBWPROCNAME = "frmPLCData.Command1_Click"
2505       If vbwProtector.vbwTraceProc Then
2506           Dim vbwProtectorParameterString As String
2507           If vbwProtector.vbwTraceParameters Then
2508               vbwProtectorParameterString = "()"
2509           End If
2510           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2511       End If
' </VB WATCH>
' <VB WATCH>
2512       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2513       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Command1_Click"

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

Private Sub Command2_Click()
' <VB WATCH>
2514       On Error GoTo vbwErrHandler
2515       Const VBWPROCNAME = "frmPLCData.Command2_Click"
2516       If vbwProtector.vbwTraceProc Then
2517           Dim vbwProtectorParameterString As String
2518           If vbwProtector.vbwTraceParameters Then
2519               vbwProtectorParameterString = "()"
2520           End If
2521           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2522       End If
' </VB WATCH>
2523       ReportToExcel
' <VB WATCH>
2524       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2525       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Command2_Click"

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







Private Sub updown1_change()
' <VB WATCH>
2526       On Error GoTo vbwErrHandler
2527       Const VBWPROCNAME = "frmPLCData.updown1_change"
2528       If vbwProtector.vbwTraceProc Then
2529           Dim vbwProtectorParameterString As String
2530           If vbwProtector.vbwTraceParameters Then
2531               vbwProtectorParameterString = "()"
2532           End If
2533           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2534       End If
' </VB WATCH>
2535       Dim sName As String

2536       If Not rsTestData.BOF Then
2537           rsTestData.MoveFirst
2538       End If

2539       If Not rsTestData.BOF Or Not rsTestData.EOF Then
2540           rsTestData.Move UpDown1.value - 1
2541       End If

2542       sName = "VibrationX"
2543       If rsTestData.Fields(sName).ActualSize <> 0 Then
2544           txtVibAx.Text = rsTestData.Fields(sName)
2545       Else
       '        txtVibAx.Text = vbNullString
2546       End If

2547       sName = "VibrationY"
2548       If rsTestData.Fields(sName).ActualSize <> 0 Then
2549           txtVibRad.Text = rsTestData.Fields(sName)
2550       Else
       '        txtVibRad.Text = vbNullString
2551       End If

2552       sName = "Remarks"
2553       If rsTestData.Fields(sName).ActualSize <> 0 Then
2554           txtTestRemarks.Text = rsTestData.Fields(sName)
2555       Else
       '        txtTestRemarks.Text = vbNullString
2556       End If

2557       sName = "ThrustBalance"
2558       If rsTestData.Fields(sName).ActualSize <> 0 Then
2559           txtThrustBal.Text = rsTestData.Fields(sName)
2560       Else
       '        txtThrustBal.Text = vbNullString
2561       End If

2562       sName = "TEMCTRG"
2563       If rsTestData.Fields(sName).ActualSize <> 0 Then
2564           txtTEMCTRGReading.Text = rsTestData.Fields(sName)
2565       Else
2566           txtTEMCTRGReading.Text = 0
       '        txtTEMCTRGReading.Text = vbNullString
2567       End If

2568       sName = "TEMCFrontThrust"
2569       If rsTestData.Fields(sName).ActualSize <> 0 Then
2570           txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
2571       Else
       '        txtTEMCFrontThrust.Text = vbNullString
2572       End If

2573       sName = "TEMCRearThrust"
2574       If rsTestData.Fields(sName).ActualSize <> 0 Then
2575           txtTEMCRearThrust.Text = rsTestData.Fields(sName)
2576       Else
       '        txtTEMCRearThrust.Text = vbNullString
2577       End If
2578       sName = "TEMCMomentArm"
2579       If rsTestData.Fields(sName).ActualSize <> 0 Then
2580           txtTEMCMomentArm.Text = rsTestData.Fields(sName)
2581       Else
       '        txtTEMCMomentArm.Text = vbNullString
2582       End If
2583       sName = "TEMCThrustRigPressure"
2584       If rsTestData.Fields(sName).ActualSize <> 0 Then
2585           txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
2586       Else
       '        txtTEMCThrustRigPressure.Text = vbNullString
2587       End If
2588       sName = "TEMCViscosity"
2589       If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
2590           txtTEMCViscosity.Text = rsTestData.Fields(sName)
2591       Else
       '        txtTEMCViscosity.Text = vbNullString
2592       End If

2593       CalculateTEMCForce

2594       rsEff.MoveFirst
2595       rsEff.Move UpDown1.value - 1
' <VB WATCH>
2596       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2597       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "updown1_change"

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
            vbwReportVariable "sName", sName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub CalculateTEMCForce()
' <VB WATCH>
2598       On Error GoTo vbwErrHandler
2599       Const VBWPROCNAME = "frmPLCData.CalculateTEMCForce"
2600       If vbwProtector.vbwTraceProc Then
2601           Dim vbwProtectorParameterString As String
2602           If vbwProtector.vbwTraceParameters Then
2603               vbwProtectorParameterString = "()"
2604           End If
2605           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2606       End If
' </VB WATCH>
2607       Dim NoOfPoles As Integer
2608       Dim Frequency As Integer
2609       Dim Additions As String
2610       Dim Frame As String
2611       Dim VOverA As Double
2612       Dim Force As Double

           'show calculated values
2613       If Val(txtTEMCFrontThrust.Text) = 0 Then
2614           If Val(txtTEMCRearThrust.Text) = 0 Then
               'no thrust entered
2615               lblTEMCFrontRear.Visible = False
2616               txtTEMCCalcForce.Text = " "
2617           Else
                   'rear thrust
2618               txtTEMCCalcForce.Text = Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5
2619               lblTEMCFrontRear.Caption = "REAR"
2620               lblTEMCFrontRear.Visible = True
2621           End If
2622       Else
               'front thrust
2623           txtTEMCCalcForce.Text = Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5
2624           lblTEMCFrontRear.Caption = "FRONT"
2625           lblTEMCFrontRear.Visible = True
2626       End If

2627       If Val(txtTEMCCalcForce.Text) < 0 Then
2628           txtTEMCCalcForce.Text = -txtTEMCCalcForce
2629           lblTEMCFrontRear.Caption = "FRONT"
2630       End If

           'see how many poles we have, it's the next to last number in the frame size
2631       If Len(txtTEMCFrameNumber) > 2 Then
2632           NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
2633       End If

2634       If cmbTEMCAdditions.ListIndex <> -1 Then
2635           Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
2636           If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
2637               Frequency = 60
2638           ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
2639               Frequency = 50
2640           Else
2641               Frequency = 0
2642           End If
2643       End If

2644       If Len(txtTEMCFrameNumber.Text) = 3 Then
2645           Frame = Left$(txtTEMCFrameNumber, 2) & "0"
2646       Else
2647           Frame = txtTEMCFrameNumber.Text
2648           If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
2649               Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
2650           Else
2651           End If
2652       End If
2653       Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
2654       If Frequency = 60 Then
2655           Force = Force / 1.2
2656       End If
2657       If Val(txtTEMCViscosity.Text) > 1# Then
2658           If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
2659               lblTEMCPassFail.Visible = True
2660               lblTEMCPassFail.ForeColor = vbRed
2661               lblTEMCPassFail.Caption = "FAIL"
2662           Else
2663               lblTEMCPassFail.Visible = True
2664               lblTEMCPassFail.ForeColor = vbGreen
2665               lblTEMCPassFail.Caption = "PASS"
2666           End If
2667       End If

2668       If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
2669           If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
2670               lblTEMCPassFail.Visible = True
2671               lblTEMCPassFail.ForeColor = vbRed
2672               lblTEMCPassFail.Caption = "FAIL"
2673           Else
2674               lblTEMCPassFail.Visible = True
2675               lblTEMCPassFail.ForeColor = vbGreen
2676               lblTEMCPassFail.Caption = "PASS"
2677           End If
2678       End If

2679       If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
2680           If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
2681               lblTEMCPassFail.Visible = True
2682               lblTEMCPassFail.ForeColor = vbRed
2683               lblTEMCPassFail.Caption = "FAIL"
2684           Else
2685               lblTEMCPassFail.Visible = True
2686               lblTEMCPassFail.ForeColor = vbGreen
2687               lblTEMCPassFail.Caption = "PASS"
2688           End If
2689       End If

2690       If (Val(txtTEMCViscosity.Text) <= 0.3) Then
2691           If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
2692               lblTEMCPassFail.Visible = True
2693               lblTEMCPassFail.ForeColor = vbRed
2694               lblTEMCPassFail.Caption = "FAIL"
2695           Else
2696               lblTEMCPassFail.Visible = True
2697               lblTEMCPassFail.ForeColor = vbGreen
2698               lblTEMCPassFail.Caption = "PASS"
2699           End If
2700       End If
2701       If NoOfPoles <> 0 Then
2702           VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles / 2)
2703       End If
2704       If Frequency = 60 Then
2705           VOverA = VOverA * 1.2
2706       End If

2707       txtTEMCPVValue.Text = Val(txtTEMCCalcForce.Text) * VOverA

2708       If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
2709           txtTEMCPVValue.Text = ""
2710           txtTEMCCalcForce.Text = ""
2711           lblTEMCPassFail.Visible = False
2712       End If

' <VB WATCH>
2713       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2714       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalculateTEMCForce"

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
            vbwReportVariable "NoOfPoles", NoOfPoles
            vbwReportVariable "Frequency", Frequency
            vbwReportVariable "Additions", Additions
            vbwReportVariable "Frame", Frame
            vbwReportVariable "VOverA", VOverA
            vbwReportVariable "Force", Force
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub UpDown2_change()
' <VB WATCH>
2715       On Error GoTo vbwErrHandler
2716       Const VBWPROCNAME = "frmPLCData.UpDown2_change"
2717       If vbwProtector.vbwTraceProc Then
2718           Dim vbwProtectorParameterString As String
2719           If vbwProtector.vbwTraceParameters Then
2720               vbwProtectorParameterString = "()"
2721           End If
2722           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2723       End If
' </VB WATCH>
2724       Dim Plothead(7, 1) As Single
2725       Dim PlotEff(7, 1) As Single
2726       Dim PlotKW(7, 1) As Single
2727       Dim PlotAmps(7, 1) As Single

2728       Dim j As Integer

2729       For j = 0 To UpDown2.value - 1
2730           Plothead(j, 0) = HeadFlow(0, j)
2731           Plothead(j, 1) = HeadFlow(1, j)

2732           PlotEff(j, 0) = EffFlow(0, j)
2733           PlotEff(j, 1) = EffFlow(1, j)
2734           PlotKW(j, 0) = KWFlow(0, j)
2735           PlotKW(j, 1) = KWFlow(1, j)
2736           PlotAmps(j, 0) = AmpsFlow(0, j)
2737           PlotAmps(j, 1) = AmpsFlow(1, j)
2738       Next j

2739       MSChart1 = Plothead
2740       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
2741       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(Plothead) / 10) + 0.5) + 1)
2742       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
2743       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

2744       MSChart3 = PlotAmps
2745       MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
2746       MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(PlotAmps) / 10) + 0.5) + 1)
2747       MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
2748       MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

2749       MSChart4 = PlotKW
2750       MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
2751       MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(PlotKW) / 10) + 0.5) + 1)
2752       MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
2753       MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

2754       MSChart5 = PlotEff
2755       MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
2756       MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(PlotEff) / 10) + 0.5) + 1)
2757       MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
2758       MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

2759       MSChart6 = Plothead
2760       MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
2761       MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(Plothead) / 10) + 0.5) + 1)
2762       MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
2763       MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

' <VB WATCH>
2764       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2765       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "UpDown2_change"

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
            vbwReportVariable "Plothead", Plothead
            vbwReportVariable "PlotEff", PlotEff
            vbwReportVariable "PlotKW", PlotKW
            vbwReportVariable "PlotAmps", PlotAmps
            vbwReportVariable "j", j
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
' <VB WATCH>
2766       On Error GoTo vbwErrHandler
2767       Const VBWPROCNAME = "frmPLCData.DataGrid1_AfterColUpdate"
2768       If vbwProtector.vbwTraceProc Then
2769           Dim vbwProtectorParameterString As String
2770           If vbwProtector.vbwTraceParameters Then
2771               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ColIndex", ColIndex) & ") "
2772           End If
2773           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2774       End If
' </VB WATCH>
2775       DoEfficiencyCalcs
' <VB WATCH>
2776       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2777       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DataGrid1_AfterColUpdate"

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
            vbwReportVariable "ColIndex", ColIndex
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub dgBalanceHoles_SelChange(Cancel As Integer)
' <VB WATCH>
2778       On Error GoTo vbwErrHandler
2779       Const VBWPROCNAME = "frmPLCData.dgBalanceHoles_SelChange"
2780       If vbwProtector.vbwTraceProc Then
2781           Dim vbwProtectorParameterString As String
2782           If vbwProtector.vbwTraceParameters Then
2783               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
2784           End If
2785           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2786       End If
' </VB WATCH>
2787       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2788           cmdModifyBalanceHoleData.Visible = False
2789       Else
2790           cmdModifyBalanceHoleData.Visible = True
2791       End If
' <VB WATCH>
2792       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2793       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "dgBalanceHoles_SelChange"

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
Private Sub Form_Load()
' <VB WATCH>
2794       On Error GoTo vbwErrHandler
2795       Const VBWPROCNAME = "frmPLCData.Form_Load"
2796       If vbwProtector.vbwTraceProc Then
2797           Dim vbwProtectorParameterString As String
2798           If vbwProtector.vbwTraceParameters Then
2799               vbwProtectorParameterString = "()"
2800           End If
2801           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2802       End If
' </VB WATCH>
2803       Dim RetVal As String
2804       Dim sSendStr As String
2805       Dim I As Integer
2806       Dim j As Integer
2807       Dim sTableName As String
2808       Dim WhichServer As String
2809       Dim WhichDatabase As String

2810       debugging = 0   'assume not debugging
2811       WhichServer = "Production"     'change to production server
2812       WhichDatabase = "Production"

2813       If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
2814           I = MsgBox("Debug?", vbYesNo)
2815           If I = vbYes Then
2816               debugging = 1
2817               WhichServer = "Production"
2818               WhichDatabase = "Production"
2819           Else
2820           End If
2821       End If

2822       If debugging Then
2823           GoTo temp
2824       End If
           'see if the mdb file is where it's supposed to be
2825       If Dir(sDevelopmentDatabase) = "" Then
2826           MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
2827           End
2828       End If

           'get the database info from the new mdb file
2829       Dim cnDevelopment As New ADODB.Connection
2830       Dim qyDevelopment As New ADODB.Command
2831       Dim rsDevelopment As New ADODB.Recordset

2832       On Error GoTo CannotConnect

2833       With cnDevelopment
2834           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDevelopmentDatabase & ";Persist Security Info=False;Jet OLEDB:Database Password=Access7277word;"
2835           .ConnectionTimeout = 10
2836           .Open
2837       End With

2838   On Error GoTo vbwErrHandler
2839       GoTo Connected

2840   CannotConnect:
2841       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2842       End

2843   Connected:

           'we're connected, get the data for the Epicor SQL server
2844       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
2845       qyDevelopment.ActiveConnection = cnDevelopment

2846       rsDevelopment.CursorLocation = adUseClient
2847       rsDevelopment.CursorType = adOpenStatic
2848       rsDevelopment.LockType = adLockOptimistic

2849       On Error GoTo NoServerData

2850       rsDevelopment.Open qyDevelopment

2851   On Error GoTo vbwErrHandler
2852       GoTo GotServerData

2853   NoServerData:

2854       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2855       End

2856   GotServerData:

2857       If rsDevelopment.RecordCount <> 1 Then
2858           GoTo NoServerData
2859       End If

           'construct Epicor connection string
2860       EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                           "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                           "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                           "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                           "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


           'make sure we can open the SQL database

2861       On Error GoTo CannotOpenEpicorSQLServer

2862       Dim cnTestEpicor As New ADODB.Connection
2863       cnTestEpicor.ConnectionString = EpicorConnectionString
2864       cnTestEpicor.Open
2865       cnTestEpicor.Close
2866       Set cnTestEpicor = Nothing
2867   On Error GoTo vbwErrHandler

2868       GoTo FoundEpicorSQLServer

2869   CannotOpenEpicorSQLServer:
2870       MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2871       End

2872   FoundEpicorSQLServer:
           'get data on rundown database
2873       rsDevelopment.Close
2874       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PumpRundown'"

2875       On Error GoTo NoRundownDatabase

2876       rsDevelopment.Open qyDevelopment

2877       GoTo FoundRundownDatabase

2878   NoRundownDatabase:
2879       MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2880       End

2881   FoundRundownDatabase:
2882       If rsDevelopment.RecordCount <> 1 Then
2883           GoTo NoRundownDatabase
2884           End
2885       End If

2886   temp:

2887       If debugging Then
2888           ParentDirectoryName = "C:\databases"
2889           sDataBaseName = "c:\databases\PumpData 2k.mdb"
2890       Else
2891           ParentDirectoryName = "\\TEI-MAIN-01\f\groups\shared\databases"
2892           sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

       '        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
2893       End If

           'see if we can open the pump rundown database
2894       On Error GoTo NoRundownDatabase
2895       With cnPumpData
2896           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False"
2897           .ConnectionTimeout = 10
2898           .Open
2899       End With
2900   On Error GoTo vbwErrHandler


2901       If debugging = 0 Then
2902           Printer.Orientation = vbPRORLandscape
2903       End If

2904       lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision

2905       boFoundPump = False

2906       Me.Show

2907       Dim k As Integer
2908       For k = 0 To 20
2909           vPlot(k, 0) = 0
2910           vPlot(k, 1) = 0
2911       Next k

2912       With MSChart1
2913           .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
2914           .Plot.Axis(VtChAxisIdY).AxisTitle = "TDH (Ft)"
2915           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
2916           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
2917           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
2918           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
2919           .Plot.UniformAxis = False
2920           .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2921           .Plot.SeriesCollection.Item(1).Pen.Width = 15
2922           .Plot.AutoLayout = False
2923           .Plot.LocationRect.Max.X = 5700
2924           .Plot.LocationRect.Max.Y = 3000
2925           .Plot.LocationRect.Min.X = -100
2926           .Plot.LocationRect.Min.Y = -100
2927       End With
2928       With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2929           .Visible = True
2930           .Size = 60
2931           .Style = VtMarkerStyleCircle
2932           .FillColor.Automatic = False
2933           .FillColor.Set 0, 0, 255
2934       End With

2935       With MSChart2
       '
2936           .Plot.AutoLayout = False
2937           .Plot.LocationRect.Max.X = 2000
2938           .Plot.LocationRect.Max.Y = 1500
2939           .Plot.LocationRect.Min.X = 0
2940           .Plot.LocationRect.Min.Y = 0

2941       End With

2942       With MSChart3
2943           .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
2944           .Plot.Axis(VtChAxisIdY).AxisTitle = "Current (Amps)"
2945           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
2946           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
2947           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
2948           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
2949           .Plot.UniformAxis = False
2950           .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2951           .Plot.SeriesCollection.Item(1).Pen.Width = 15
2952           .Plot.AutoLayout = False
2953           .Plot.LocationRect.Max.X = 5700
2954           .Plot.LocationRect.Max.Y = 3000
2955           .Plot.LocationRect.Min.X = -100
2956           .Plot.LocationRect.Min.Y = -100
2957       End With
2958       With MSChart3.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2959           .Visible = True
2960           .Size = 60
2961           .Style = VtMarkerStyleCircle
2962           .FillColor.Automatic = False
2963           .FillColor.Set 0, 0, 255
2964       End With
2965       With MSChart4
2966           .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
2967           .Plot.Axis(VtChAxisIdY).AxisTitle = "kW"
2968           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
2969           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
2970           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
2971           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
2972           .Plot.UniformAxis = False
2973           .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2974           .Plot.SeriesCollection.Item(1).Pen.Width = 15
2975           .Plot.AutoLayout = False
2976           .Plot.LocationRect.Max.X = 5700
2977           .Plot.LocationRect.Max.Y = 3000
2978           .Plot.LocationRect.Min.X = -100
2979           .Plot.LocationRect.Min.Y = -100
2980       End With
2981       With MSChart4.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2982           .Visible = True
2983           .Size = 60
2984           .Style = VtMarkerStyleCircle
2985           .FillColor.Automatic = False
2986           .FillColor.Set 0, 0, 255
2987       End With
2988       With MSChart5
2989           .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
2990           .Plot.Axis(VtChAxisIdY).AxisTitle = "Efficiency (%)"
2991           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
2992           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
2993           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
2994           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
2995           .Plot.UniformAxis = False
2996           .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2997           .Plot.SeriesCollection.Item(1).Pen.Width = 15
2998           .Plot.AutoLayout = False
2999           .Plot.LocationRect.Max.X = 5700
3000           .Plot.LocationRect.Max.Y = 3000
3001           .Plot.LocationRect.Min.X = -100
3002           .Plot.LocationRect.Min.Y = -100
3003       End With
3004       With MSChart5.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
3005           .Visible = True
3006           .Size = 60
3007           .Style = VtMarkerStyleCircle
3008           .FillColor.Automatic = False
3009           .FillColor.Set 0, 0, 255
3010       End With

3011       With MSChart6
3012           .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
3013           .Plot.Axis(VtChAxisIdY).AxisTitle = "TDH (Ft)"
3014           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
3015           .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
3016           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
3017           .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
3018           .Plot.UniformAxis = False
3019           .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
3020           .Plot.SeriesCollection.Item(1).Pen.Width = 15
3021           .Plot.AutoLayout = False
3022           .Plot.LocationRect.Max.X = 5700
3023           .Plot.LocationRect.Max.Y = 3000
3024           .Plot.LocationRect.Min.X = -100
3025           .Plot.LocationRect.Min.Y = -100
3026       End With
3027       With MSChart6.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
3028           .Visible = True
3029           .Size = 60
3030           .Style = VtMarkerStyleCircle
3031           .FillColor.Automatic = False
3032           .FillColor.Set 0, 0, 255
3033       End With

           'assure that the timers are off
3034       frmPLCData.tmrGetDDE.Enabled = False

3035       frmPLCData.tmrStartUp.Enabled = False

           'initialize the PLC network
3036       RetVal = NetWorkInitialize()
3037       If RetVal <> 0 Then
3038           MsgBox ("Can't Initialize Network. Exiting...")
3039           End
3040       End If

3041       If debugging = 0 Then
               'load array of plcs
3042           I = 0
3043           Open rsDevelopment.Fields("ServerName") & "plcaddresses.txt" For Input As 1
3044           While Not EOF(1)
3045               Input #1, Description(I)
3046               For j = 0 To 125
3047                   Input #1, aDevices(I).Address(j)
3048               Next j
3049               Input #1, j
3050               I = I + 1
3051           Wend
3052           Close #1

3053           DeviceCount = I

3054           If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
       '            Dim k As Integer
3055               For k = 0 To DeviceCount - 1
3056                   If InStr(Description(k), "MWSC") <> 0 Then
3057                       Exit For
3058                   End If
3059               Next k
3060               Description(DeviceCount) = Description(0)
3061               Description(0) = Description(k)
3062               Description(k) = Description(DeviceCount)

3063               aDevices(DeviceCount) = aDevices(0)
3064               aDevices(0) = aDevices(k)
3065               aDevices(k) = aDevices(DeviceCount)

3066           End If

3067           Dim PLCAddress As String
3068           For I = 0 To DeviceCount - 1
3069               PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
3070               RetVal = PingSilent(PLCAddress)
3071               If RetVal <> 0 Then
3072                   frmPLCData.cmbPLCLoop.AddItem Description(I)
3073                   frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
3074               End If
3075           Next I
3076       End If

3077       frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

           'turn on the PLC led

3078       frmPLCData.cmbPLCLoop.ListIndex = 0
3079       frmPLCData.tmrGetDDE.Enabled = True

           'hook up to the various databases

3080       DataEnvironment2.Connection1.ConnectionString = cnPumpData
3081       DataEnvironment3.Connection1.ConnectionString = cnPumpData

3082       With cnEffData
3083           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
3084           .Open
3085       End With

           'open some recordsets
3086       rsPumpData.Index = "SerialNumber"
3087       rsTestSetup.Index = "FindData"
3088       rsTestData.Index = "PrimaryKey"
3089       rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
3090       rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
3091       rsTestData.Filter = "SerialNumber = ''"
3092       rsTestData.CursorLocation = adUseClient
3093       rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
3094       rsEff.CursorLocation = adUseClient
3095       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

3096       qyBalanceHoles.ActiveConnection = cnPumpData
3097       rsBalanceHoles.CursorLocation = adUseClient
3098       rsBalanceHoles.CursorType = adOpenStatic
3099       rsBalanceHoles.LockType = adLockOptimistic
       '    qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 FROM BalanceHoles;"
       '    rsBalanceHoles.Open qyBalanceHoles
       '    rsBalanceHoles.Filter = "SerialNo = ''"



3100       If debugging <> 1 Then
3101           FindMagtrols
3102       Else
3103           cmbMagtrol.AddItem "Add Manually"
3104           cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
3105           cmbMagtrol.ListIndex = 0
3106       End If
3107       optKW(1).value = True
3108       optKW_Click (1)


           'blank out data grid
3109       Set DataGrid1.DataSource = rsTestData

           'load the combo boxes
3110       LoadCombo cmbStatorFill, "StatorFill"
3111       LoadCombo cmbCirculationPath, "CirculationPath"
3112       LoadCombo cmbVoltage, "Voltage"
3113       LoadCombo cmbFrequency, "Frequency"
3114       LoadCombo cmbMotor, "Motor"
3115       LoadCombo cmbDesignPressure, "DesignPressure"
3116       LoadCombo cmbRPM, "RPM"
3117       LoadCombo cmbOrificeNumber, "OrificeNumber"
3118       LoadCombo cmbTestSpec, "TestSpecification"
3119       LoadCombo cmbLoopNumber, "LoopNumber"
3120       LoadCombo cmbSuctDia, "SuctionDiameter"
3121       LoadCombo cmbDischDia, "DischargeDiameter"
3122       LoadCombo cmbModel, "Model"
3123       LoadCombo cmbModelGroup, "ModelGroup"
3124       LoadCombo cmbMounting, "Mounting"
3125       LoadCombo cmbTachID, "TachID"
3126       LoadCombo cmbAnalyzerNo, "AnalyzerNo"
3127       LoadCombo cmbPLCNo, "PLCNo"
3128       LoadCombo cmbFlowMeter, "Flowmeter"
       '    LoadInstrumentationCombo cmbTachID, "TachID"
       '    LoadInstrumentationCombo cmbAnalyzerNo, "AnalyzerNo"
       '    LoadInstrumentationCombo cmbPLCNo, "PLCNo"
       '    LoadInstrumentationCombo cmbFlowMeter, "Flowmeter"

           'load the TEMC combo boxes, too
3129       LoadCombo cmbTEMCAdapter, "TEMCAdapter"
3130       LoadCombo cmbTEMCAdditions, "TEMCAdditions"
3131       LoadCombo cmbTEMCCirculation, "TEMCCirculation"
3132       LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
3133       LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
3134       LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
3135       LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
3136       LoadCombo cmbTEMCInsulation, "TEMCInsulation"
3137       LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
3138       LoadCombo cmbTEMCMaterials, "TEMCMaterials"
3139       LoadCombo cmbTEMCModel, "TEMCModel"
3140       LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
3141       LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
3142       LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
3143       LoadCombo cmbTEMCVoltage, "TEMCVoltage"
3144       LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
3145       LoadCombo cmbTEMCTRG, "TEMCTRG"

3146       LoadCombo frmSearch.cmbSearchModel, "Model"

           'fill memory arrays for dlookups
3147       FillArrays

           'choose the first tab
3148       frmPLCData.SSTab1.Tab = 0

           'set the grid column names
3149       Dim c As Column
3150       For Each c In DataGrid1.Columns
3151           Select Case c.DataField
               Case "TestDataID"
3152               c.Visible = False
3153           Case "SerialNumber"
3154               c.Visible = False
3155           Case "Date"
3156               c.Visible = False
3157           Case Else ' Show all other columns.
3158               c.Visible = True
3159               c.Alignment = dbgRight
3160           End Select
3161       Next c

3162       Set dgBalanceHoles.DataSource = rsBalanceHoles

3163       For Each c In dgBalanceHoles.Columns
3164           Select Case c.DataField
               Case "BalanceHoleID"
3165               c.Visible = False
3166           Case "SerialNo"
3167               c.Visible = False
3168           Case "Date"
3169               c.Visible = True
3170               c.Alignment = dbgCenter
3171               c.Width = 2000
3172           Case "Number"
3173               c.Visible = True
3174               c.Alignment = dbgCenter
3175               c.Width = 700
3176           Case "Diameter"
3177               c.Visible = False
3178           Case "Diameter1"
3179               c.Caption = "Diameter"
3180               c.Visible = True
3181               c.Alignment = dbgCenter
3182               c.Width = 700
3183           Case "BoltCircle1"
3184               c.Caption = "Bolt Circle"
3185               c.Visible = True
3186               c.Alignment = dbgCenter
3187               c.Width = 800
3188           Case "BoltCircle"
3189               c.Visible = False
3190           Case "SetNo"
3191               c.Visible = False
3192           Case Else ' Show all other columns.
3193               c.Visible = False
3194           End Select
3195       Next c

3196       BlankData

       '    If debugging <> 1 Then
               'get user initials
3197           frmLogin.Show
       '    End If

           'setup eff.mdb file
3198       DataEnvironment1.Connection2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
       '    DataEnvironment3.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"

           'set dropdown width
3199       SendMessage cmbAnalyzerNo.hwnd, &H160, SendMessage(cmbAnalyzerNo.hwnd, &H15F, 0, 0) + 100, 0
3200       SendMessage cmbPLCNo.hwnd, &H160, SendMessage(cmbPLCNo.hwnd, &H15F, 0, 0) + 100, 0
3201       SendMessage cmbTachID.hwnd, &H160, SendMessage(cmbTachID.hwnd, &H15F, 0, 0) + 100, 0
3202       SendMessage cmbFlowMeter.hwnd, &H160, SendMessage(cmbFlowMeter.hwnd, &H15F, 0, 0) + 100, 0

3203       FromStoredData = False

' <VB WATCH>
3204       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3205       Exit Sub
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
            vbwReportVariable "RetVal", RetVal
            vbwReportVariable "sSendStr", sSendStr
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "sTableName", sTableName
            vbwReportVariable "WhichServer", WhichServer
            vbwReportVariable "WhichDatabase", WhichDatabase
            vbwReportVariable "k", k
            vbwReportVariable "PLCAddress", PLCAddress
            vbwReportVariable "cnDevelopment", cnDevelopment
            vbwReportVariable "qyDevelopment", qyDevelopment
            vbwReportVariable "rsDevelopment", rsDevelopment
            vbwReportVariable "cnTestEpicor", cnTestEpicor
            vbwReportVariable "c", c
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
3206       On Error GoTo vbwErrHandler
3207       Const VBWPROCNAME = "frmPLCData.Form_Unload"
3208       If vbwProtector.vbwTraceProc Then
3209           Dim vbwProtectorParameterString As String
3210           If vbwProtector.vbwTraceParameters Then
3211               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
3212           End If
3213           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3214       End If
' </VB WATCH>
3215       End
' <VB WATCH>
3216       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3217       Exit Sub
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

Private Sub Label15_Click()
' <VB WATCH>
3218       On Error GoTo vbwErrHandler
3219       Const VBWPROCNAME = "frmPLCData.Label15_Click"
3220       If vbwProtector.vbwTraceProc Then
3221           Dim vbwProtectorParameterString As String
3222           If vbwProtector.vbwTraceParameters Then
3223               vbwProtectorParameterString = "()"
3224           End If
3225           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3226       End If
' </VB WATCH>
3227       frmDiagram.Show
' <VB WATCH>
3228       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3229       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Label15_Click"

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

Private Sub lblAutoMan_Click(Index As Integer)
           '0 - Flow
           '1 - Suction
           '2 - Discharge
           '3 - Temperature
           '4 - A1 - Circ Flow
           '5 - A2 - RBH Temp
           '6 - A3 - RBH Press
           '7 - A4
' <VB WATCH>
3230       On Error GoTo vbwErrHandler
3231       Const VBWPROCNAME = "frmPLCData.lblAutoMan_Click"
3232       If vbwProtector.vbwTraceProc Then
3233           Dim vbwProtectorParameterString As String
3234           If vbwProtector.vbwTraceParameters Then
3235               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3236           End If
3237           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3238       End If
' </VB WATCH>

3239       Dim blnEnabled As Boolean

3240       If lblAutoMan(Index).Caption = "Auto" Then
3241           lblAutoMan(Index).Caption = "Man"
3242           blnEnabled = True
3243       Else
3244           lblAutoMan(Index).Caption = "Auto"
3245           blnEnabled = False
3246       End If

3247       Select Case Index
               Case 0
3248               txtFlowDisplay.Enabled = blnEnabled
3249           Case 1
3250               txtSuctionDisplay.Enabled = blnEnabled
3251           Case 2
3252               txtDischargeDisplay.Enabled = blnEnabled
3253           Case 3
3254               txtTemperatureDisplay.Enabled = blnEnabled
3255           Case 4
3256               txtAI1Display.Enabled = blnEnabled
3257           Case 5
3258               txtAI2Display.Enabled = blnEnabled
3259           Case 6
3260               txtAI3Display.Enabled = blnEnabled
3261           Case 7
3262               txtAI4Display.Enabled = blnEnabled
3263       End Select

' <VB WATCH>
3264       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3265       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "lblAutoMan_Click"

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
            vbwReportVariable "Index", Index
            vbwReportVariable "blnEnabled", blnEnabled
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Private Sub txtNPSHFileLocation_Click()
' <VB WATCH>
3266       On Error GoTo vbwErrHandler
3267       Const VBWPROCNAME = "frmPLCData.txtNPSHFileLocation_Click"
3268       If vbwProtector.vbwTraceProc Then
3269           Dim vbwProtectorParameterString As String
3270           If vbwProtector.vbwTraceParameters Then
3271               vbwProtectorParameterString = "()"
3272           End If
3273           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3274       End If
' </VB WATCH>
3275       Dim sTempDir As String
3276       On Error Resume Next
3277       sTempDir = CurDir    'Remember the current active directory
3278       CommonDialog2.DialogTitle = "Select a directory" 'titlebar
3279       CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
3280       CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
3281       CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
3282       CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
3283       CommonDialog2.CancelError = True 'allow escape key/cancel
3284       CommonDialog2.ShowSave   'show the dialog screen

3285       If Err <> 32755 Then    ' User didn't chose Cancel.
               'Me.SDir.Text = CurDir
3286       End If

       '    ChDir sTempDir  'restore path to what it was at entering

3287   Me.txtNPSHFileLocation.Text = CommonDialog2.filename

' <VB WATCH>
3288       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3289       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtNPSHFileLocation_Click"

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
            vbwReportVariable "sTempDir", sTempDir
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Private Sub txtTitle_LostFocus(Index As Integer)
' <VB WATCH>
3290       On Error GoTo vbwErrHandler
3291       Const VBWPROCNAME = "frmPLCData.txtTitle_LostFocus"
3292       If vbwProtector.vbwTraceProc Then
3293           Dim vbwProtectorParameterString As String
3294           If vbwProtector.vbwTraceParameters Then
3295               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3296           End If
3297           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3298       End If
' </VB WATCH>

3299       ChangeTitles Index

' <VB WATCH>
3300       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3301       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTitle_LostFocus"

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
            vbwReportVariable "Index", Index
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub ChangeTitles(ChannelNo As Integer)
' <VB WATCH>
3302       On Error GoTo vbwErrHandler
3303       Const VBWPROCNAME = "frmPLCData.ChangeTitles"
3304       If vbwProtector.vbwTraceProc Then
3305           Dim vbwProtectorParameterString As String
3306           If vbwProtector.vbwTraceParameters Then
3307               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ChannelNo", ChannelNo) & ") "
3308           End If
3309           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3310       End If
' </VB WATCH>
3311       Dim I As Integer
3312       Dim S As String

3313       If txtTitle(ChannelNo).Locked = True Then
' <VB WATCH>
3314       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3315           Exit Sub
3316       End If

3317       Dim qy As New ADODB.Command
3318       Dim rs As New ADODB.Recordset

3319       qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
3320       qy.CommandText = "SELECT * FROM AITitles " & _
               "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
               "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
               "AND ((AITitles.Channel)=" & ChannelNo & "));"

3321       With rs     'open the recordset for the query
3322           .CursorLocation = adUseClient
3323           .CursorType = adOpenStatic
3324           .LockType = adLockOptimistic
3325           .Open qy
3326       End With

3327       If (rs.BOF = True And rs.EOF = True) Then  'new record
3328           rs.AddNew
3329           rs.Fields("SerialNo") = txtSN.Text
3330           rs.Fields("Date") = cmbTestDate.Text
3331           rs.Fields("Channel") = CByte(ChannelNo)
3332           rs.Fields("Title") = txtTitle(ChannelNo).Text
3333           rs.Update
3334       Else    'we have an entry, modify it
3335           rs.Fields("SerialNo") = txtSN.Text
3336           rs.Fields("Date") = cmbTestDate.Text
3337           rs.Fields("Channel") = CByte(ChannelNo)
3338           rs.Fields("Title") = txtTitle(ChannelNo).Text
3339           rs.Update
3340       End If

3341       rs.Close
3342       Set rs = Nothing
3343       Set qy = Nothing

' <VB WATCH>
3344       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3345       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChangeTitles"

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
            vbwReportVariable "ChannelNo", ChannelNo
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub optKW_Click(Index As Integer)
' <VB WATCH>
3346       On Error GoTo vbwErrHandler
3347       Const VBWPROCNAME = "frmPLCData.optKW_Click"
3348       If vbwProtector.vbwTraceProc Then
3349           Dim vbwProtectorParameterString As String
3350           If vbwProtector.vbwTraceParameters Then
3351               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3352           End If
3353           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3354       End If
' </VB WATCH>
3355       Select Case Index
               Case 0  'add 3 powers
3356               txtKW.Enabled = False
3357           Case 1  'enter kw
3358               txtKW.Enabled = True
3359           Case 2  'use analog in 4
3360               txtKW.Enabled = False
3361       End Select
' <VB WATCH>
3362       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3363       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "optKW_Click"

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
            vbwReportVariable "Index", Index
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub optMfr_Click(Index As Integer)
' <VB WATCH>
3364       On Error GoTo vbwErrHandler
3365       Const VBWPROCNAME = "frmPLCData.optMfr_Click"
3366       If vbwProtector.vbwTraceProc Then
3367           Dim vbwProtectorParameterString As String
3368           If vbwProtector.vbwTraceParameters Then
3369               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3370           End If
3371           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3372       End If
' </VB WATCH>
3373       frmTEMC.Visible = optMfr(1).value
3374       frmChempump.Visible = optMfr(0).value
3375       frmTEMCData.Visible = optMfr(1).value
3376       txtModelNo_Change
' <VB WATCH>
3377       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3378       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "optMfr_Click"

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
            vbwReportVariable "Index", Index
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub optReport_Click(Index As Integer)
           'choose a report to view/print
' <VB WATCH>
3379       On Error GoTo vbwErrHandler
3380       Const VBWPROCNAME = "frmPLCData.optReport_Click"
3381       If vbwProtector.vbwTraceProc Then
3382           Dim vbwProtectorParameterString As String
3383           If vbwProtector.vbwTraceParameters Then
3384               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3385           End If
3386           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3387       End If
' </VB WATCH>

           'see if we have balance hole data for this pump
3388       Dim strBH As String
3389       Dim I As Integer

3390       If Index = 6 Then       'cancel pressed
3391           frmReport.Visible = False
' <VB WATCH>
3392       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3393           Exit Sub
3394       End If

3395       If Index <> 7 Then
3396           If boGotBalanceHoles Then
3397               If rsBalanceHoles.State = adStateClosed Then
3398                   rsBalanceHoles.ActiveConnection = cnPumpData
3399                   rsBalanceHoles.Open
3400               End If

3401               If rsBalanceHoles.RecordCount <> 0 Then
3402                   rsBalanceHoles.MoveFirst
3403                   strBH = Space(20) & "Balance Hole Data" & vbCrLf & "Date" & Space(24) & "Number" & Space(6) & "Dia" & Space(7) & "BC"
3404                   For I = 1 To rsBalanceHoles.RecordCount
3405                       strBH = strBH & vbCrLf & _
                           Left$(rsBalanceHoles.Fields("Date") & Space(30), 30) & _
                           Left$(rsBalanceHoles.Fields("Number") & Space(20), 10) & _
                           Left$(rsBalanceHoles.Fields("Diameter1") & Space(20), 10) & _
                           rsBalanceHoles.Fields("BoltCircle1")
3406                       rsBalanceHoles.MoveNext
3407                   Next I
3408               Else
3409               End If
3410           Else
       '            MsgBox "There are no balance holes for this test date", vbOKOnly, "No Balance Holes"
       '            frmReport.Visible = False
       '            Exit Sub
3411           End If
3412       End If


3413       Dim strVersion As String
3414       Dim isChart As Boolean

3415       isChart = False

3416       strVersion = "V" & App.Major & "." & App.Minor & "." & App.Revision

3417       Dim dr As DataReport
3418       frmReport.Visible = False
3419       optReport(Index).value = False

3420       Select Case Index
               Case 0  'Customer no circ flow
3421               Set dr = drCustNoCirc
3422           Case 1  'customer circ flow
3423               Set dr = drCustCirc
3424               frmReportOptions.Show 1

3425               Dim PosRPM As Integer
3426               Dim PosAxPos As Integer
3427               Dim PosCircFlow As Integer
3428               Dim PosVib As Integer
3429               Dim PosRem As Integer
3430               Dim PosTRG As Integer

3431                       PosTRG = frmReportOptions.chkTRG.value * 7920
3432                       PosRPM = frmReportOptions.chkSelectRPM.value * (7920 + frmReportOptions.chkTRG.value * 720)
3433                       PosAxPos = frmReportOptions.chkSelectAxPos.value * (7920 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)
3434                       PosCircFlow = frmReportOptions.chkSelectCircFlow.value * (7920 + frmReportOptions.chkSelectAxPos.value * 720 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)
3435                       PosVib = frmReportOptions.chkVibration.value * (7920 + frmReportOptions.chkSelectCircFlow.value * 720 + frmReportOptions.chkSelectAxPos.value * 720 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)
3436                       PosRem = 7920 + (frmReportOptions.chkVibration.value * 2 * 720 + frmReportOptions.chkSelectCircFlow.value * 720 + frmReportOptions.chkSelectAxPos.value * 720 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)

3437                       drCustCirc.Sections(2).Controls("labelTRG").Left = PosTRG
3438                       drCustCirc.Sections(3).Controls("textTRG").Left = PosTRG
3439                       drCustCirc.Sections(2).Controls("labelRemarks").Left = PosRem
3440                       drCustCirc.Sections(3).Controls("textRemarks").Left = PosRem
3441                       drCustCirc.Sections(2).Controls("labelRPM").Left = PosRPM
3442                       drCustCirc.Sections(3).Controls("textRPM").Left = PosRPM
3443                       drCustCirc.Sections(2).Controls("labelAxPos").Left = PosAxPos
3444                       drCustCirc.Sections(3).Controls("textAxPos").Left = PosAxPos
3445                       drCustCirc.Sections(2).Controls("labelCircFlow").Left = PosCircFlow
3446                       drCustCirc.Sections(3).Controls("textCircflow").Left = PosCircFlow
3447                       drCustCirc.Sections(2).Controls("labelSelectVibX").Left = PosVib
3448                       drCustCirc.Sections(3).Controls("textVibX").Left = PosVib
3449                       drCustCirc.Sections(2).Controls("labelSelectVibY").Left = PosVib + 720
3450                       drCustCirc.Sections(3).Controls("textVibY").Left = PosVib + 720

3451                       drCustCirc.Sections(2).Controls("labelTRG").Visible = PosTRG
3452                       drCustCirc.Sections(3).Controls("textTRG").Visible = PosTRG
3453                       drCustCirc.Sections(2).Controls("labelRPM").Visible = PosRPM
3454                       drCustCirc.Sections(3).Controls("textRPM").Visible = PosRPM
3455                       drCustCirc.Sections(2).Controls("labelAxPos").Visible = PosAxPos
3456                       drCustCirc.Sections(3).Controls("textAxPos").Visible = PosAxPos
3457                       drCustCirc.Sections(2).Controls("labelCircFlow").Visible = PosCircFlow
3458                       drCustCirc.Sections(3).Controls("textCircflow").Visible = PosCircFlow
3459                       drCustCirc.Sections(2).Controls("labelSelectVibX").Visible = PosVib
3460                       drCustCirc.Sections(3).Controls("textVibX").Visible = PosVib
3461                       drCustCirc.Sections(2).Controls("labelSelectVibY").Visible = PosVib
3462                       drCustCirc.Sections(3).Controls("textVibY").Visible = PosVib


3463           Case 2  'customer vibration
3464               Set dr = drCustVib
3465           Case 3  'internal
3466               Set dr = drInternal
3467               drInternal2.Sections(1).Controls("lblCustomer").Caption = txtShpNo
3468               drInternal2.Sections(1).Controls("lblBillTo").Caption = txtBilNo
3469               drInternal.Sections(1).Controls("lblBillTo").Caption = txtBilNo
3470               drInternal2.Sections(1).Controls("lblmodel").Caption = txtModelNo
3471               drInternal2.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
3472               drInternal2.Sections(1).Controls("lblSN").Caption = txtSN
3473               drInternal2.Sections(1).Controls("lblToday").Caption = Now
3474               drInternal2.Sections(1).Controls("lblVersion").Caption = strVersion
3475               drInternal2.Orientation = rptOrientLandscape
3476               drInternal2.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
3477               drInternal2.Sections(2).Controls("lblTC1").Caption = txtTitle(0).Text
3478               drInternal2.Sections(2).Controls("lblTC1A").Caption = txtTitle(1).Text
3479               drInternal2.Sections(2).Controls("lblTC2").Caption = txtTitle(2).Text
3480               drInternal2.Sections(2).Controls("lblTC2A").Caption = txtTitle(3).Text
3481               drInternal2.Sections(2).Controls("lblTC3").Caption = txtTitle(4).Text
3482               drInternal2.Sections(2).Controls("lblTC3A").Caption = txtTitle(5).Text
3483               drInternal2.Sections(2).Controls("lblTC4").Caption = txtTitle(6).Text
3484               drInternal2.Sections(2).Controls("lblTC4A").Caption = txtTitle(7).Text
3485               drInternal2.Sections(2).Controls("lblAI1").Caption = txtTitle(20).Text
3486               drInternal2.Sections(2).Controls("lblAI1A").Caption = txtTitle(21).Text
3487               drInternal2.Sections(2).Controls("lblAI2").Caption = txtTitle(22).Text
3488               drInternal2.Sections(2).Controls("lblAI2A").Caption = txtTitle(23).Text
3489               drInternal2.Sections(2).Controls("lblAI3").Caption = txtTitle(24).Text
3490               drInternal2.Sections(2).Controls("lblAI3A").Caption = txtTitle(25).Text
3491               drInternal2.Sections(2).Controls("lblAI4").Caption = txtTitle(26).Text
3492               drInternal2.Sections(2).Controls("lblAI4A").Caption = txtTitle(27).Text

3493           Case 5  'charts
3494               Set dr = drChart
3495               isChart = True
3496               drChart.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
3497               drChart.Sections(1).Controls("lblCustomer").Caption = txtShpNo
3498               drChart.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
3499               drChart.Sections(1).Controls("lblmodel").Caption = txtModelNo
3500               drChart.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
3501               drChart.Sections(1).Controls("lblSN").Caption = txtSN
3502               drChart.Sections(1).Controls("lblToday").Caption = Now
3503               drChart.Sections(1).Controls("lblVersion").Caption = strVersion

       '            Set drChart.Sections(1).Controls("Image1").Picture = CWGraphKw.ControlImage
       '            Set drChart.Sections(1).Controls("Image2").Picture = CWGraphEff.ControlImage
       '            Set drChart.Sections(1).Controls("Image3").Picture = CWGraphTDH.ControlImage
       '            Set drChart.Sections(1).Controls("Image4").Picture = CWGraphAmps.ControlImage

3504           Case 6  'escape out
' <VB WATCH>
3505       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3506               Exit Sub
3507           Case 7  'Unapproved pumps
3508               drApproved.Sections(1).Controls("lblNow").Caption = Now
3509               drApproved.Sections(1).Controls("lblVersion").Caption = strVersion
3510               drApproved.Show
' <VB WATCH>
3511       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3512               Exit Sub
3513           Case 4  'balance holes
3514               Set dr = drBalanceHoles
3515               If DataEnvironment3.Recordsets.Item(1).State = adStateOpen Then
3516                   DataEnvironment3.Recordsets.Item(1).Close
3517               End If
3518               If (LenB(frmPLCData.txtSN.Text) = 0) Or (LenB(cmbTestDate.List(cmbTestDate.ListIndex)) = 0) Then
' <VB WATCH>
3519       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3520                   Exit Sub
3521               End If
3522               DataEnvironment3.GetBalanceHoles frmPLCData.txtSN.Text, cmbTestDate.List(cmbTestDate.ListCount - 1)
       '            DataEnvironment3.GetBalanceHoles frmPLCData.txtSN.Text, cmbTestDate.List(cmbTestDate.ListIndex)
3523               drBalanceHoles.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
3524               drBalanceHoles.Sections(1).Controls("lblCustomer").Caption = txtShpNo
3525               drBalanceHoles.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
3526               drBalanceHoles.Sections(1).Controls("lblmodel").Caption = txtModelNo
3527               drBalanceHoles.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
3528               drBalanceHoles.Sections(1).Controls("lblSN").Caption = txtSN
3529               drBalanceHoles.Sections(1).Controls("lblToday").Caption = Now
3530               drBalanceHoles.Sections(1).Controls("lblVersion").Caption = strVersion
3531               drBalanceHoles.Show
' <VB WATCH>
3532       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3533               Exit Sub

3534           Case 8  'export to excel
3535               ExportToExcel
' <VB WATCH>
3536       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3537               Exit Sub

3538           Case 9  'Customer no circ flow no axial
3539               Set dr = drCustNoCircNoAxial

3540           Case 10  'Customer vibration no axial
3541               Set dr = drCustVibNoAxial

3542           Case 11 ' TEMC Inspection Report
3543               Set dr = drTEMCInspectionSheet
3544       End Select

3545       If Index = 11 Then
3546           dr.Sections(1).Controls("lblOutlineDimensions").Caption = "Are the following dimensions in accordance with those shown on the outline drawing, with tolerances per procedure A-29368?" & Chr(13) & "*   Base anchor bolt hole bolt circle pitch" & Chr(13) & "*   Base anchor bolt hole diameter" & Chr(13) & "*   Suction and discharge flange bolt hole bolt circle pitch" & Chr(13) & "*   Suction and discharge flange bolt hole diameter" & Chr(13) & "*   Face-to-face dimension between suction and discharge flange faces" & Chr(13) & "*   Location dimensions for all other fluid connections"
3547           dr.Sections(1).Controls("lblMotorNoLoadTest").Caption = "Record motor input current and power: **_________A **_______kW" & Chr(13) & "Confirm test was completed"
3548           dr.Sections(1).Controls("lblMotorLockedRotorTest").Caption = "Record motor input power at rated current: **_________A **_______kW" & Chr(13) & "Confirm test was completed"
3549           dr.Sections(1).Controls("lblHydraulicTest").Caption = "Are hydraulic test results per A-15852 acceptable?  Specifically, are the following within acceptable limits?" & Chr(13) & "*   rated head at rated flow" & Chr(13) & "*   input power" & Chr(13) & "*   input current" & Chr(13) & "*   TRG reading" & Chr(13) & "*   Axial thrust force/P-V" & Chr(13) & "*   Vibration" & Chr(13) & "(Note: test data recorded through lab data acquisition system.)"
3550           dr.Sections(1).Controls("lblNPSHTest").Caption = "Is NPSH required at design flow no greater than the NPSH required value provided to the customer?" & Chr(13) & "(Note: test data recorded through lab data acquisition system.)"
3551           dr.Sections(1).Controls("lblCustomer").Caption = txtShpNo
3552           dr.Sections(1).Controls("lblJobNo").Caption = txtJobNum.Text
3553           If InStr(txtSalesOrderNumber.Text, "/") <> 0 Then
3554               dr.Sections(1).Controls("lblOrderNo").Caption = Left$(txtSalesOrderNumber.Text, InStr(txtSalesOrderNumber.Text, "/") - 1)
3555               dr.Sections(1).Controls("lblItemNo").Caption = Right$(txtSalesOrderNumber.Text, Len(txtSalesOrderNumber.Text) - InStr(txtSalesOrderNumber.Text, "/"))
3556           Else
3557               dr.Sections(1).Controls("lblOrderNo").Caption = txtSalesOrderNumber.Text
3558               dr.Sections(1).Controls("lblItemNo").Caption = "****"
3559           End If
3560           dr.Sections(1).Controls("lblProductNo").Caption = txtSN
3561           dr.Sections(1).Controls("lblType").Caption = txtModelNo
3562           dr.Sections(1).Controls("lblDateInspected").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
3563           dr.Sections(1).Controls("lblFreq").Caption = cmbFrequency.List(cmbFrequency.ListIndex)
3564           dr.Sections(1).Controls("lblVolts").Caption = cmbVoltage.List(cmbVoltage.ListIndex)
3565           dr.Sections(1).Controls("lblFt").Caption = Format(Val(txtDesignTDH) * 0.3048, "###,##0.00")
3566           dr.Sections(1).Controls("lblCapacityM").Caption = Format(Val(txtDesignFlow) * 0.2271247, "##,##0.00")
3567           dr.Sections(1).Controls("lblCapacityG").Caption = txtDesignFlow

3568           dr.Sections(1).Controls("lblInsulationResistance").Caption = Val(txtTestAndInspection(0)) & " V Megger " & Val(txtTestAndInspection(1)) & " MOhms Above "
3569           dr.Sections(1).Controls("lblDielectricStrength").Caption = "AC " & Val(txtTestAndInspection(2)) & " X  " & Val(txtTestAndInspection(3)) & " min "
3570           dr.Sections(1).Controls("lblHydrostaticTest").Caption = Val(txtTestAndInspection(4)) & " " & cmbTestAndInspection(0).Text & " X " & Val(txtTestAndInspection(5)) & " min "
3571           dr.Sections(1).Controls("lblPneumaticTest").Caption = Val(txtTestAndInspection(6)) & " " & cmbTestAndInspection(1).Text & " X " & Val(txtTestAndInspection(7)) & " min "

3572           For I = 1 To 4
3573               If TestAndInspectionGood(I - 1).value = 1 Then
3574                   dr.Sections(1).Controls("lblGood" & Trim(str(I))).Caption = "Good _X__"
3575               Else
3576                   dr.Sections(1).Controls("lblGood" & Trim(str(I))).Caption = "Good ____"
3577               End If
3578           Next I

               'print initials and date
3579           For I = 1 To 11
3580               If I < 11 Then  '11 is supervisor yes/no
3581                   dr.Sections(1).Controls("lblInitials" & Trim(str(I))).Caption = txtWho.Text
3582                   dr.Sections(1).Controls("lbldate" & Trim(str(I))).Caption = Date
3583               End If
3584               If TestAndInspectionGood(I + 3).value = 1 Then
3585                   dr.Sections(1).Controls("lblYesNo" & Trim(str(I))).Caption = "Yes_X No__"
3586               Else
3587                   dr.Sections(1).Controls("lblYesNo" & Trim(str(I))).Caption = "Yes__ No_X"
3588               End If
3589           Next I

               'fix
3590           dr.Sections(1).Controls("lblPhase").Caption = txtNoPhases.Text
3591           dr.Sections(1).Controls("lblNPSHReq").Caption = txtNPSHr.Text
3592           dr.Sections(1).Controls("lblRatedOutput").Caption = txtRatedInputPower.Text
3593           dr.Sections(1).Controls("lblLiquid").Caption = txtLiquid.Text
3594           dr.Sections(1).Controls("lblAmps").Caption = txtAmps.Text
3595           dr.Sections(1).Controls("lblThermalClass").Caption = txtThermalClass.Text
3596           dr.Sections(1).Controls("lblViscosity").Caption = txtViscosity.Text
3597           dr.Sections(1).Controls("lblEXPClass").Caption = txtExpClass.Text
3598           dr.Sections(1).Controls("lblLiquidTemp").Caption = txtLiquidTemperature.Text

               'fix
3599           If txtSN <> "" Then
3600               Select Case optMfr(0).value
                       Case True
       '                    dr.Sections(1).Controls("lblType").Caption = ""
3601                       dr.Sections(1).Controls("lblPole").Caption = IIf(cmbRPM.List(cmbRPM.ListIndex) = 3450, "2", "4")
3602                   Case False
       '                    dr.Sections(1).Controls("lblType").Caption = cmbTEMCModel.List(cmbTEMCModel.ListIndex)
3603                       dr.Sections(1).Controls("lblPole").Caption = IIf(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1) = "1", "2", "4")
3604               End Select
3605           End If
3606           dr.Sections(1).Controls("lblSG").Caption = txtSpGr

       '        Set dr.Sections(2).Controls("Image2").Picture = CWGraphKw.ControlImage
       '        Set dr.Sections(2).Controls("Image3").Picture = CWGraphEff.ControlImage
       '        Set dr.Sections(2).Controls("Image4").Picture = CWGraphTDH.ControlImage
       '        Set dr.Sections(2).Controls("Image5").Picture = CWGraphAmps.ControlImage

3607       End If

3608       If Not isChart And Index <> 11 Then
       '        dr.Sections(1).Controls("lblTestDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
3609           dr.Sections(1).Controls("lblCustomer").Caption = txtShpNo
3610           dr.Sections(1).Controls("lblmodel").Caption = txtModelNo
3611           dr.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
3612           dr.Sections(1).Controls("lblSN").Caption = txtSN
3613           dr.Sections(1).Controls("lblGPM").Caption = txtDesignFlow
3614           dr.Sections(1).Controls("lblFt").Caption = txtDesignTDH
       '        dr.Sections(1).Controls("lblBaroPress").Caption = txtInHgDisplay
3615           dr.Sections(1).Controls("lblSuctGageHt").Caption = txtSuctHeight
3616           dr.Sections(1).Controls("lblDischGageHt").Caption = txtDischHeight
3617           dr.Sections(1).Controls("lblVersion").Caption = strVersion

3618           If chkTrimmed.value = 1 Then
3619               If Val(txtImpTrim.Text) <> 0 Then
3620                   dr.Sections(1).Controls("lblImpDia").Caption = txtImpTrim
3621               Else
3622                   dr.Sections(1).Controls("lblImpDia").Caption = txtImpellerDia
3623               End If
3624           Else
3625               dr.Sections(1).Controls("lblImpDia").Caption = txtImpellerDia
3626           End If
3627           Dim stemp As String
3628           dr.Sections(1).Controls("lblRPM").Caption = IIf(optMfr(0).value = True, cmbRPM.List(cmbRPM.ListIndex), IIf(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1) = "1", "3450", "1750"))
3629           dr.Sections(1).Controls("lblSpGr").Caption = txtSpGr
3630           dr.Sections(1).Controls("lblMotor").Caption = IIf(optMfr(0).value = True, cmbMotor.List(cmbMotor.ListIndex), txtTEMCFrameNumber.Text)
       '        dr.Sections(1).Controls("lblMotor").Caption = cmbMotor.List(cmbMotor.ListIndex)
3631           If Len(cmbTEMCVoltage.List(cmbTEMCVoltage.ListIndex)) > 6 Then
3632               stemp = Right$(cmbTEMCVoltage.List(cmbTEMCVoltage.ListIndex), Len(cmbTEMCVoltage.List(cmbTEMCVoltage.ListIndex)) - 6)
3633           Else
3634               stemp = ""
3635           End If
3636           dr.Sections(1).Controls("lblVoltage").Caption = IIf(optMfr(0).value = True, cmbVoltage.List(cmbVoltage.ListIndex), stemp)
3637           dr.Sections(1).Controls("lblEndPlay").Caption = txtEndPlay
3638           If Len(cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)) > 6 Then
3639               stemp = Right$(cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex), Len(cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)) - 6)
3640           Else
3641               stemp = ""
3642           End If
3643           dr.Sections(1).Controls("lblDesPressure").Caption = IIf(optMfr(0).value = True, cmbDesignPressure.List(cmbDesignPressure.ListIndex), stemp)
3644           dr.Sections(1).Controls("lblStatorFill").Caption = IIf(optMfr(0).value = True, cmbStatorFill.List(cmbStatorFill.ListIndex), "Dry")
3645           If Len(cmbTEMCModel.List(cmbTEMCModel.ListIndex)) > 11 Then
3646               stemp = Right$(cmbTEMCModel.List(cmbTEMCModel.ListIndex), Len(cmbTEMCModel.List(cmbTEMCModel.ListIndex)) - 11)
3647           Else
3648               stemp = ""
3649           End If

3650           dr.Sections(1).Controls("lblCircPath").Caption = IIf(optMfr(0).value = True, cmbCirculationPath.List(cmbCirculationPath.ListIndex), stemp)

3651           dr.Sections(1).Controls("lblKWMult").Caption = txtKWMult
3652           If Val(txtHDCor) = 0 Then
3653               dr.Sections(1).Controls("lblHDCor").Caption = 0
3654           Else
3655               dr.Sections(1).Controls("lblHDCor").Caption = txtHDCor
3656           End If
3657           dr.Sections(1).Controls("lblSuctPipeDia").Caption = cmbSuctDia.List(cmbSuctDia.ListIndex)
3658           dr.Sections(1).Controls("lblDischPipeDia").Caption = cmbDischDia.List(cmbDischDia.ListIndex)
3659           dr.Sections(1).Controls("lblTestSpec").Caption = cmbTestSpec.List(cmbTestSpec.ListIndex)

3660           dr.Sections(1).Controls("lblToday").Caption = Now

3661           dr.Sections(1).Controls("lblSuctionID").Caption = txtSuctionID
3662           dr.Sections(1).Controls("lblDischargeID").Caption = txtDischargeID
3663           dr.Sections(1).Controls("lblTempID").Caption = txtTemperatureID
3664           dr.Sections(1).Controls("lblCircFlowID").Caption = txtMagflowID
3665           dr.Sections(1).Controls("lblFlowID").Caption = txtFlowmeterID

3666           dr.Sections(1).Controls("lblAnalyzerID").Caption = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)
3667           dr.Sections(1).Controls("lblLoopID").Caption = cmbLoopNumber.List(cmbLoopNumber.ListIndex)
3668           dr.Sections(1).Controls("lblTachID").Caption = cmbTachID.List(cmbTachID.ListIndex)
3669           dr.Sections(1).Controls("lblOrificeID").Caption = cmbOrificeNumber.List(cmbOrificeNumber.ListIndex)
3670           dr.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)

3671           If chkFeathered.value = 1 Then
3672               dr.Sections(1).Controls("lblImpFeathered").Visible = True
3673           Else
3674               dr.Sections(1).Controls("lblImpFeathered").Visible = False
3675           End If

3676           If chkOrifice.value = 1 Then
3677               dr.Sections(1).Controls("lblDischOrifice").Visible = True
3678               dr.Sections(1).Controls("lblDischOrificeValue").Visible = True
3679               dr.Sections(1).Controls("lblDischOrificeValue").Caption = txtOrifice
3680           Else
3681               dr.Sections(1).Controls("lblDischOrifice").Visible = False
3682               dr.Sections(1).Controls("lblDischOrificeValue").Visible = False
3683           End If

3684           If chkCircOrifice.value = 1 Then
3685               dr.Sections(1).Controls("lblCircFlowOrifice").Visible = True
3686               dr.Sections(1).Controls("lblCircFlowOrificeValue").Visible = True
3687               dr.Sections(1).Controls("lblCircFlowOrificeValue").Caption = frmPLCData.txtCircOrifice
3688           Else
3689               dr.Sections(1).Controls("lblCircFlowOrifice").Visible = False
3690               dr.Sections(1).Controls("lblCircFlowOrificeValue").Visible = False
3691           End If

3692           dr.Sections(4).Controls("lblOther").Caption = txtOtherMods
3693           dr.Sections(4).Controls("lblRemarks").Caption = txtRemarks

3694           dr.Sections(4).Controls("lblTestRunRemarks").Caption = txtTestSetupRemarks

3695           dr.Orientation = rptOrientLandscape
3696       End If

3697       Printer.Orientation = vbPRORLandscape

       '    If DataEnvironment1.rsEff.State = adStateOpen Then
       '        DataEnvironment1.rsEff.Close
       '    End If
       '    DataEnvironment1.rsEff.Open
       '    DataEnvironment1.rsEff.MoveFirst

       '    DataEnvironment1.Commands.Item(1).Parameters(0).Value = 7

       '    DataEnvironment1.rsEff.Requery

3698       Dim dm As String
3699       Dim RsNum As Integer
3700       rsEff.Requery
3701       RsNum = 9 - UpDown2.value

3702       dm = "Recs" & Trim(str(frmPLCData.UpDown2.value))

3703       Select Case Index
               Case 0
3704               Set drCustNoCirc.DataSource = DataEnvironment1
3705               drCustNoCirc.DataMember = dm
3706               DoEvents
3707               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3708                   DataEnvironment1.Recordsets(RsNum).Open
3709               End If
3710               DataEnvironment1.Recordsets(RsNum).Requery
3711               dr.Sections(4).Controls("lblBH").Caption = strBH
3712               DoEvents
3713               drCustNoCirc.Show
3714           Case 1
3715               Set drCustCirc.DataSource = DataEnvironment1
3716               drCustCirc.DataMember = dm
3717               DoEvents
3718               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3719                   DataEnvironment1.Recordsets(RsNum).Open
3720               End If
3721               DataEnvironment1.Recordsets(RsNum).Requery
3722               dr.Sections(4).Controls("lblBH").Caption = strBH
3723               DoEvents
3724               drCustCirc.Show
3725           Case 2
3726               Set drCustVib.DataSource = DataEnvironment1
3727               drCustVib.DataMember = dm
3728               DoEvents
3729               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3730                   DataEnvironment1.Recordsets(RsNum).Open
3731               End If
3732               DataEnvironment1.Recordsets(RsNum).Requery
3733               dr.Sections(4).Controls("lblBH").Caption = strBH
3734               DoEvents
3735               drCustVib.Show
3736           Case 3
3737               Set drInternal.DataSource = DataEnvironment1
3738               drInternal.DataMember = dm
3739               DoEvents
3740               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3741                   DataEnvironment1.Recordsets(RsNum).Open
3742               End If
3743               DataEnvironment1.Recordsets(RsNum).Requery
3744               dr.Sections(4).Controls("lblBH").Caption = strBH
3745               DoEvents
3746               drInternal.Show
3747               Set drInternal2.DataSource = DataEnvironment1
3748               drInternal2.DataMember = dm
3749               DoEvents
3750               drInternal2.Show
3751           Case 4
3752               Set drAnalysis.DataSource = DataEnvironment1
3753               drAnalysis.DataMember = dm
3754               DoEvents
3755               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3756                   DataEnvironment1.Recordsets(RsNum).Open
3757               End If
3758               DataEnvironment1.Recordsets(RsNum).Requery
3759               DoEvents
3760               drAnalysis.Show
3761           Case 5
3762               Set drChart.DataSource = DataEnvironment1
3763               drChart.DataMember = dm
3764               DoEvents
3765               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3766                   DataEnvironment1.Recordsets(RsNum).Open
3767               End If
3768               DataEnvironment1.Recordsets(RsNum).Requery
3769               DoEvents
3770               drChart.Show
3771           Case 6

3772           Case 9
3773               Set drCustNoCircNoAxial.DataSource = DataEnvironment1
3774               drCustNoCircNoAxial.DataMember = dm
3775               DoEvents
3776               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3777                   DataEnvironment1.Recordsets(RsNum).Open
3778               End If
3779               DataEnvironment1.Recordsets(RsNum).Requery
3780               dr.Sections(4).Controls("lblBH").Caption = strBH
3781               DoEvents
3782               drCustNoCircNoAxial.Show
3783           Case 10
3784               Set drCustVibNoAxial.DataSource = DataEnvironment1
3785               drCustVibNoAxial.DataMember = dm
3786               DoEvents
3787               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3788                   DataEnvironment1.Recordsets(RsNum).Open
3789               End If
3790               DataEnvironment1.Recordsets(RsNum).Requery
3791               dr.Sections(4).Controls("lblBH").Caption = strBH
3792               DoEvents
3793               drCustVibNoAxial.Show
3794           Case 11
3795               Printer.Orientation = vbPRORLandscape
3796               Printer.PaperSize = vbPRPSLetter
3797               drTEMCInspectionSheet.Orientation = rptOrientPortrait
3798               drTEMCInspectionSheet.Height = 16000
3799               drTEMCInspectionSheet.RightMargin = 0
3800               drTEMCInspectionSheet.LeftMargin = 0
3801               drTEMCInspectionSheet.TopMargin = 0
3802               drTEMCInspectionSheet.BottomMargin = 0
3803               Set drTEMCInspectionSheet.DataSource = DataEnvironment1
3804               drTEMCInspectionSheet.DataMember = dm
3805               DoEvents
3806               If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
3807                   DataEnvironment1.Recordsets(RsNum).Open
3808               End If
3809               DataEnvironment1.Recordsets(RsNum).Requery
       '            dr.Sections(4).Controls("lblBH").Caption = strBH
3810               DoEvents

3811               drTEMCInspectionSheet.Show
3812       End Select

       '    rsEff.Close
       '    rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

       '    rsEff.Requery

' <VB WATCH>
3813       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3814       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "optReport_Click"

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
            vbwReportVariable "Index", Index
            vbwReportVariable "strBH", strBH
            vbwReportVariable "I", I
            vbwReportVariable "strVersion", strVersion
            vbwReportVariable "isChart", isChart
            vbwReportVariable "PosRPM", PosRPM
            vbwReportVariable "PosAxPos", PosAxPos
            vbwReportVariable "PosCircFlow", PosCircFlow
            vbwReportVariable "PosVib", PosVib
            vbwReportVariable "PosRem", PosRem
            vbwReportVariable "PosTRG", PosTRG
            vbwReportVariable "stemp", stemp
            vbwReportVariable "dm", dm
            vbwReportVariable "RsNum", RsNum
            vbwReportVariable "dr", dr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub tmrGetDDE_Timer()
' <VB WATCH>
3815       On Error GoTo vbwErrHandler
3816       Const VBWPROCNAME = "frmPLCData.tmrGetDDE_Timer"
3817       If vbwProtector.vbwTraceProc Then
3818           Dim vbwProtectorParameterString As String
3819           If vbwProtector.vbwTraceParameters Then
3820               vbwProtectorParameterString = "()"
3821           End If
3822           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3823       End If
' </VB WATCH>

       'get here every second... get plc and magtrol data

3824       Dim sSendStr As String
3825       Dim I As Integer
3826       Dim VoltMul As Double

3827       If Calibrating Then
' <VB WATCH>
3828       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3829           Exit Sub
3830       End If

3831       If debugging Then
               'Exit Sub
3832       End If

3833       If boPLCOperating = True Then
3834           frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on
3835           DoEvents

               'convert the plc data into real numbers
               'the following data are type real
3836           txtFlow.Text = ConvertToReal("4050")
3837           txtSuction.Text = ConvertToReal("4052")
3838           txtDischarge.Text = ConvertToReal("4054")
3839           txtTemperature.Text = ConvertToReal("4056")

3840           txtValvePosition.Text = ConvertToLong("2004")

3841           frmPLCData.txtTC1.Text = ConvertToLong("2200")
3842           frmPLCData.txtTC2.Text = ConvertToLong("2202")
3843           frmPLCData.txtTC3.Text = ConvertToLong("2204")
3844           frmPLCData.txtTC4.Text = ConvertToLong("2206")

3845           frmPLCData.txtAI1.Text = ConvertToReal("4060")
3846           frmPLCData.txtAI2.Text = ConvertToReal("4062")
3847           frmPLCData.txtAI3.Text = ConvertToReal("4064")
3848           frmPLCData.txtAI4.Text = ConvertToReal("4066")

3849           frmPLCData.txtPCoef.Text = ConvertToLong("4036")
3850           frmPLCData.txtICoef.Text = ConvertToLong("4037")
3851           frmPLCData.txtDCoef.Text = ConvertToLong("4040")

3852           frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
3853           frmPLCData.txtInHg.Text = ConvertToLong("1460")


               'modify the data from PLC format to format that we can use
               'and update the screen
3854           If txtFlowDisplay.Enabled = False Then
3855               frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
3856           End If
3857           If txtSuctionDisplay.Enabled = False Then
3858               frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
3859           End If
3860           If txtDischargeDisplay.Enabled = False Then
3861               frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
3862           End If
3863           If txtTemperatureDisplay.Enabled = False Then
3864               frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
3865           End If
3866           frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

3867           frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
3868           frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
3869           frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
3870           frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

3871           If txtAI1Display.Enabled = False Then
3872               frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
3873           End If
3874           If txtAI2Display.Enabled = False Then
3875               frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
3876           End If
3877           If txtAI3Display.Enabled = False Then
3878               frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
3879           End If
3880           If txtAI4Display.Enabled = False Then
3881               frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
3882           End If

3883           frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

3884           frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

3885           frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off
3886           DoEvents

3887           frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
3888           DoEvents
3889       End If

3890       If boMagtrolOperating = True Then


               'get the data from the Magtrol
3891           If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
3892               sSendStr = vbCrLf
3893               sData = Space$(68)
3894               VoltMul = Sqr(3)
3895           Else
3896               sSendStr = "OT" & vbCrLf
3897               sData = Space$(183)
3898               VoltMul = 1#
3899           End If
3900           ibwrt iUD, sSendStr
3901           ibrd iUD, sData

               'parse the Magrol response
               'vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)

3902               Dim vSplit() As String
3903               vSplit = Split(Right(sData, Len(sData) - 1), ",")
3904               If UBound(vSplit) > 0 Then
3905                   ReDim vResponse(UBound(vSplit))
3906               End If
3907               For I = 0 To UBound(vSplit) - 1
3908                   If Len(vSplit(I)) <> 0 Then
3909                       vResponse(I) = CDbl(vSplit(I))
3910                   End If
3911               Next I

               'format the parsed response
3912           Dim dd As String
3913           dd = "- -"

3914           On Error GoTo noresponse
3915           If Not IsEmpty(vResponse) Then
               '8 entries for 5300 and 12 for the 6530
3916               If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                       'put the responses into the correct text box
3917                   txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

3918                   Select Case vResponse(0)
                           Case Is < 1
3919                           txtI1.Text = Format$(vResponse(0), "0.0000")
3920                       Case Is < 10
3921                           txtI1.Text = Format$(vResponse(0), "0.000")
3922                       Case Is < 100
3923                           txtI1.Text = Format$(vResponse(0), "00.00")
3924                       Case Else
3925                           txtI1.Text = Format$(vResponse(0), "000.0")
3926                   End Select

3927                   Select Case vResponse(3)
                           Case Is < 1
3928                           txtI2.Text = Format$(vResponse(3), "0.0000")
3929                       Case Is < 10
3930                           txtI2.Text = Format$(vResponse(3), "0.000")
3931                       Case Is < 100
3932                           txtI2.Text = Format$(vResponse(3), "00.00")
3933                       Case Else
3934                           txtI2.Text = Format$(vResponse(3), "000.0")
3935                   End Select

3936                   Select Case vResponse(6)
                           Case Is < 1
3937                           txtI3.Text = Format$(vResponse(6), "0.0000")
3938                       Case Is < 10
3939                           txtI3.Text = Format$(vResponse(6), "0.000")
3940                       Case Is < 100
3941                           txtI3.Text = Format$(vResponse(6), "00.00")
3942                       Case Else
3943                           txtI3.Text = Format$(vResponse(6), "000.0")
3944                   End Select

3945                   txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
3946                   txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                       'txtI2.Text = Format$(vResponse(3), "###0.0")
3947                   txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
3948                   txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                       'txtI3.Text = Format$(vResponse(6), "###0.0")
3949                   txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
3950                   If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                           'if we have some measured current
                           'pf = sum of power/sum of VA
3951                       If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                               'add kw responses and / by 1000 to get to kW
3952                           txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
3953                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
3954                       Else
3955                           txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
3956                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
3957                       End If
3958                       Select Case Val(txtKW.Text)
                               Case Is < 1
3959                               txtKW.Text = Format$(txtKW.Text, "0.00000")
3960                           Case Is < 10
3961                               txtKW.Text = Format$(txtKW.Text, "0.0000")
3962                           Case Is < 100
3963                               txtKW.Text = Format$(txtKW.Text, "00.000")
3964                           Case Else
3965                               txtKW.Text = Format$(txtKW.Text, "000.00")
3966                       End Select
3967                   Else
3968                       txtPF = dd
3969                   End If
3970               Else
                       'no response, show all -- in text boxes
3971                   txtV1.Text = dd
3972                   txtI1.Text = dd
3973                   txtP1.Text = dd
3974                   txtV2.Text = dd
3975                   txtI2.Text = dd
3976                   txtP2.Text = dd
3977                   txtV3.Text = dd
3978                   txtI3.Text = dd
3979                   txtP3.Text = dd
3980                   txtPF = dd
3981                   txtKW = dd
3982               End If
       '            If UBound(vResponse) = 8 Then
       '                txtV1.Text = Format$(Sqr(3) * vResponse(1), "###0.0")
       '                txtI1.Text = Format$(vResponse(0), "###0.0")
       '                txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")
       '                txtV2.Text = Format$(Sqr(3) * vResponse(4), "###0.0")
       '                txtI2.Text = Format$(vResponse(3), "###0.0")
       '                txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
       '                txtV3.Text = Format$(Sqr(3) * vResponse(7), "###0.0")
       '                txtI3.Text = Format$(vResponse(6), "###0.0")
       '                txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
       '                If (vResponse(0) + vResponse(3) + vResponse(6)) <> 0 Then
       '                    txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
       '                Else
       '                    txtPF = dd
       '                End If
       '                txtKW = Format$((vResponse(2) + vResponse(5) + vResponse(8)) / 1000, "###0.0")
       '            Else
       '                txtV1.Text = dd
       '                txtI1.Text = dd
       '                txtP1.Text = dd
       '                txtV2.Text = dd
       '                txtI2.Text = dd
       '                txtP2.Text = dd
       '                txtV3.Text = dd
       '                txtI3.Text = dd
       '                txtP3.Text = dd
       '                txtPF = dd
       '                txtKW = dd
       '            End If
3983           End If
3984       Else    'magtrol not operating
3985           Dim dbl As Double

3986           If optKW(0).value = True Then   'add 3 powers
3987               txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
3988           End If
3989           If optKW(1).value = True Then   'enter kw
3990               txtP1.Text = Val(txtKW.Text) / 3
3991               txtP2.Text = Val(txtKW.Text) / 3
3992               txtP3.Text = Val(txtKW.Text) / 3
3993           End If
3994           If optKW(2).value = True Then   'use ai4
3995               txtKW.Text = txtAI4Display.Text
3996               txtP1.Text = Val(txtKW.Text) / 3
3997               txtP2.Text = Val(txtKW.Text) / 3
3998               txtP3.Text = Val(txtKW.Text) / 3
3999           End If

4000           dbl = Val(txtV1.Text) * Val(txtI1.Text)
4001           dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
4002           dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
4003           If dbl <> 0 Then
4004               txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
4005           End If
4006       End If

4007   noresponse:
4008   On Error GoTo vbwErrHandler
4009       frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off
4010       DoEvents

           'update the little PLC chart
4011       For I = 0 To 19
4012           vPlot(I, 0) = vPlot(I + 1, 0)
4013           vPlot(I, 1) = vPlot(I + 1, 1)
4014       Next I
4015       vPlot(20, 0) = Val(txtSetPointDisplay.Text)
4016       vPlot(20, 1) = Val(txtFlowDisplay.Text)

       '    If Not (txtSetPointDisplay.Text = "" Or txtFlowDisplay.Text = "") Then
       '       If vPlot(0, 0) <> Empty Then
4017               MSChart2 = vPlot

4018               MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
4019               MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(vPlot) / 10) + 0.5) + 1)
4020               MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
4021               MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

       '            CWGraph1.PlotY vPlot
       '        End If
       '    End If

           'do NPSH stuff
4022       Dim SuctVelHead As Single
4023       Dim DischVelHead As Single
4024       Dim Conversion As Single
4025       Dim SuctionPSIA As Single
4026       Dim DischargePSIA As Single
4027       Dim VaporPress As Single
4028       Dim SpecVolume As Single
4029       Dim NPSHa As Single
4030       Dim TDH As Single
4031       Dim pd As Single


           'velocity head
4032       If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
4033           SuctVelHead = 0
4034       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
4035           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
4036           SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
4037       End If

4038       If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
4039           DischVelHead = 0
4040       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
4041           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
4042           DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
4043       End If

           'convert gauges to absolute
4044       If txtInHgDisplay.Text = "" Then
4045           Conversion = 0
4046       Else
4047           Conversion = txtInHgDisplay * 0.491
4048       End If

4049       SuctionPSIA = Val(txtSuctionDisplay) + Conversion
4050       DischargePSIA = Val(txtDischargeDisplay) + Conversion


           'lookup vapor pressure and specific volume in the arrays that we made
           'if temp is out of range, say so and exit
4051       If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
4052           txtNPSHa = 0
' <VB WATCH>
4053       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4054           Exit Sub
4055       Else
4056           I = Val(txtTemperatureDisplay) - 40
       '        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
       '        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
4057           VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
4058           SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
4059       End If

4060       If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
               'NPSHa
4061           NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
       '        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
4062           txtNPSHa = Format$(NPSHa, "##0.00")

               'tdh
4063           TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
4064           txtTDH = Format$(TDH, "##0.00")

4065       Else
4066           txtNPSHa = 0
4067       End If
' <VB WATCH>
4068       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4069       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tmrGetDDE_Timer"

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
            vbwReportVariable "sSendStr", sSendStr
            vbwReportVariable "I", I
            vbwReportVariable "VoltMul", VoltMul
            vbwReportVariable "vSplit", vSplit
            vbwReportVariable "dd", dd
            vbwReportVariable "dbl", dbl
            vbwReportVariable "SuctVelHead", SuctVelHead
            vbwReportVariable "DischVelHead", DischVelHead
            vbwReportVariable "Conversion", Conversion
            vbwReportVariable "SuctionPSIA", SuctionPSIA
            vbwReportVariable "DischargePSIA", DischargePSIA
            vbwReportVariable "VaporPress", VaporPress
            vbwReportVariable "SpecVolume", SpecVolume
            vbwReportVariable "NPSHa", NPSHa
            vbwReportVariable "TDH", TDH
            vbwReportVariable "pd", pd
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub tmrStartUp_Timer()
           'we waited for a while, disable the timer
' <VB WATCH>
4070       On Error GoTo vbwErrHandler
4071       Const VBWPROCNAME = "frmPLCData.tmrStartUp_Timer"
4072       If vbwProtector.vbwTraceProc Then
4073           Dim vbwProtectorParameterString As String
4074           If vbwProtector.vbwTraceParameters Then
4075               vbwProtectorParameterString = "()"
4076           End If
4077           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4078       End If
' </VB WATCH>
4079       tmrStartUp.Enabled = False
' <VB WATCH>
4080       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4081       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tmrStartUp_Timer"

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
Public Function SetCombo(cmbComboName As ComboBox, sName As String, rs As ADODB.Recordset)
       'set the pump parameter combo box to the right data based upon
       'the number in the database
' <VB WATCH>
4082       On Error GoTo vbwErrHandler
4083       Const VBWPROCNAME = "frmPLCData.SetCombo"
4084       If vbwProtector.vbwTraceProc Then
4085           Dim vbwProtectorParameterString As String
4086           If vbwProtector.vbwTraceParameters Then
4087               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
4088               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sName", sName) & ", "
4089               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
4090           End If
4091           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4092       End If
' </VB WATCH>

4093       Dim I As Integer
4094       Dim sParam As String
4095       Dim qy As New ADODB.Command
4096       Dim rs1 As New ADODB.Recordset

4097       If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
4098           sParam = rs.Fields(sName)                'get the index number
4099           qy.ActiveConnection = cnPumpData
4100           qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
4101           Set rs1 = qy.Execute()                                  'get the record for the index number

4102           If rs1.BOF = True And rs1.EOF = True Then
4103               cmbComboName.ListIndex = -1                             'else, remove any pointer
' <VB WATCH>
4104       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4105               Exit Function
4106           End If

4107           For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
4108               If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
4109                   cmbComboName.ListIndex = I                                              'if we do, set the combo box
4110                   Exit For                                            'and we're done
4111               End If
4112               cmbComboName.ListIndex = -1                             'else, remove any pointer
4113           Next I
4114       Else
4115           cmbComboName.ListIndex = -1
4116       End If

' <VB WATCH>
4117       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4118       Exit Function
' <VB WATCH>
4119       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4120       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetCombo"

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
            vbwReportVariable "sName", sName
            vbwReportVariable "I", I
            vbwReportVariable "sParam", sParam
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportVariable "rs", rs
            vbwReportVariable "qy", qy
            vbwReportVariable "rs1", rs1
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Private Function SetComboTestSetup(cmbComboName As ComboBox, sFieldName As String, sTableName As String, rs As ADODB.Recordset)
       'set the pump parameter combo box to the right data based upon
       'the number in the database
' <VB WATCH>
4121       On Error GoTo vbwErrHandler
4122       Const VBWPROCNAME = "frmPLCData.SetComboTestSetup"
4123       If vbwProtector.vbwTraceProc Then
4124           Dim vbwProtectorParameterString As String
4125           If vbwProtector.vbwTraceParameters Then
4126               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
4127               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sFieldName", sFieldName) & ", "
4128               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ", "
4129               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
4130           End If
4131           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4132       End If
' </VB WATCH>

       'same as setcombo, except here we also pass in the field name

4133       FromStoredData = True

4134       Dim I As Integer
4135       Dim sParam As String
4136       Dim qy As New ADODB.Command
4137       Dim rs1 As New ADODB.Recordset

4138       If rs.Fields(sFieldName).ActualSize <> 0 Then
4139           sParam = rs.Fields(sFieldName)
4140           qy.ActiveConnection = cnPumpData
4141           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
4142           Set rs1 = qy.Execute()

4143           For I = 0 To cmbComboName.ListCount - 1
4144               If cmbComboName.ItemData(I) = rs1.Fields(0) Then
4145                   cmbComboName.ListIndex = I
4146                   Exit For
4147               End If
4148               cmbComboName.ListIndex = -1
4149           Next I
4150       Else
4151           cmbComboName.ListIndex = -1
4152       End If

4153       FromStoredData = False

' <VB WATCH>
4154       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4155       Exit Function
' <VB WATCH>
4156       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4157       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetComboTestSetup"

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
            vbwReportVariable "sFieldName", sFieldName
            vbwReportVariable "sTableName", sTableName
            vbwReportVariable "I", I
            vbwReportVariable "sParam", sParam
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportVariable "rs", rs
            vbwReportVariable "qy", qy
            vbwReportVariable "rs1", rs1
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Private Sub DisablePumpDataControls()
           'disable the pump data controls cause we're just showing what we found
' <VB WATCH>
4158       On Error GoTo vbwErrHandler
4159       Const VBWPROCNAME = "frmPLCData.DisablePumpDataControls"
4160       If vbwProtector.vbwTraceProc Then
4161           Dim vbwProtectorParameterString As String
4162           If vbwProtector.vbwTraceParameters Then
4163               vbwProtectorParameterString = "()"
4164           End If
4165           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4166       End If
' </VB WATCH>

4167       txtSalesOrderNumber.Enabled = False
4168       frmMfr.Enabled = False
4169       txtShpNo.Enabled = False
4170       txtBilNo.Enabled = False
4171       txtDesignFlow.Enabled = False
4172       txtDesignTDH.Enabled = False

4173       frmMiscPumpData.Enabled = False

4174       txtModelNo.Enabled = False
4175       txtImpellerDia.Enabled = False

4176       frmTEMC.Enabled = False
4177       frmChempump.Enabled = False

4178       txtRemarks.Enabled = False
4179       Me.cmdAddNewTestDate.Visible = False

4180       cmdEnterPumpData.Enabled = False

' <VB WATCH>
4181       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4182       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisablePumpDataControls"

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
Private Sub DisableTestSetupDataControls()
' <VB WATCH>
4183       On Error GoTo vbwErrHandler
4184       Const VBWPROCNAME = "frmPLCData.DisableTestSetupDataControls"
4185       If vbwProtector.vbwTraceProc Then
4186           Dim vbwProtectorParameterString As String
4187           If vbwProtector.vbwTraceParameters Then
4188               vbwProtectorParameterString = "()"
4189           End If
4190           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4191       End If
' </VB WATCH>

4192       cmbTestSpec.Enabled = False
4193       txtWho.Enabled = False
4194       txtRMA.Enabled = False

4195       frmLoopAndXducer.Enabled = False
4196       frmElecData.Enabled = False
4197       frmPerfMods.Enabled = False
4198       frmOtherFiles.Enabled = False
4199       frmInstrumentTags.Enabled = False
4200       frmTAndI.Enabled = False
4201       frmThrustBalMods.Enabled = False
4202       txtTestSetupRemarks.Enabled = False

4203       cmdEnterTestSetupData.Enabled = False
4204       cmbPLCNo.Enabled = False
' <VB WATCH>
4205       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4206       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisableTestSetupDataControls"

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
Private Sub DisableTestDataControls()
' <VB WATCH>
4207       On Error GoTo vbwErrHandler
4208       Const VBWPROCNAME = "frmPLCData.DisableTestDataControls"
4209       If vbwProtector.vbwTraceProc Then
4210           Dim vbwProtectorParameterString As String
4211           If vbwProtector.vbwTraceParameters Then
4212               vbwProtectorParameterString = "()"
4213           End If
4214           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4215       End If
' </VB WATCH>

4216       cmbPLCLoop.Enabled = False
4217       frmPumpData.Enabled = False
4218       frmThermocouples.Enabled = False
4219       frmAI.Enabled = False
4220       frmMagtrol.Enabled = False
4221       fmrMiscTestData.Enabled = False
4222       frmPLCMisc.Enabled = False
4223       DataGrid1.Enabled = False
4224       DataGrid2.Enabled = False
4225       cmdEnterTestData.Enabled = False

' <VB WATCH>
4226       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4227       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisableTestDataControls"

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
Private Sub EnableTestSetupDataControls()
' <VB WATCH>
4228       On Error GoTo vbwErrHandler
4229       Const VBWPROCNAME = "frmPLCData.EnableTestSetupDataControls"
4230       If vbwProtector.vbwTraceProc Then
4231           Dim vbwProtectorParameterString As String
4232           If vbwProtector.vbwTraceParameters Then
4233               vbwProtectorParameterString = "()"
4234           End If
4235           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4236       End If
' </VB WATCH>

4237       cmbTestSpec.Enabled = True
4238       txtWho.Enabled = True
4239       txtRMA.Enabled = True

4240       frmLoopAndXducer.Enabled = True
4241       frmElecData.Enabled = True
4242       frmPerfMods.Enabled = True
4243       frmOtherFiles.Enabled = True
4244       frmInstrumentTags.Enabled = True
4245       frmTAndI.Enabled = True
4246       frmThrustBalMods.Enabled = True
4247       txtTestSetupRemarks.Enabled = True

4248       cmdEnterTestSetupData.Enabled = True
4249       cmbPLCNo.Enabled = True
' <VB WATCH>
4250       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4251       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableTestSetupDataControls"

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
Private Sub EnableTestDataControls()
' <VB WATCH>
4252       On Error GoTo vbwErrHandler
4253       Const VBWPROCNAME = "frmPLCData.EnableTestDataControls"
4254       If vbwProtector.vbwTraceProc Then
4255           Dim vbwProtectorParameterString As String
4256           If vbwProtector.vbwTraceParameters Then
4257               vbwProtectorParameterString = "()"
4258           End If
4259           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4260       End If
' </VB WATCH>

4261       cmbPLCLoop.Enabled = True
4262       frmPumpData.Enabled = True
4263       frmThermocouples.Enabled = True
4264       frmAI.Enabled = True
4265       frmMagtrol.Enabled = True
4266       fmrMiscTestData.Enabled = True
4267       frmPLCMisc.Enabled = True
4268       DataGrid1.Enabled = True
4269       DataGrid2.Enabled = True
4270       cmdEnterTestData.Enabled = True

' <VB WATCH>
4271       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4272       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableTestDataControls"

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
Private Sub EnablePumpDataControls()
           'disable the pump data controls cause we're just showing what we found
' <VB WATCH>
4273       On Error GoTo vbwErrHandler
4274       Const VBWPROCNAME = "frmPLCData.EnablePumpDataControls"
4275       If vbwProtector.vbwTraceProc Then
4276           Dim vbwProtectorParameterString As String
4277           If vbwProtector.vbwTraceParameters Then
4278               vbwProtectorParameterString = "()"
4279           End If
4280           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4281       End If
' </VB WATCH>

4282       txtSalesOrderNumber.Enabled = True
4283       frmMfr.Enabled = True
4284       txtShpNo.Enabled = True
4285       txtBilNo.Enabled = True
4286       txtDesignFlow.Enabled = True
4287       txtDesignTDH.Enabled = True

4288       frmMiscPumpData.Enabled = True

4289       txtModelNo.Enabled = True
4290       txtImpellerDia.Enabled = True

4291       frmTEMC.Enabled = True
4292       frmChempump.Enabled = True

4293       txtRemarks.Enabled = True
4294       Me.cmdAddNewTestDate.Visible = True

4295       cmdEnterPumpData.Enabled = True

' <VB WATCH>
4296       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4297       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnablePumpDataControls"

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
Private Sub EnableMagtrolFields()
' <VB WATCH>
4298       On Error GoTo vbwErrHandler
4299       Const VBWPROCNAME = "frmPLCData.EnableMagtrolFields"
4300       If vbwProtector.vbwTraceProc Then
4301           Dim vbwProtectorParameterString As String
4302           If vbwProtector.vbwTraceParameters Then
4303               vbwProtectorParameterString = "()"
4304           End If
4305           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4306       End If
' </VB WATCH>
4307       txtV1.Enabled = True
4308       txtV2.Enabled = True
4309       txtV3.Enabled = True
4310       txtI1.Enabled = True
4311       txtI2.Enabled = True
4312       txtI3.Enabled = True
4313       txtP1.Enabled = True
4314       txtP2.Enabled = True
4315       txtP3.Enabled = True
4316       optKW(0).Visible = True
4317       optKW(1).Visible = True
4318       optKW(2).Visible = True
4319       optKW(1).value = True
4320       optKW_Click (1)
' <VB WATCH>
4321       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4322       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableMagtrolFields"

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
Private Sub DisableMagtrolFields()
' <VB WATCH>
4323       On Error GoTo vbwErrHandler
4324       Const VBWPROCNAME = "frmPLCData.DisableMagtrolFields"
4325       If vbwProtector.vbwTraceProc Then
4326           Dim vbwProtectorParameterString As String
4327           If vbwProtector.vbwTraceParameters Then
4328               vbwProtectorParameterString = "()"
4329           End If
4330           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4331       End If
' </VB WATCH>
4332       txtV1.Enabled = False
4333       txtV2.Enabled = False
4334       txtV3.Enabled = False
4335       txtI1.Enabled = False
4336       txtI2.Enabled = False
4337       txtI3.Enabled = False
4338       txtP1.Enabled = False
4339       txtP2.Enabled = False
4340       txtP3.Enabled = False
4341       txtKW.Enabled = False
4342       optKW(0).Visible = False
4343       optKW(1).Visible = False
4344       optKW(2).Visible = False
' <VB WATCH>
4345       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4346       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisableMagtrolFields"

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
Private Sub EnablePLCFields()
' <VB WATCH>
4347       On Error GoTo vbwErrHandler
4348       Const VBWPROCNAME = "frmPLCData.EnablePLCFields"
4349       If vbwProtector.vbwTraceProc Then
4350           Dim vbwProtectorParameterString As String
4351           If vbwProtector.vbwTraceParameters Then
4352               vbwProtectorParameterString = "()"
4353           End If
4354           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4355       End If
' </VB WATCH>
4356       frmPLCData.txtAI1Display.Enabled = True
4357       frmPLCData.txtAI2Display.Enabled = True
4358       frmPLCData.txtAI3Display.Enabled = True
4359       frmPLCData.txtAI4Display.Enabled = True
4360       frmPLCData.txtTC1Display.Enabled = True
4361       frmPLCData.txtTC2Display.Enabled = True
4362       frmPLCData.txtTC3Display.Enabled = True
4363       frmPLCData.txtTC4Display.Enabled = True
4364       frmPLCData.txtFlowDisplay.Enabled = True
4365       frmPLCData.txtSuctionDisplay.Enabled = True
4366       frmPLCData.txtDischargeDisplay.Enabled = True
4367       frmPLCData.txtTemperatureDisplay.Enabled = True
4368       frmPLCData.txtInHgDisplay.Enabled = True
' <VB WATCH>
4369       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4370       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnablePLCFields"

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
Private Sub DisablePLCFields()
' <VB WATCH>
4371       On Error GoTo vbwErrHandler
4372       Const VBWPROCNAME = "frmPLCData.DisablePLCFields"
4373       If vbwProtector.vbwTraceProc Then
4374           Dim vbwProtectorParameterString As String
4375           If vbwProtector.vbwTraceParameters Then
4376               vbwProtectorParameterString = "()"
4377           End If
4378           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4379       End If
' </VB WATCH>
4380       frmPLCData.txtAI1Display.Enabled = False
4381       frmPLCData.txtAI2Display.Enabled = False
4382       frmPLCData.txtAI3Display.Enabled = False
4383       frmPLCData.txtAI4Display.Enabled = False
4384       frmPLCData.txtTC1Display.Enabled = False
4385       frmPLCData.txtTC2Display.Enabled = False
4386       frmPLCData.txtTC3Display.Enabled = False
4387       frmPLCData.txtTC4Display.Enabled = False
4388       frmPLCData.txtFlowDisplay.Enabled = False
4389       frmPLCData.txtSuctionDisplay.Enabled = False
4390       frmPLCData.txtDischargeDisplay.Enabled = False
4391       frmPLCData.txtTemperatureDisplay.Enabled = False
4392       frmPLCData.txtInHgDisplay.Enabled = False
' <VB WATCH>
4393       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4394       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisablePLCFields"

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
Private Sub BlankData()
' <VB WATCH>
4395       On Error GoTo vbwErrHandler
4396       Const VBWPROCNAME = "frmPLCData.BlankData"
4397       If vbwProtector.vbwTraceProc Then
4398           Dim vbwProtectorParameterString As String
4399           If vbwProtector.vbwTraceParameters Then
4400               vbwProtectorParameterString = "()"
4401           End If
4402           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4403       End If
' </VB WATCH>
4404       txtShpNo.Text = vbNullString
4405       txtBilNo.Text = vbNullString
4406       txtModelNo.Text = vbNullString
4407       cmbMotor.ListIndex = -1
4408       cmbStatorFill.ListIndex = -1
4409       cmbVoltage.ListIndex = -1
4410       cmbDesignPressure.ListIndex = -1
4411       cmbFrequency.ListIndex = -1
4412       cmbCirculationPath.ListIndex = -1
4413       cmbRPM.ListIndex = -1
4414       cmbModel.ListIndex = -1
4415       cmbModelGroup.ListIndex = -1
4416       txtSpGr.Text = vbNullString
4417       txtImpellerDia.Text = vbNullString
4418       txtEndPlay.Text = vbNullString
4419       txtGGap.Text = vbNullString
4420       txtDesignFlow.Text = vbNullString
4421       txtDesignTDH.Text = vbNullString
4422       txtOtherMods.Text = vbNullString
4423       txtRemarks.Text = vbNullString
4424       txtSalesOrderNumber.Text = vbNullString
4425       txtTestSetupRemarks.Text = vbNullString
4426       txtNPSHFile.Text = vbNullString
4427       txtPicturesFile.Text = vbNullString
4428       txtVibrationFile.Text = vbNullString
       '    cmbOrificeNumber.ListIndex = 18
       '    cmbTestSpec.ListIndex = 6       'default = Rev7
4429       cmbLoopNumber.ListIndex = -1
4430       cmbSuctDia.ListIndex = -1
4431       cmbDischDia.ListIndex = -1
4432       cmbTachID.ListIndex = -1
4433       cmbAnalyzerNo.ListIndex = -1
4434       txtTestRemarks.Text = vbNullString
4435       txtHDCor.Text = 0
4436       txtDischHeight.Text = 0
4437       txtSuctHeight.Text = 0
4438       txtKWMult.Text = 1
4439       txtWho.Text = LogInInitials
4440       txtRMA.Text = vbNullString
4441       frmPLCData.chkNPSH.value = 0
4442       frmPLCData.chkPictures.value = 0
4443       frmPLCData.chkVibration.value = 0
4444       frmPLCData.txtFlowmeterID = vbNullString
4445       frmPLCData.txtSuctionID = vbNullString
4446       frmPLCData.txtDischargeID = vbNullString
4447       frmPLCData.txtTemperatureID = vbNullString
4448       frmPLCData.txtMagflowID = vbNullString
4449       frmPLCData.chkBalanceHoles.value = 0
4450       frmPLCData.chkCircOrifice = 0
4451       frmPLCData.txtCircOrifice = vbNullString
4452       frmPLCData.txtImpTrim = vbNullString
4453       frmPLCData.txtOrifice = vbNullString
4454       frmPLCData.chkFeathered.value = 0
4455       frmPLCData.chkTrimmed.value = 0
4456       frmPLCData.chkCircOrifice.value = 0
4457       frmPLCData.txtThrustBal = vbNullString
4458       frmPLCData.txtRPM = vbNullString
4459       frmPLCData.txtVibAx = vbNullString
4460       frmPLCData.txtVibRad = vbNullString
4461       frmPLCData.txtTEMCTRGReading = vbNullString
4462       dgBalanceHoles.Visible = False
' <VB WATCH>
4463       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4464       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BlankData"

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
Private Sub AddTestData()
' <VB WATCH>
4465       On Error GoTo vbwErrHandler
4466       Const VBWPROCNAME = "frmPLCData.AddTestData"
4467       If vbwProtector.vbwTraceProc Then
4468           Dim vbwProtectorParameterString As String
4469           If vbwProtector.vbwTraceParameters Then
4470               vbwProtectorParameterString = "()"
4471           End If
4472           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4473       End If
' </VB WATCH>
4474       Dim I As Integer
4475       Dim sFilter As String

4476       ClearEff
4477       rsEff.MoveFirst

4478       For I = 1 To 8
4479           rsTestData.AddNew
4480           rsTestData!SerialNumber = txtSN
4481           rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
4482           rsTestData!testnumber = I
4483           rsTestData!DataWritten = False
4484           rsTestData.Update
4485           DoEfficiencyCalcs
4486           rsEff.MoveNext
4487           rsTestData.MoveNext
4488       Next I
4489       boFoundTestData = True
           'rsTestData.Update
4490       rsTestData.Requery
4491       rsTestData.Resync

          'select the entries from testdata
4492       sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

4493       rsTestData.Filter = sFilter

4494       Set DataGrid1.DataSource = rsTestData

           ' fix the datagrid

4495       Dim c As Column
4496       For Each c In DataGrid1.Columns
4497          Select Case c.DataField
              Case "TestDataID"
4498             c.Visible = False
4499          Case "SerialNumber"
4500             c.Visible = False
4501          Case "Date"
4502             c.Visible = False
4503          Case Else ' Hide all other columns.
4504             c.Visible = True
4505             c.Alignment = dbgRight
4506          End Select
4507       Next c

4508       rsEff.Requery

4509       DataGrid1.Refresh
4510       DataGrid2.Refresh

          ' fix the datagrid
       '   Set DataGrid1.DataSource = rsTestData
       '   Set DataGrid2.DataSource = rsEff



       '    ClearEff
' <VB WATCH>
4511       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4512       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AddTestData"

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
            vbwReportVariable "sFilter", sFilter
            vbwReportVariable "c", c
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub DoEfficiencyCalcs()
' <VB WATCH>
4513       On Error GoTo vbwErrHandler
4514       Const VBWPROCNAME = "frmPLCData.DoEfficiencyCalcs"
4515       If vbwProtector.vbwTraceProc Then
4516           Dim vbwProtectorParameterString As String
4517           If vbwProtector.vbwTraceParameters Then
4518               vbwProtectorParameterString = "()"
4519           End If
4520           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4521       End If
' </VB WATCH>
4522       Dim KW As Single, VI As Single, VITemp As Single
4523       Dim Vave As Single, Iave As Single
4524       Dim I As Integer
4525       Dim j As Integer
4526       Dim HeightDiff As Single

4527       If Not IsNull(rsTestData.Fields("TotalPower")) Then
4528           KW = rsTestData.Fields("TotalPower")
4529       Else
               'if we wrote data with an old version, we will not have written total power
               'if total power = 0 and the three individual powers are not 0, add them

4530           If rsTestData.Fields("PowerA") > 0 Then
4531               If rsTestData.Fields("PowerB") > 0 Then
4532                   If rsTestData.Fields("PowerC") > 0 Then
4533                       KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
4534                   End If
4535               End If
4536           End If
4537      End If

       '
       '    If Not IsNull(rsTestData.Fields("PowerA")) Then
       '        If rsTestData.Fields("PowerA") > 0 Then        'Magtrol 6530 is negative
       '            KW = rsTestData.Fields("PowerA")
       '        End If
       '    End If
       '    If Not IsNull(rsTestData.Fields("PowerB")) Then
       '        If Val(rsTestData.Fields("PowerB")) > 0 Then        'Magtrol 6530 is negative
       '            KW = KW + rsTestData.Fields("PowerB")
       '        End If
       '    End If
       '    If Not IsNull(rsTestData.Fields("PowerC")) Then
       '        If Val(rsTestData.Fields("PowerC")) > 0 Then        'Magtrol 6530 is negative
       '            KW = KW + rsTestData.Fields("PowerC")
       '        End If
       '    End If

4538       I = 0
4539       Vave = 0
4540       Iave = 0
4541       If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
4542           VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
4543           Vave = rsTestData.Fields("VoltageA")
4544           Iave = rsTestData.Fields("CurrentA")
4545           If VI <> 0 Then
4546               I = I + 1
4547           End If
4548       End If
4549       If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
4550           VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
4551           If VITemp <> 0 Then
4552               I = I + 1
4553               VI = VI + VITemp
4554               Vave = Vave + rsTestData.Fields("VoltageB")
4555               Iave = Iave + rsTestData.Fields("CurrentB")
4556           End If
4557       End If
4558       If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
4559           VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
4560           If VITemp <> 0 Then
4561               I = I + 1
4562               VI = VI + VITemp
4563               Vave = Vave + rsTestData.Fields("VoltageC")
4564               Iave = Iave + rsTestData.Fields("CurrentC")
4565           End If
4566       End If
4567       If KW = 0 Then
4568           For j = 1 To rsEff.Fields.Count - 1
4569               rsEff.Fields(j) = 0
4570           Next j
       '        Exit Sub
4571       End If
4572       If VI <> 0 Then
4573           rsEff.Fields("Volts") = Vave / I
4574           rsEff.Fields("Amps") = Iave / I
4575           rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
4576           rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
4577       Else
4578           rsEff.Fields("PowerFactor") = 0
4579       End If

4580       If optMfr(0).value = True Then
4581           If cmbStatorFill.ListIndex = -1 Then
4582               rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

4583           Else
4584               rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
       '            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
4585           End If
4586       Else
4587           rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
4588       End If

4589       Dim sHDCor As Single
4590       Dim sDisc As Single
4591       Dim sSuct As Single
4592       If IsNull(rsTestSetup.Fields("HDCor")) Then
4593           sHDCor = 0
4594       Else
4595           sHDCor = rsTestSetup.Fields("HDCor")
4596       End If
4597       If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
4598           sDisc = 0
4599       Else
4600           sDisc = rsTestSetup.Fields("DischargeGageHeight")
4601       End If
4602       If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
4603           sSuct = 0
4604       Else
4605           sSuct = rsTestSetup.Fields("SuctionGageHeight")
4606       End If
4607       HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
4608       If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
4609           rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
4610       End If
       '    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
4611       rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
4612       rsEff.Fields("ElecHP") = 1000 * KW / 746
       '    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
4613           If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
4614               If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
           '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4615               rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
           '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4616               rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
4617               If rsEff.Fields("MotorEfficiency") <> 0 Then
4618                   rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
4619               Else
4620                   rsEff.Fields("HydraulicEfficiency") = 0
4621               End If
4622           Else
4623               rsEff.Fields("LiquidHP") = 0
4624               rsEff.Fields("OverallEfficiency") = 0
4625           End If
4626       Else
4627           rsEff.Fields("LiquidHP") = 0
4628           rsEff.Fields("OverallEfficiency") = 0
4629       End If

4630       I = rsEff.AbsolutePosition
4631       If Not IsNull(rsTestData.Fields("Flow")) Then
4632           rsEff.Fields("Flow") = rsTestData.Fields("Flow")
4633           HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
4634           HeadFlow(1, I - 1) = rsEff.Fields("TDH")
4635           FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
4636           FlowHead(I - 1, 1) = rsEff.Fields("TDH")

4637           EffFlow(0, I - 1) = rsTestData.Fields("Flow")
4638           EffFlow(1, I - 1) = rsEff.Fields("OverallEfficiency")
4639           KWFlow(0, I - 1) = rsTestData.Fields("Flow")
4640           KWFlow(1, I - 1) = KW
4641           AmpsFlow(0, I - 1) = rsTestData.Fields("Flow")
4642           AmpsFlow(1, I - 1) = rsEff.Fields("Amps")
4643       Else
4644           HeadFlow(0, I - 1) = 0
4645           HeadFlow(1, I - 1) = 0

4646           EffFlow(0, I - 1) = 0
4647           EffFlow(1, I - 1) = 0
4648           KWFlow(0, I - 1) = 0
4649           KWFlow(1, I - 1) = 0
4650           AmpsFlow(0, I - 1) = 0
4651           AmpsFlow(1, I - 1) = 0
4652       End If

4653       Dim Plothead(7, 1) As Single
4654       Dim HeadPlot(7, 1) As Single
       '    Dim PlotEff() As Single
       '    Dim PlotKW() As Single
       '    Dim PlotAmps() As Single
       '    ReDim PlotHead(0, 0)
       '    ReDim PlotEff(0, 0)
       '    ReDim PlotKW(0, 0)
       '
4655       For j = 0 To UpDown2.value - 1
       '        If HeadFlow(1, j) <> 0 Then
                   'ReDim Preserve Plothead(1, j)
                   'ReDim Preserve HeadPlot(j, 1)
       '            Plothead(0, j) = HeadFlow(0, j)
       '            Plothead(1, j) = HeadFlow(1, j)
4656               HeadPlot(j, 0) = FlowHead(j, 0)
4657               HeadPlot(j, 1) = FlowHead(j, 1)
4658               Plothead(j, 0) = FlowHead(j, 0)
4659               Plothead(j, 1) = FlowHead(j, 1)

       '            ReDim Preserve PlotEff(1, j)
       '            PlotEff(0, j) = EffFlow(0, j)
       '            PlotEff(1, j) = EffFlow(1, j)
       '            ReDim Preserve PlotKW(1, j)
       '            PlotKW(0, j) = KWFlow(0, j)
       '            PlotKW(1, j) = KWFlow(1, j)
       '            ReDim Preserve PlotAmps(1, j)
       '            PlotAmps(0, j) = AmpsFlow(0, j)
       '            PlotAmps(1, j) = AmpsFlow(1, j)
       '        End If
4660       Next j

4661       MSChart1 = Plothead
4662       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
4663       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(Plothead) / 10) + 0.5) + 1)
4664       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
4665       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5



       '        CWGraph2.PlotXY HeadFlow
       '        CWGraphTDH.PlotXY HeadFlow
       '        CWGraphEff.PlotXY EffFlow
       '        CWGraphKw.PlotXY KWFlow

       '    CWGraphTDH.Axes(2).Maximum = SetGraphMax(Plothead())
       '    CWGraphAmps.Axes(2).Maximum = SetGraphMax(PlotAmps())
       '    CWGraph2.Axes(2).Maximum = SetGraphMax(Plothead())

       '    SetGraphMax (Plothead())
       '    If UBound(PlotHead()) <> 0 Then
       '        CWGraph2.PlotXY Plothead
       '        CWGraphTDH.PlotXY Plothead
       '        CWGraphEff.PlotXY PlotEff
       '        CWGraphKw.PlotXY PlotKW
       '        CWGraphAmps.PlotXY PlotAmps
       '    End If

           'copy fields for reports
4666       rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
4667       rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
       '    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
       '    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
4668       rsEff.Fields("KW") = KW
4669       rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
4670       rsEff.Fields("RPM") = rsTestData.Fields("RPM")
4671       rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
4672       rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
4673       rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
4674       rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
4675       rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
4676       rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
4677       rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
4678       rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
4679       rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
4680       rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
4681       rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
4682       rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
4683       rsEff.Fields("TC1") = rsTestData.Fields("TC1")
4684       rsEff.Fields("TC2") = rsTestData.Fields("TC2")
4685       rsEff.Fields("TC3") = rsTestData.Fields("TC3")
4686       rsEff.Fields("TC4") = rsTestData.Fields("TC4")
4687       rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
4688       rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
4689       rsEff.Fields("AI4") = rsTestData.Fields("AI4")
4690       rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
4691       rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4692       rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
4693       rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
4694       rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
4695       rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
4696       rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
4697       If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
4698           txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
4699       End If
4700       If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
4701           txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
4702       End If
4703       If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
4704           txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
4705       End If
4706       If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
4707           txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
4708       End If
4709       If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
4710           txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
4711       End If

4712       CalculateTEMCForce

4713       If Not IsNull(txtTEMCCalcForce.Text) Then
4714           rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
4715       Else
4716           rsEff.Fields("TEMCCalculatedForce") = 0
4717       End If

4718       If Not IsNull(txtTEMCPVValue.Text) Then
4719           rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
4720       Else
4721           rsEff.Fields("TEMCPV") = 0
4722       End If

4723       If Val(txtTEMCFrontThrust.Text) <> 0 Then
4724           rsEff.Fields("TEMCFR") = "F"
       '        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4725       Else
4726           If Val(txtTEMCRearThrust.Text) = 0 Then
                   'no thrust
4727               rsEff.Fields("TEMCFR") = " "
4728               rsEff.Fields("TEMCFrontThrust") = 0
4729           Else
4730               rsEff.Fields("TEMCFR") = "R"
       '            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
4731           End If
4732       End If

4733       rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

4734       rsEff.Update

4735       If rsEffDisp.State = adStateOpen Then
4736           rsEffDisp.Close
4737       End If

4738       Dim qyEffDisp As New ADODB.Command
4739       qyEffDisp.ActiveConnection = cnEffData
4740       qyEffDisp.CommandText = "SELECT Flow, TDH, KW, Volts, Amps, OverallEfficiency FROM Efficiency;"

4741       With rsEffDisp     'open the recordset for the query
4742           .CursorLocation = adUseClient
4743           .CursorType = adOpenStatic
4744           .LockType = adLockOptimistic
4745           .Open qyEffDisp
4746       End With

4747       rsEffDisp.Requery
4748       DataGrid2.Refresh


' <VB WATCH>
4749       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4750       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DoEfficiencyCalcs"

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
            vbwReportVariable "KW", KW
            vbwReportVariable "VI", VI
            vbwReportVariable "VITemp", VITemp
            vbwReportVariable "Vave", Vave
            vbwReportVariable "Iave", Iave
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "HeightDiff", HeightDiff
            vbwReportVariable "sHDCor", sHDCor
            vbwReportVariable "sDisc", sDisc
            vbwReportVariable "sSuct", sSuct
            vbwReportVariable "Plothead", Plothead
            vbwReportVariable "HeadPlot", HeadPlot
            vbwReportVariable "qyEffDisp", qyEffDisp
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub ClearEff()
       '    Dim I As Integer, j As Integer
' <VB WATCH>
4751       On Error GoTo vbwErrHandler
4752       Const VBWPROCNAME = "frmPLCData.ClearEff"
4753       If vbwProtector.vbwTraceProc Then
4754           Dim vbwProtectorParameterString As String
4755           If vbwProtector.vbwTraceParameters Then
4756               vbwProtectorParameterString = "()"
4757           End If
4758           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4759       End If
' </VB WATCH>
4760       Dim qy As New ADODB.Command

4761       If rsEff.State = adStateOpen Then
4762           If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
4763               rsEff.CancelUpdate
4764           End If
4765           rsEff.Close
4766       End If
4767       qy.ActiveConnection = cnEffData
4768       qy.CommandText = "DROP TABLE Efficiency"
4769       rsEff.Open qy
4770       qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
4771       rsEff.Open qy
4772       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

4773       rsEff.Requery
4774       DataGrid2.Refresh

4775       Dim c As Column
4776       For Each c In DataGrid2.Columns
4777           c.Alignment = dbgCenter
4778           c.Width = 750
4779           Select Case c.ColIndex
                   Case 1
4780                   c.Caption = "Flow"
4781                   c.NumberFormat = "###0.00"
4782               Case 2
4783                   c.Caption = "TDH"
4784                   c.NumberFormat = "00.0"
4785               Case 3
4786                   c.Caption = "Overall Eff"
4787                   c.NumberFormat = "00.00"
4788                   c.Width = 850
4789               Case 4
4790                   c.Caption = "PF"
4791                   c.NumberFormat = "00.0"
4792               Case 5
4793                   c.Caption = "Vel Head"
4794                   c.NumberFormat = "00.00"
4795               Case 6
4796                   c.Caption = "Elec HP"
4797                   c.NumberFormat = "#00.0"
4798               Case 7
4799                   c.Caption = "Liq HP"
4800                   c.NumberFormat = "#00.0"
4801               Case Else
4802                   c.Visible = False
4803           End Select
4804       Next c

' <VB WATCH>
4805       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4806       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ClearEff"

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
            vbwReportVariable "qy", qy
            vbwReportVariable "c", c
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Function JustAlphaNumeric(char As String) As String
' <VB WATCH>
4807       On Error GoTo vbwErrHandler
4808       Const VBWPROCNAME = "frmPLCData.JustAlphaNumeric"
4809       If vbwProtector.vbwTraceProc Then
4810           Dim vbwProtectorParameterString As String
4811           If vbwProtector.vbwTraceParameters Then
4812               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("char", char) & ") "
4813           End If
4814           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4815       End If
' </VB WATCH>
4816       Select Case Asc(char)
               Case 42             ' *
4817               JustAlphaNumeric = char
4818           Case 48 To 57       ' 0 - 9
4819               JustAlphaNumeric = char
4820           Case 65 To 90       ' A - Z
4821               JustAlphaNumeric = char
4822           Case 97 To 122      ' a - z
4823               JustAlphaNumeric = UCase(char)
4824           Case Else
4825               JustAlphaNumeric = ""
4826       End Select
' <VB WATCH>
4827       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4828       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "JustAlphaNumeric"

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
            vbwReportVariable "char", char
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function



Private Sub txtI1_Change()
' <VB WATCH>
4829       On Error GoTo vbwErrHandler
4830       Const VBWPROCNAME = "frmPLCData.txtI1_Change"
4831       If vbwProtector.vbwTraceProc Then
4832           Dim vbwProtectorParameterString As String
4833           If vbwProtector.vbwTraceParameters Then
4834               vbwProtectorParameterString = "()"
4835           End If
4836           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4837       End If
' </VB WATCH>
4838       txtI2.Text = txtI1.Text
4839       txtI3.Text = txtI1.Text
' <VB WATCH>
4840       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4841       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtI1_Change"

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

Private Sub txtModelNo_Change()
' <VB WATCH>
4842       On Error GoTo vbwErrHandler
4843       Const VBWPROCNAME = "frmPLCData.txtModelNo_Change"
4844       If vbwProtector.vbwTraceProc Then
4845           Dim vbwProtectorParameterString As String
4846           If vbwProtector.vbwTraceParameters Then
4847               vbwProtectorParameterString = "()"
4848           End If
4849           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4850       End If
' </VB WATCH>
4851       Dim I As Integer
4852       Dim S As String
4853       Dim sFull As String
4854       Dim boDone As Boolean
4855       Dim boRepeat As Boolean

4856       Static bo3Digits As Boolean         '3 digits in frame number
4857       Static bo2Digits As Boolean         '2 digits in stages

4858       If optMfr(0).value = True Then
' <VB WATCH>
4859       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4860           Exit Sub
4861       End If

4862       cmbTEMCAdapter.ListIndex = -1
4863       cmbTEMCAdditions.ListIndex = -1
4864       cmbTEMCCirculation.ListIndex = -1
4865       cmbTEMCDesignPressure.ListIndex = -1
4866       cmbTEMCNominalDischargeSize.ListIndex = -1
4867       cmbTEMCDivisionType.ListIndex = -1
4868       cmbTEMCImpellerType.ListIndex = -1
4869       cmbTEMCInsulation.ListIndex = -1
4870       cmbTEMCJacketGasket.ListIndex = -1
4871       cmbTEMCMaterials.ListIndex = -1
4872       cmbTEMCModel.ListIndex = -1
4873       cmbTEMCNominalImpSize.ListIndex = -1
4874       cmbTEMCOtherMotor.ListIndex = -1
4875       cmbTEMCPumpStages.ListIndex = -1
4876       cmbTEMCNominalSuctionSize.ListIndex = -1
4877       cmbTEMCTRG.ListIndex = -1
4878       cmbTEMCVoltage.ListIndex = -1


           'first, get rid of spaces, dashes, etc

4879       S = ""
4880       For I = 1 To Len(txtModelNo.Text)
4881           S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
4882       Next I

           'next, fill out the model number to it's max length of 24 characters

4883       boDone = False
4884       boRepeat = False

4885       Do While Not boDone
4886           sFull = ""
4887           For I = 1 To Len(S)
4888               Select Case I
                       Case 1
                           'type
4889                       sFull = sFull & Mid$(S, I, 1)
4890                   Case 2
                           'adapter
4891                       If IsNumeric(Mid$(S, I, 1)) Then
4892                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4893                           boRepeat = True
4894                           Exit For
4895                       Else
4896                           sFull = sFull & Mid$(S, I, 1)
4897                           boRepeat = False
4898                       End If
4899                   Case 3
                           'materials
4900                       sFull = sFull & Mid$(S, I, 1)
4901                   Case 4
                       'design pressure
4902                       sFull = sFull & Mid$(S, I, 1)
4903                   Case 5
                       'motor frame number - digit 1
4904                       sFull = sFull & Mid$(S, I, 1)
4905                   Case 6
                       'motor frame number - digit 2
4906                       sFull = sFull & Mid$(S, I, 1)
4907                   Case 7
                       'motor frame number - digit 3
4908                       sFull = sFull & Mid$(S, I, 1)
4909                   Case 8
                       'motor frame number - digit 4
4910                       If IsNumeric(Mid$(S, I, 1)) Then
4911                           sFull = sFull & Mid$(S, I, 1)
4912                           boRepeat = False
4913                       Else    '3 digits
       '                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
4914                           S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
4915                           boRepeat = True
4916                           Exit For
4917                       End If
4918                   Case 9
                       'insulation
4919                       sFull = sFull & Mid$(S, I, 1)
4920                   Case 10
                       'voltage
4921                       sFull = sFull & Mid$(S, I, 1)
4922                   Case 11
                       'other motor specs
4923                       If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
4924                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4925                           boRepeat = True
4926                           Exit For
4927                       Else
4928                           sFull = sFull & Mid$(S, I, 1)
4929                           boRepeat = False
4930                       End If
4931                   Case 12
                       ' TRG
4932                       sFull = sFull & Mid$(S, I, 1)
4933                   Case 13
                       'Nominal discharge - digit 1
4934                       sFull = sFull & Mid$(S, I, 1)
4935                   Case 14
                       'nominal discharge - digit 2
4936                       sFull = sFull & Mid$(S, I, 1)
4937                   Case 15
                       'nominal suction - digit 1
4938                       sFull = sFull & Mid$(S, I, 1)
4939                   Case 16
                       'nominal suction - digit 2
4940                       sFull = sFull & Mid$(S, I, 1)
4941                   Case 17
                       'nominal impeller size
4942                       sFull = sFull & Mid$(S, I, 1)
4943                   Case 18
                       'impeller type
4944                       If Val(Mid$(sFull, 5, 1)) < 3 Then
4945                           If IsNumeric(Mid$(S, I, 1)) Or Mid$(S, I, 1) = "*" Then
4946                               sFull = sFull & Mid$(S, I, 1)
4947                               boRepeat = False
4948                           Else
4949                               S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4950                               boRepeat = True
4951                               Exit For
4952                           End If
4953                       Else
4954                           If Mid$(S, I, 1) = "*" Then
4955                               boRepeat = False
4956                           Else
4957                               S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4958                               boRepeat = True
4959                               Exit For
4960                           End If
4961                       End If
4962                   Case 19
                       'Division type
4963                       If IsNumeric(Mid$(S, I, 1)) Then
4964                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4965                           boRepeat = True
4966                           Exit For
4967                       Else
4968                           sFull = sFull & Mid$(S, I, 1)
4969                           boRepeat = False
4970                       End If
4971                   Case 20
                       'pump stages - digit 1
4972                       sFull = sFull & Mid$(S, I, 1)
4973                   Case 21
                       'pump stages - digit 2
4974                       If IsNumeric(Mid$(S, I, 1)) Or Mid$(S, I, 1) = "*" Then
4975                           sFull = sFull & Mid$(S, I, 1)
4976                           boRepeat = False
4977                       Else
4978                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4979                           boRepeat = True
4980                           Exit For
4981                       End If
4982                   Case 22
                       'pump jacket
4983                       If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                               Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
4984                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4985                           boRepeat = True
4986                       Else
4987                           sFull = sFull & Mid$(S, I, 1)
4988                           boRepeat = False
4989                       End If
4990                   Case 23
                       'additions
4991                         sFull = sFull & Mid$(S, I, 1)
4992                   Case 24
                       'circulation
4993                         sFull = sFull & Mid$(S, I, 1)
4994               End Select
4995           Next I
4996           If Not boRepeat Then
4997               boDone = True
4998           End If
4999       Loop

5000       sFull = S
5001       For I = 1 To Len(sFull)
5002           Select Case I
                   Case 1
5003                   ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
5004               Case 2
5005                   ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
5006               Case 3
5007                   ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
5008               Case 4
5009                   ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
5010               Case 5
       '                If IsNumeric(Mid$(sFull, i, 1)) Then  '4 digit frame number
5011                       If Val(Mid$(sFull, I, 1)) = 0 Then
5012                           txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
5013                       Else
5014                           txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
5015                       End If
       '                    bo3Digits = False
       '                Else
       '                    txtTEMCFrameNumber.Text = Mid$(sFull, 5, 3)
       '                    ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, i, 1)
       '                    bo3Digits = True
       '                End If
5016               Case 9
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, i, 1)
       '                Else
5017                       ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
       '                End If
5018               Case 10
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, i, 1)
       '                Else
5019                       ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
       '                End If
5020               Case 11
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, i, 1)
       '                Else
5021                       ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
       '                End If
5022               Case 12
       '                If bo3Digits Then
       '                Else
5023                       ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
       '                End If
5024               Case 13
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCNominalDischargeSize, Right$(sFull, 2)
       '                Else
       '                End If
5025                       ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
5026               Case 14
       '                If bo3Digits Then
       '                Else
       '                End If
5027               Case 15
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCNominalSuctionSize, Right$(sFull, 2)
       '                Else
       '                End If
5028                       ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
5029               Case 16
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, i, 1)
       '                Else
       '                End If
5030               Case 17
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, i, 1)
       '                Else
5031                       ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
       '                End If
5032               Case 18
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, i, 1)
       '                Else
5033                       ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
       '                End If
5034               Case 19
       '                If bo3Digits Then
       '                    ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, i, 1)
       '                Else
5035                       ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
       '                End If
5036               Case 20
       '                If bo3Digits Then
       '                    If IsNumeric(Mid$(sFull, i, 1)) Then  '2 digit stages
       '                        ParseTEMCModelNo cmbTEMCPumpStages, Right$(sFull, 2)
       '                        bo2Digits = True
       '                    Else
       '                        ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, i, 1)
       '                        bo2Digits = False
       '                    End If
       '                Else
5037                       If IsNumeric(Mid$(sFull, I + 1, 1)) Then
5038                           ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 2)
5039                       Else
5040                           ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
5041                       End If
       '                End If
5042               Case 21
       '                If bo3Digits Then
       '                    If bo2Digits Then
       '                        ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, i, 1)
       '                    Else
       '                        ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, i, 1)
       '                    End If
       '                Else
       '                    If IsNumeric(Mid$(sFull, i, 1)) Then  '2 digit stages
       '                        ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, i, 2)
       '                        bo2Digits = True
       '                    Else
       '                        ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, i, 1)
       '                        bo2Digits = False
       '                    End If
       '                End If
5043               Case 22
       '                If bo3Digits Then
       '                    If bo2Digits Then
       '                        ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, i, 1)
       '                    Else
       '                        ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, i, 1)
       '                    End If
       '                Else
       '                    If bo2Digits Then
5044                           ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
       '                    Else
       '                        ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, i, 1)
       '                    End If
       '                End If
5045               Case 23
       '                If bo3Digits Then
       '                    If bo2Digits Then
       '                        ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, i, 1)
       '                    Else
       '                    End If
       '                Else
       '                    If bo2Digits Then
5046                           ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
       '                    Else
       '                        ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, i, 1)
       '                    End If
       '                End If
5047               Case 24
       '                If bo3Digits Then
       '                    If bo2Digits Then
       '                    Else
       '                    End If
       '                Else
       '                    If bo2Digits Then
5048                           ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)
       '                    Else
       '                    End If
       '                End If

5049           End Select
5050       Next I
' <VB WATCH>
5051       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5052       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNo_Change"

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
            vbwReportVariable "S", S
            vbwReportVariable "sFull", sFull
            vbwReportVariable "boDone", boDone
            vbwReportVariable "boRepeat", boRepeat
            vbwReportVariable "bo3Digits", bo3Digits
            vbwReportVariable "bo2Digits", bo2Digits
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub txtModelNo_Validate(Cancel As Boolean)
' <VB WATCH>
5053       On Error GoTo vbwErrHandler
5054       Const VBWPROCNAME = "frmPLCData.txtModelNo_Validate"
5055       If vbwProtector.vbwTraceProc Then
5056           Dim vbwProtectorParameterString As String
5057           If vbwProtector.vbwTraceParameters Then
5058               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
5059           End If
5060           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5061       End If
' </VB WATCH>
5062       Dim I As Integer
5063       Dim S As String

       '    s = txtModelNo.Text
       '    S = Replace(S, "-", "")
       '    S = Replace(S, " ", "")
       '    S = Replace(S, "/", "")

       '    txtModelNo.Text = ""

       '    For i = 1 To Len(s)
       '        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
       '    Next i
5064       txtModelNo_Change

' <VB WATCH>
5065       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5066       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNo_Validate"

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
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub txtNPSHFile_GotFocus()
' <VB WATCH>
5067       On Error GoTo vbwErrHandler
5068       Const VBWPROCNAME = "frmPLCData.txtNPSHFile_GotFocus"
5069       If vbwProtector.vbwTraceProc Then
5070           Dim vbwProtectorParameterString As String
5071           If vbwProtector.vbwTraceParameters Then
5072               vbwProtectorParameterString = "()"
5073           End If
5074           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5075       End If
' </VB WATCH>
5076       On Error GoTo FileCancel
5077       If LenB(txtNPSHFile.Text) <> 0 Then
5078           CommonDialog1.filename = txtNPSHFile.Text
5079       End If
5080       CommonDialog1.ShowOpen
5081       txtNPSHFile.Text = CommonDialog1.filename
' <VB WATCH>
5082       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5083       Exit Sub
5084   FileCancel:
5085   On Error GoTo vbwErrHandler
5086       CommonDialog1.CancelError = False
' <VB WATCH>
5087       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5088       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtNPSHFile_GotFocus"

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

Private Sub txtP1_Change()
' <VB WATCH>
5089       On Error GoTo vbwErrHandler
5090       Const VBWPROCNAME = "frmPLCData.txtP1_Change"
5091       If vbwProtector.vbwTraceProc Then
5092           Dim vbwProtectorParameterString As String
5093           If vbwProtector.vbwTraceParameters Then
5094               vbwProtectorParameterString = "()"
5095           End If
5096           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5097       End If
' </VB WATCH>
5098       txtP2.Text = txtP1.Text
5099       txtP3.Text = txtP1.Text
' <VB WATCH>
5100       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5101       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtP1_Change"

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

Private Sub txtPicturesFile_gotfocus()
' <VB WATCH>
5102       On Error GoTo vbwErrHandler
5103       Const VBWPROCNAME = "frmPLCData.txtPicturesFile_gotfocus"
5104       If vbwProtector.vbwTraceProc Then
5105           Dim vbwProtectorParameterString As String
5106           If vbwProtector.vbwTraceParameters Then
5107               vbwProtectorParameterString = "()"
5108           End If
5109           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5110       End If
' </VB WATCH>
5111       CommonDialog1.CancelError = True
5112       On Error GoTo FileCancel
5113       If LenB(txtPicturesFile.Text) <> 0 Then
5114           CommonDialog1.filename = txtPicturesFile.Text
5115       End If
5116       CommonDialog1.ShowOpen
5117       txtPicturesFile.Text = CommonDialog1.filename
' <VB WATCH>
5118       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5119       Exit Sub
5120   FileCancel:
5121   On Error GoTo vbwErrHandler
5122       CommonDialog1.CancelError = False
' <VB WATCH>
5123       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5124       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtPicturesFile_gotfocus"

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

Private Sub txtSN_Change()
' <VB WATCH>
5125       On Error GoTo vbwErrHandler
5126       Const VBWPROCNAME = "frmPLCData.txtSN_Change"
5127       If vbwProtector.vbwTraceProc Then
5128           Dim vbwProtectorParameterString As String
5129           If vbwProtector.vbwTraceParameters Then
5130               vbwProtectorParameterString = "()"
5131           End If
5132           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5133       End If
' </VB WATCH>
5134       cmdFindPump.Default = True
' <VB WATCH>
5135       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5136       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtSN_Change"

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

Private Sub txtTEMCFrontThrust_Change()
' <VB WATCH>
5137       On Error GoTo vbwErrHandler
5138       Const VBWPROCNAME = "frmPLCData.txtTEMCFrontThrust_Change"
5139       If vbwProtector.vbwTraceProc Then
5140           Dim vbwProtectorParameterString As String
5141           If vbwProtector.vbwTraceParameters Then
5142               vbwProtectorParameterString = "()"
5143           End If
5144           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5145       End If
' </VB WATCH>
5146       CalculateTEMCForce
' <VB WATCH>
5147       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5148       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCFrontThrust_Change"

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

Private Sub txtTEMCMomentArm_Change()
' <VB WATCH>
5149       On Error GoTo vbwErrHandler
5150       Const VBWPROCNAME = "frmPLCData.txtTEMCMomentArm_Change"
5151       If vbwProtector.vbwTraceProc Then
5152           Dim vbwProtectorParameterString As String
5153           If vbwProtector.vbwTraceParameters Then
5154               vbwProtectorParameterString = "()"
5155           End If
5156           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5157       End If
' </VB WATCH>
5158       CalculateTEMCForce
' <VB WATCH>
5159       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5160       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCMomentArm_Change"

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

Private Sub txtTEMCRearThrust_Change()
' <VB WATCH>
5161       On Error GoTo vbwErrHandler
5162       Const VBWPROCNAME = "frmPLCData.txtTEMCRearThrust_Change"
5163       If vbwProtector.vbwTraceProc Then
5164           Dim vbwProtectorParameterString As String
5165           If vbwProtector.vbwTraceParameters Then
5166               vbwProtectorParameterString = "()"
5167           End If
5168           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5169       End If
' </VB WATCH>
5170       CalculateTEMCForce
' <VB WATCH>
5171       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5172       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCRearThrust_Change"

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

Private Sub txtTEMCThrustRigPressure_Change()
' <VB WATCH>
5173       On Error GoTo vbwErrHandler
5174       Const VBWPROCNAME = "frmPLCData.txtTEMCThrustRigPressure_Change"
5175       If vbwProtector.vbwTraceProc Then
5176           Dim vbwProtectorParameterString As String
5177           If vbwProtector.vbwTraceParameters Then
5178               vbwProtectorParameterString = "()"
5179           End If
5180           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5181       End If
' </VB WATCH>
5182       CalculateTEMCForce
' <VB WATCH>
5183       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5184       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCThrustRigPressure_Change"

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

Private Sub txtTEMCViscosity_Change()
' <VB WATCH>
5185       On Error GoTo vbwErrHandler
5186       Const VBWPROCNAME = "frmPLCData.txtTEMCViscosity_Change"
5187       If vbwProtector.vbwTraceProc Then
5188           Dim vbwProtectorParameterString As String
5189           If vbwProtector.vbwTraceParameters Then
5190               vbwProtectorParameterString = "()"
5191           End If
5192           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5193       End If
' </VB WATCH>
5194       CalculateTEMCForce
' <VB WATCH>
5195       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5196       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCViscosity_Change"

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

Private Sub txtV1_Change()
' <VB WATCH>
5197       On Error GoTo vbwErrHandler
5198       Const VBWPROCNAME = "frmPLCData.txtV1_Change"
5199       If vbwProtector.vbwTraceProc Then
5200           Dim vbwProtectorParameterString As String
5201           If vbwProtector.vbwTraceParameters Then
5202               vbwProtectorParameterString = "()"
5203           End If
5204           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5205       End If
' </VB WATCH>
5206       txtV2.Text = txtV1.Text
5207       txtV3.Text = txtV1.Text
' <VB WATCH>
5208       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5209       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtV1_Change"

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

Private Sub txtVibrationFile_gotfocus()
' <VB WATCH>
5210       On Error GoTo vbwErrHandler
5211       Const VBWPROCNAME = "frmPLCData.txtVibrationFile_gotfocus"
5212       If vbwProtector.vbwTraceProc Then
5213           Dim vbwProtectorParameterString As String
5214           If vbwProtector.vbwTraceParameters Then
5215               vbwProtectorParameterString = "()"
5216           End If
5217           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5218       End If
' </VB WATCH>
5219       On Error GoTo FileCancel
5220       If LenB(txtVibrationFile.Text) <> 0 Then
5221           CommonDialog1.filename = txtVibrationFile.Text
5222       End If
5223       CommonDialog1.ShowOpen
5224       txtVibrationFile.Text = CommonDialog1.filename
' <VB WATCH>
5225       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5226       Exit Sub
5227   FileCancel:
5228   On Error GoTo vbwErrHandler
5229       CommonDialog1.CancelError = False
' <VB WATCH>
5230       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5231       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtVibrationFile_gotfocus"

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

Private Sub ExportToExcel()
' <VB WATCH>
5232       On Error GoTo vbwErrHandler
5233       Const VBWPROCNAME = "frmPLCData.ExportToExcel"
5234       If vbwProtector.vbwTraceProc Then
5235           Dim vbwProtectorParameterString As String
5236           If vbwProtector.vbwTraceParameters Then
5237               vbwProtectorParameterString = "()"
5238           End If
5239           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5240       End If
' </VB WATCH>

5241       Dim SaveFileName As String
5242       Dim WorkSheetName As String

5243       Dim I As Integer
5244       Dim iRowNo As Integer
5245       Dim sImp As String
5246       Dim ans As Integer

5247       Dim bCanShowSpeed As Boolean
5248       Dim CantShowReason As String

       'close any running excel processes
5249       Dim objWMIService, colProcesses
5250       Set objWMIService = GetObject("winmgmts:")
5251       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
5252       If colProcesses.Count > 0 Then
5253           Set xlApp = Excel.Application
5254       Else
               'use existing copy
       '        Set xlApp = New Excel.Application
5255           Set xlApp = CreateObject("Excel.Application")
5256       End If


5257       CommonDialog1.CancelError = True        'in case the user
5258       On Error GoTo ErrHandler                '  chooses the cancel button

           'set up dialog box
5259       CommonDialog1.DialogTitle = "Open Excel Files"
5260       CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
5261       CommonDialog1.InitDir = App.Path
       '    CommonDialog1.InitDir = "C:\"    'in this directory
5262       CommonDialog1.ShowOpen                              'open the file selection dialog box

5263       If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
5264           SaveFileName = CommonDialog1.filename           'get the name of the file
5265           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
5266                xlApp.Workbooks.Close
5267           End If
               ' Create the Excel Workbook Object.
5268   On Error GoTo vbwErrHandler
5269           Set xlBook = xlApp.Workbooks.Add                'add a workbook
5270           WorkSheetName = NewWorkBook                     'do some stuff for the new workbook
5271           ActiveWorkbook.CheckCompatibility = False
5272           xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                                 FileFormat:=xlNormal                        'save the file
5273       Else                                                'the file name already exists
5274           SaveFileName = CommonDialog1.filename
               ' Create the Excel Workbook Object.
5275           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
5276                xlApp.Workbooks.Close
5277           End If
5278           Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
5279           If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
5280               MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
' <VB WATCH>
5281       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5282               Exit Sub
5283           Else
5284           End If
5285       End If

5286   On Error GoTo vbwErrHandler

           'see if we can export Speed and SG and if we can, ask user if s/he wants it
           'assume that we can show speed calcs

5287       bCanShowSpeed = False
       'open the template and copy the data from the sheet
       '  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
           'write the data to the spreadsheet
5288       With xlApp

5289       Dim xlTemplateName As String
5290       xlTemplateName = ParentDirectoryName & "\PumpData Excel Template.xls"
5291       Dim xlTemplate As Excel.Workbook
5292       Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
5293       Dim TemplateWS As Excel.Worksheet
5294       Dim sheetName As String
5295       sheetName = xlTemplate.Sheets(1).Name
5296       xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

5297       xlTemplate.Close savechanges:=False

5298       Set xlTemplate = Nothing

5299       Application.DisplayAlerts = False
5300       ActiveWorkbook.Worksheets(WorkSheetName).Delete
5301       Application.DisplayAlerts = True
5302       ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

           'WorkSheetName = sheetName

           'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
           ' if there is, get the coefficients, else make the coefficients 0

5303           Dim ACoef As Double
5304           Dim BCoef As Double
5305           Dim CCoef As Double

5306           Dim qy As New ADODB.Command
5307           Dim rs As New ADODB.Recordset
5308           qy.ActiveConnection = cnPumpData
       '        Dim VoltageForLookup As Integer
       '        If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
       '            VoltageForLookup = 460
       '        ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
       '            VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
       '        End If

5309           If rsPumpData!ChempumpPump = True Then
5310               qy.CommandText = "SELECT * FROM CalculatedRPM WHERE Frame = '" & cmbMotor.List(cmbMotor.ListIndex) & "'"
5311           Else
5312               qy.CommandText = "SELECT * FROM CalculatedRPM WHERE Frame = '" & txtTEMCFrameNumber.Text & "'"
5313           End If

5314           rs.CursorLocation = adUseClient
5315           rs.CursorType = adOpenStatic

5316           rs.Open qy
5317           If rs.RecordCount = 0 Then
5318               ACoef = 0
5319               BCoef = 0
5320               CCoef = 0
5321               MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                          " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                          " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
5322           Else
5323               ACoef = rs.Fields("x2")
5324               BCoef = rs.Fields("x")
5325               CCoef = rs.Fields("b")
5326               .Range("H8").Select
5327               .ActiveCell.FormulaR1C1 = rs.Fields("Poles")
5328               .Range("H54").Select
5329               .ActiveCell.FormulaR1C1 = rs.Fields("Rotor OD")
5330               .Range("H55").Select
5331               .ActiveCell.FormulaR1C1 = rs.Fields("Rotor Length")
5332           End If


           'write header data
               'first write the revision
5333           Dim RundownRev As String
5334           RundownRev = App.Major & "." & App.Minor & "." & App.Revision

5335           .Range("AM3").Select
5336           .ActiveCell.FormulaR1C1 = RundownRev

5337           .Range("A2").Select
5338           .ActiveCell.FormulaR1C1 = "Serial Number"
5339           .Range("C2").Select
5340           .ActiveCell.FormulaR1C1 = txtSN

5341           .Range("F1").Select
5342           .ActiveCell.FormulaR1C1 = "Customer"
5343           .Range("H1").Select
5344           .ActiveCell.FormulaR1C1 = txtShpNo

5345           .Range("A3").Select
5346           .ActiveCell.FormulaR1C1 = "Model"
5347           .Range("C3").Select
5348           .ActiveCell.FormulaR1C1 = txtModelNo

5349           .Range("F2").Select
5350           .ActiveCell.FormulaR1C1 = "Sales Order"
5351           .Range("H2").Select
5352           .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

5353           .Range("A9").Select
5354           .ActiveCell.FormulaR1C1 = "Design Flow"
5355           .Range("C9").Select
5356           .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

5357           .Range("A10").Select
5358           .ActiveCell.FormulaR1C1 = "Design Head"
5359           .Range("C10").Select
5360           .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

5361           .Range("P13").Select
5362           .ActiveCell.FormulaR1C1 = "Barometric Pressure"
5363           .Range("R13").Select
5364           .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

5365           .Range("P11").Select
5366           .ActiveCell.FormulaR1C1 = "Suction Gage Height"
5367           .Range("R11").Select
5368           .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

5369           .Range("P12").Select
5370           .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
5371           .Range("R12").Select
5372           .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

5373           .Range("A1").Select
5374           .ActiveCell.FormulaR1C1 = "Run Date"
5375           .Range("C1").Select
5376           .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

5377           .Range("D10:E10").Select
5378           With xlApp.Selection
5379               .HorizontalAlignment = xlCenter
5380               .VerticalAlignment = xlBottom
5381               .WrapText = False
5382               .Orientation = 0
5383               .AddIndent = False
5384               .IndentLevel = 0
5385               .ShrinkToFit = False
5386               .ReadingOrder = xlContext
5387               .MergeCells = False
5388           End With
5389           xlApp.Selection.Merge

               'determine rpm

5390           Dim RPMvalue As String
5391           If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
               '1 says 2 pole
5392               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
5393                   RPMvalue = "2900"
5394               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
5395                   RPMvalue = "3450"
5396               Else
                       'vfd or other, no rpm
5397                   RPMvalue = ""
5398               End If
5399           Else
               '2 says 4 pole
5400               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
5401                   RPMvalue = "1450"
5402               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
5403                   RPMvalue = "1750"
5404               Else
                       'vfd or other, no rpm
5405                   RPMvalue = ""
5406               End If
5407           End If

       '        .Range("G1").Select
       '        .ActiveCell.FormulaR1C1 = "RPM"
       '        .Range("I1").Select
       '        .ActiveCell.FormulaR1C1 = RPMvalue

5408           .Range("A5").Select
5409           .ActiveCell.FormulaR1C1 = "Sp Gravity"
5410           .Range("C5").Select
5411           .ActiveCell.FormulaR1C1 = txtSpGr

5412           .Range("A6").Select
5413           .ActiveCell.FormulaR1C1 = "Viscosity"
5414           .Range("C6").Select
5415           .ActiveCell.FormulaR1C1 = txtViscosity

5416           .Range("F4").Select
5417           .ActiveCell.FormulaR1C1 = "Motor"
5418           .Range("H4").Select
5419           If rsPumpData!ChempumpPump = True Then
5420               .ActiveCell.FormulaR1C1 = Me.cmbMotor.List(Me.cmbMotor.ListIndex)
5421           Else
5422               .ActiveCell.FormulaR1C1 = Me.txtTEMCFrameNumber.Text
5423           End If

5424           .Range("H12").Select
       '        .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text

5425           .Range("F5").Select
5426           .ActiveCell.FormulaR1C1 = "Voltage"
5427           .Range("H5").Select
5428           .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

5429           .Range("K6").Select
5430           .ActiveCell.FormulaR1C1 = "End Play"
5431           .Range("M6").Select
5432           .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

5433           .Range("K7").Select
5434           .ActiveCell.FormulaR1C1 = "G-Gap"
5435           .Range("M7").Select
5436           .ActiveCell.FormulaR1C1 = txtGGap.Text

5437           .Range("A8").Select
5438           .ActiveCell.FormulaR1C1 = "Design Pressure"
5439           .Range("C8").Select

5440           If rsPumpData!ChempumpPump = False Then
5441               Dim DesPress As String
5442               DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
5443               Dim j As Integer
5444               j = InStrRev(DesPress, "-")
5445               .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)
5446           Else
5447               .ActiveCell.FormulaR1C1 = Me.cmbDesignPressure.List(Me.cmbDesignPressure.ListIndex)
5448           End If

       '        .Range("G8").Select
       '        .ActiveCell.FormulaR1C1 = "Stator Fill"
       '        .Range("I8").Select
       '        .ActiveCell.FormulaR1C1 = "Dry"

5449           .Range("K4").Select
5450           .ActiveCell.FormulaR1C1 = "Circulation Path"

5451           .Range("M4").Select
5452           If rsPumpData!ChempumpPump = False Then
5453               .ActiveCell.FormulaR1C1 = Me.cmbTEMCModel.List(Me.cmbTEMCModel.ListIndex)
5454           Else
5455               .ActiveCell.FormulaR1C1 = Me.cmbCirculationPath.List(Me.cmbCirculationPath.ListIndex)
5456           End If

5457           .Range("M8").Select
5458           .ActiveCell.FormulaR1C1 = txtNPSHr.Text

5459           .Range("K1").Select
5460           .ActiveCell.FormulaR1C1 = "Impeller Dia"
5461           .Range("M1").Select


       '        If LenB(txtImpTrim) <> 0 Then
       '            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
       '        Else
       '            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
       '        End If
       '
5462           If chkTrimmed.value = 1 Then
5463               If Val(txtImpTrim.Text) <> 0 Then
5464                   .ActiveCell.FormulaR1C1 = txtImpTrim
5465               Else
5466                   .ActiveCell.FormulaR1C1 = txtImpellerDia
5467               End If
5468           Else
5469               .ActiveCell.FormulaR1C1 = txtImpellerDia
5470           End If



       '        .Range("K1").Select
       '        .ActiveCell.FormulaR1C1 = "KW Mult"
       '        .Range("N1").Select
       '        .ActiveCell.FormulaR1C1 = Val(txtKWMult)

       '        .Range("K2").Select
       '        .ActiveCell.FormulaR1C1 = "HD Cor"
       '        .Range("N2").Select
       '        If Val(txtHDCor) = 0 Then
       '            .ActiveCell.FormulaR1C1 = 0
       '        Else
       '            .ActiveCell.FormulaR1C1 = Val(txtHDCor)
       '        End If

5471           .Range("P9").Select
5472           .ActiveCell.FormulaR1C1 = "Suction Dia"
5473           .Range("R9").Select
5474           .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

5475           .Range("P10").Select
5476           .ActiveCell.FormulaR1C1 = "Discharge Dia"
5477           .Range("R10").Select
5478           .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

5479           .Range("A11").Select
5480           .ActiveCell.FormulaR1C1 = "Test Spec"
5481           .Range("C11").Select
5482           .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

5483           .Range("K3").Select
5484           .ActiveCell.FormulaR1C1 = "Impeller Feathered"
5485           .Range("M3").Select
5486           If chkFeathered.value = 1 Then
5487               .ActiveCell.FormulaR1C1 = "Yes"
5488           Else
5489               .ActiveCell.FormulaR1C1 = "No"
5490           End If

5491           .Range("K2").Select
5492           .ActiveCell.FormulaR1C1 = "Disch Orifice"
5493           .Range("M2").Select
5494           If chkOrifice.value = 1 Then
5495               .ActiveCell.FormulaR1C1 = Val(txtOrifice)
5496           Else
5497               .ActiveCell.FormulaR1C1 = "None"
5498           End If


5499           .Range("K5").Select
5500           .ActiveCell.FormulaR1C1 = "Circulation Orifice"
5501           .Range("M5").Select
5502           If chkCircOrifice.value = 1 Then
5503               .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
5504           Else
5505               .ActiveCell.FormulaR1C1 = "None"
5506           End If

5507           .Range("A13").Select
5508           .ActiveCell.FormulaR1C1 = "Other Mods"
5509           .Range("C13").Select
5510           .ActiveCell.FormulaR1C1 = txtOtherMods

5511           .Range("A14").Select
5512           .ActiveCell.FormulaR1C1 = "Remarks"
5513           .Range("C14").Select
5514           .ActiveCell.FormulaR1C1 = txtRemarks

5515           .Range("A15").Select
5516           .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
5517           .Range("C15").Select
5518           .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

5519           .Range("P1").Select
5520           .ActiveCell.FormulaR1C1 = "Suct ID"
5521           .Range("R1").Select
5522           .ActiveCell.FormulaR1C1 = Me.txtSuctionID.Text

5523           .Range("P2").Select
5524           .ActiveCell.FormulaR1C1 = "Disch ID"
5525           .Range("R2").Select
5526           .ActiveCell.FormulaR1C1 = Me.txtDischargeID.Text

5527           .Range("P3").Select
5528           .ActiveCell.FormulaR1C1 = "Temp ID"
5529           .Range("R3").Select
5530           .ActiveCell.FormulaR1C1 = Me.txtTemperatureID.Text
5531           .Range("P4").Select
5532           .ActiveCell.FormulaR1C1 = "Circ Flow ID"
5533           .Range("R4").Select
5534           .ActiveCell.FormulaR1C1 = Me.txtMagflowID.Text

5535           .Range("P5").Select
5536           .ActiveCell.FormulaR1C1 = "Flow ID"
5537           .Range("R5").Select
5538           .ActiveCell.FormulaR1C1 = Me.txtFlowmeterID.Text 'cmbFlowMeter.List(cmbFlowMeter.ListIndex)

5539           .Range("P6").Select
5540           .ActiveCell.FormulaR1C1 = "Analyzer ID"
5541           .Range("R6").Select
5542           .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

5543           .Range("P7").Select
5544           .ActiveCell.FormulaR1C1 = "Loop ID"
5545           .Range("R7").Select
5546           .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

5547           .Range("A4").Select
5548           .ActiveCell.FormulaR1C1 = "Fluid"
5549           .Range("C4").Select
5550           .ActiveCell.FormulaR1C1 = txtLiquid.Text

5551           .Range("F3").Select
5552           .ActiveCell.FormulaR1C1 = "RMA"
5553           .Range("H3").Select
5554           .ActiveCell.FormulaR1C1 = Me.txtRMA.Text

5555           .Range("F12").Select
5556           .ActiveCell.FormulaR1C1 = "No Of Diodes"
5557           .Range("H12").Select
5558           .ActiveCell.FormulaR1C1 = Me.txtNoOfDiodes.Text

       '        .ActiveCell.FormulaR1C1 = txtRMA.Text
       '        If rsPumpData.Fields("RVSPartNo") <> "" Then
       '            .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
       '        End If
       '        If rsPumpData.Fields("CustPN") <> "" Then
       '            .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
       '        End If

5559           .Range("A7").Select
5560           .ActiveCell.FormulaR1C1 = "Temperature"
5561           .Range("C7").Select
5562           .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

5563           .Range("F6").Select
5564           .ActiveCell.FormulaR1C1 = "Frequency"
5565           .Range("H6").Select
5566           If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
5567               .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
5568           Else
5569               .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
5570           End If
       '        .Range("K2").Select
       '        .ActiveCell.FormulaR1C1 = "Disch Orifice"
       '        .Range("M2").Select
       '        .ActiveCell.FormulaR1C1 = txtOrifice.Text

       '        .Range("K12").Select
       '        .ActiveCell.FormulaR1C1 = "Flow Orifice"
       '        .Range("L12").Select
       '        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

5571           .Range("P8").Select
5572           .ActiveCell.FormulaR1C1 = "PLC No"
5573           .Range("R8").Select
5574           .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

5575           .Range("F7").Select
5576           .ActiveCell.FormulaR1C1 = "Phases"
5577           .Range("H7").Select
5578           .ActiveCell.FormulaR1C1 = txtNoPhases.Text

5579           .Range("F8").Select
5580           .ActiveCell.FormulaR1C1 = "Poles"

5581           .Range("F9").Select
5582           .ActiveCell.FormulaR1C1 = "Rated Current"
5583           .Range("H9").Select
5584           .ActiveCell.FormulaR1C1 = txtAmps.Text

5585           .Range("F10").Select
5586           .ActiveCell.FormulaR1C1 = "Rated Input Power"
5587           .Range("H10").Select
5588           .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

5589           .Range("F11").Select
5590           .ActiveCell.FormulaR1C1 = "Insulation Class"
5591           .Range("H11").Select
5592           .ActiveCell.FormulaR1C1 = txtThermalClass.Text

       '        .Range("P8").Select
       '        .ActiveCell.FormulaR1C1 = "Tach ID"
       '        .Range("R8").Select
       '        .ActiveCell.FormulaR1C1 = cmbTachID.List(cmbTachID.ListIndex)
       '
       '        .Range("P9").Select
       '        .ActiveCell.FormulaR1C1 = "Orifice ID"
       '        .Range("R9").Select
       '        '.ActiveCell.FormulaR1C1 = cmbOrificeNumber.List(cmbOrificeNumber.ListIndex)

           'list the columns starting at row17

5593           .Range("A17").Select
5594           .ActiveCell.FormulaR1C1 = "Flow"
5595           .Range("A18").Select
5596           .ActiveCell.FormulaR1C1 = "(GPM)"

5597           .Range("B17").Select
5598           .ActiveCell.FormulaR1C1 = "TDH"
5599           .Range("B18").Select
5600           .ActiveCell.FormulaR1C1 = "(Ft)"

5601           .Range("C17").Select
5602           .ActiveCell.FormulaR1C1 = "KW"

5603           .Range("D17").Select
5604           .ActiveCell.FormulaR1C1 = "Ave"
5605           .Range("D18").Select
5606           .ActiveCell.FormulaR1C1 = "Volts"

5607           .Range("E17").Select
5608           .ActiveCell.FormulaR1C1 = "Ave"
5609           .Range("E18").Select
5610           .ActiveCell.FormulaR1C1 = "Amps"

5611           .Range("F17").Select
5612           .ActiveCell.FormulaR1C1 = "Power"
5613           .Range("F18").Select
5614           .ActiveCell.FormulaR1C1 = "Factor"

5615           .Range("G17").Select
5616           .ActiveCell.FormulaR1C1 = "Overall"
5617           .Range("G18").Select
5618           .ActiveCell.FormulaR1C1 = "Eff"

5619           .Range("H17").Select
5620           .ActiveCell.FormulaR1C1 = "Measured"
5621           .Range("H18").Select
5622           .ActiveCell.FormulaR1C1 = "RPM"

5623           .Range("I17").Select
5624           .ActiveCell.FormulaR1C1 = "Calculated"
5625           .Range("I18").Select
5626           .ActiveCell.FormulaR1C1 = "RPM"

5627           .Range("J17").Select
5628           .ActiveCell.FormulaR1C1 = "Suction"
5629           .Range("J18").Select
5630           .ActiveCell.FormulaR1C1 = "Temp(F)"

5631           .Range("K17").Select
5632           .ActiveCell.FormulaR1C1 = "Disch"
5633           .Range("K18").Select
5634           .ActiveCell.FormulaR1C1 = "Pressure"

5635           .Range("L17").Select
5636           .ActiveCell.FormulaR1C1 = "Suction"
5637           .Range("L18").Select
5638           .ActiveCell.FormulaR1C1 = "Pressure"

5639           .Range("M17").Select
5640           .ActiveCell.FormulaR1C1 = "Vel"
5641           .Range("M18").Select
5642           .ActiveCell.FormulaR1C1 = "Head"

5643           .Range("N17").Select
5644           .ActiveCell.FormulaR1C1 = "Axial"
5645           .Range("N18").Select
5646           .ActiveCell.FormulaR1C1 = "Position"

5647           .Range("O17").Select
5648           .ActiveCell.FormulaR1C1 = "Pct of"
5649           .Range("O18").Select
5650           .ActiveCell.FormulaR1C1 = "End Play"

5651           .Range("P17").Select
5652           .ActiveCell.FormulaR1C1 = "Hydraulic"
5653           .Range("P18").Select
5654           .ActiveCell.FormulaR1C1 = "Efficiency"

       '        .Range("P17").Select
       '        .ActiveCell.FormulaR1C1 = "Circ"
       '        .Range("P18").Select
       '        .ActiveCell.FormulaR1C1 = "Flow"

5655           .Range("Q17").Select
5656           .ActiveCell.FormulaR1C1 = "Motor"
5657           .Range("Q18").Select
5658           .ActiveCell.FormulaR1C1 = "Efficiency"

5659           .Range("S17").Select
5660           .ActiveCell.FormulaR1C1 = "NPSHa"

5661           .Range("T17").Select
5662           .ActiveCell.FormulaR1C1 = "Phase 1"
5663           .Range("T18").Select
5664           .ActiveCell.FormulaR1C1 = "Current"

5665           .Range("U17").Select
5666           .ActiveCell.FormulaR1C1 = "Phase 2"
5667           .Range("U18").Select
5668           .ActiveCell.FormulaR1C1 = "Current"

5669           .Range("V17").Select
5670           .ActiveCell.FormulaR1C1 = "Phase 3"
5671           .Range("V18").Select
5672           .ActiveCell.FormulaR1C1 = "Current"

5673           .Range("W17").Select
5674           .ActiveCell.FormulaR1C1 = "Phase 1"
5675           .Range("W18").Select
5676           .ActiveCell.FormulaR1C1 = "Voltage"

5677           .Range("X17").Select
5678           .ActiveCell.FormulaR1C1 = "Phase 2"
5679           .Range("X18").Select
5680           .ActiveCell.FormulaR1C1 = "Voltage"

5681           .Range("Y17").Select
5682           .ActiveCell.FormulaR1C1 = "Phase 3"
5683           .Range("Y18").Select
5684           .ActiveCell.FormulaR1C1 = "Voltage"

5685           .Range("Z17").Select
5686           .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

5687           .Range("Z18").Select
5688           .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

5689           .Range("AA17").Select
5690           .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

5691           .Range("AA18").Select
5692           .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

5693           .Range("AB17").Select
5694           .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

5695           .Range("AB18").Select
5696           .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

       '        .Range("AC17").Select
       '        .ActiveCell.FormulaR1C1 = "HR"

       '        .Range("AC18").Select
       '        .ActiveCell.FormulaR1C1 = "(ft)"

5697           .Range("AC17").Select
5698           .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

5699           .Range("AC18").Select
5700           .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

5701           .Range("AD17").Select
5702           .ActiveCell.FormulaR1C1 = "TRG"
5703           .Range("AD18").Select
5704           .ActiveCell.FormulaR1C1 = "Position"

5705           .Range("AE17").Select
5706           .ActiveCell.FormulaR1C1 = "Thrust"

5707           .Range("AF17").Select
5708           .ActiveCell.FormulaR1C1 = "F/R"

5709           .Range("AG17").Select
5710           .ActiveCell.FormulaR1C1 = "Moment"
5711           .Range("AG18").Select
5712           .ActiveCell.FormulaR1C1 = "Arm"

5713           .Range("AH17").Select
5714           .ActiveCell.FormulaR1C1 = "Rig"
5715           .Range("AH18").Select
5716           .ActiveCell.FormulaR1C1 = "Pressure"

       '        .Range("AI17").Select
       '        .ActiveCell.FormulaR1C1 = "Viscosity"

5717           .Range("AI19").Select
5718           .ActiveCell.FormulaR1C1 = "Rear"
5719           .Range("AI18").Select
5720           .ActiveCell.FormulaR1C1 = "Force"

5721           .Range("AJ17").Select
5722           .ActiveCell.FormulaR1C1 = "PV"

5723           .Range("R17").Select
5724           .ActiveCell.FormulaR1C1 = "Shaft"
5725           .Range("R18").Select
5726           .ActiveCell.FormulaR1C1 = "Power"

       '        .Range("AM17").Select
       '        .ActiveCell.FormulaR1C1 = "Pct Full"
       '        .Range("AM18").Select
       '        .ActiveCell.FormulaR1C1 = "Scale"

5727           .Range("AK17").Select
5728           .ActiveCell.FormulaR1C1 = "NPSHr"

5729           .Range("AL17").Select
5730           .ActiveCell.FormulaR1C1 = "Remarks"




               'now output the data

5731           iRowNo = 20

5732           rsEff.MoveFirst
5733           For I = 1 To frmPLCData.UpDown2.value
5734               .Range("A" & iRowNo).Select
5735               .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

5736               .Range("B" & iRowNo).Select
5737               .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

5738               .Range("C" & iRowNo).Select
5739               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

5740               .Range("D" & iRowNo).Select
5741               .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

5742               .Range("E" & iRowNo).Select
5743               .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

5744               .Range("F" & iRowNo).Select
5745               .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

5746               .Range("G" & iRowNo).Select
5747               .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

5748               .Range("H" & iRowNo).Select
5749               .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

5750               .Range("I" & iRowNo).Select
                   'use the coefficients from above to calculate rpm
5751               Dim f As Double
5752               f = .Range("H6").value
5753               .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)

5754               .Range("J" & iRowNo).Select
5755               .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

5756               .Range("K" & iRowNo).Select
5757               .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

5758               .Range("L" & iRowNo).Select
5759               .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

5760               .Range("M" & iRowNo).Select
5761               .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

5762               .Range("N" & iRowNo).Select
5763               .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

5764               .Range("O" & iRowNo).Select
5765               If Val(txtEndPlay) > 0 Then
5766                   .ActiveCell.FormulaR1C1 = 100 * rsEff.Fields("Pos") / Val(txtEndPlay)
5767               End If

5768               .Range("P" & iRowNo).Select
5769               .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

       '            .Range("P" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5770               .Range("Q" & iRowNo).Select
5771               .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

5772               .Range("S" & iRowNo).Select
5773               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

5774               .Range("T" & iRowNo).Select
5775               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

5776               .Range("U" & iRowNo).Select
5777               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

5778               .Range("V" & iRowNo).Select
5779               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

5780               .Range("W" & iRowNo).Select
5781               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

5782               .Range("X" & iRowNo).Select
5783               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

5784               .Range("Y" & iRowNo).Select
5785               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

       '            .Range("Y" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC1")
       '
       '            .Range("Z" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC2")
       '
       '            .Range("AA" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC3")
       '
       '            .Range("AB" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC4")

5786               .Range("Z" & iRowNo).Select
5787               .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5788               .Range("AA" & iRowNo).Select
5789               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

5790               .Range("AB" & iRowNo).Select
5791               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

       '            .Range("AC" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

5792               .Range("AC" & iRowNo).Select
5793               .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

5794               .Range("AD" & iRowNo).Select
5795               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

5796               .Range("AE" & iRowNo).Select

5797               If rsEff.Fields("TEMCFrontThrust") = 0 Then
5798                   If rsEff.Fields("TEMCRearThrust") = 0 Then
5799                       .ActiveCell.FormulaR1C1 = " "
5800                       .Range("AF" & iRowNo).Select
5801                       .ActiveCell.FormulaR1C1 = " "
5802                   Else
5803                       .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
5804                       .Range("AF" & iRowNo).Select
5805                       .ActiveCell.FormulaR1C1 = "R"
5806                   End If
5807               Else
5808                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
5809                   .Range("AF" & iRowNo).Select
5810                   .ActiveCell.FormulaR1C1 = "F"
5811               End If

5812               .Range("AG" & iRowNo).Select
5813               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

5814               .Range("AH" & iRowNo).Select
5815               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

       '            .Range("AJ" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

5816               .Range("AI" & iRowNo).Select
5817               If rsEff.Fields("TEMCForceDirection") = "F" Then
5818                   .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
5819               Else
5820                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
5821               End If

5822               .Range("AJ" & iRowNo).Select
5823               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

5824               .Range("R" & iRowNo).Select
5825               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

5826               .Range("AK" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")

       '            If RatedKW = 999 Then
       '                .ActiveCell.FormulaR1C1 = ""
       '            Else
       '                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
       '            End If

5827               .Range("AL" & iRowNo).Select
5828               .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


5829               rsEff.MoveNext
5830               iRowNo = iRowNo + 1
5831           Next I

5832           .Range("A20:AS30").Select
5833           .Selection.NumberFormat = "0.00"

5834           .Range("N20:N27").Select
5835           .Selection.NumberFormat = "0.000"

           'set up formulas to calculate BEP
           '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
           '  the formulas for doing that are in E68, F68 and G68
           '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
           '
5836       Dim AColumnRow As String
5837       Dim PColumnRow As String

5838       AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
5839       PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

5840           .Range("E68").Select
       '        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

5841           .Range("F68").Select
       '        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

5842           .Range("G68").Select
       '        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

           'export balance holes
5843       If boGotBalanceHoles Then
5844           If rsBalanceHoles.State = adStateClosed Then
5845               rsBalanceHoles.ActiveConnection = cnPumpData
5846               rsBalanceHoles.Open
5847           End If 'rsBalanceHoles.State = adStateClosed

5848           If rsBalanceHoles.RecordCount <> 0 Then

5849               .Range("K9:N9").Merge
5850               .Range("K9:N9").Formula = "Balance Hole Data"
5851               .Range("K9:N9").HorizontalAlignment = xlCenter

5852               .Range("K10").Select
5853               .ActiveCell.Formula = "Date"

5854               .Range("L10").Select
5855               .ActiveCell.Formula = "Number"

5856               .Range("M10").Select
5857               .ActiveCell.Formula = "Diameter"

5858               .Range("N10").Select
5859               .ActiveCell.Formula = "Bolt Circle"

5860               iRowNo = 11

5861               If rsBalanceHoles.RecordCount > 3 Then
5862                   For I = 1 To rsBalanceHoles.RecordCount - 3
5863                       Rows("13:13").Select
5864                       Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
5865                   Next I
5866               End If

5867               rsBalanceHoles.MoveFirst
5868               For I = 1 To rsBalanceHoles.RecordCount

5869                   .Range("K" & iRowNo).Select
5870                   .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
5871                   .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
5872                   .Range("L" & iRowNo).Select
5873                   .ActiveCell = rsBalanceHoles.Fields("Number")
5874                   .ActiveCell.NumberFormat = "0"
5875                   .Range("M" & iRowNo).Select
5876                   If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
5877                       .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
5878                       .ActiveCell.NumberFormat = "0.0000"
5879                   Else
5880                       .ActiveCell = rsBalanceHoles.Fields("Diameter1")
5881                   End If

5882                   .Range("N" & iRowNo).Select
5883                   If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
5884                       .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
5885                       .ActiveCell.NumberFormat = "0.0000"
5886                   Else
5887                       .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
5888                   End If

5889                   rsBalanceHoles.MoveNext
5890                   iRowNo = iRowNo + 1
5891               Next I
5892               .Range("K10:N" & iRowNo - 1).Select
5893               With .Selection.Interior
5894                   .ColorIndex = 34
5895                   .Pattern = xlSolid
5896               End With
5897           End If 'rsBalanceHoles.RecordCount <> 0
5898       End If ' boGotBalanceHoles

           'plot graphs

5899       Dim SeriesName As String
5900       Dim XVals As String
5901       Dim YVals As String
5902       Dim RowNo As Long
5903       Dim RowStr As String
5904       Dim LastPoint As Integer
5905       Dim LineType As String
5906       Dim AxisGroup As Integer
5907       Dim LabelPos As Integer
5908       Dim LineColor As Long

5909           .ActiveSheet.ChartObjects("HydRepChart").Activate
5910           Dim S As Series
               'For Each S In ActiveChart.SeriesCollection
               '    S.Delete
               'Next S

              'determine how many rows of data we have

       '        Range("J86", "J93").Select
       '        With Application.WorksheetFunction
       '            LastPoint = .Match(.Max(Selection), Selection)
       '            RowNo = LastPoint + 85
       '        End With
       '        RowStr = Trim(str(RowNo))

               'find max values to scale chart

               'first TDH
5911           Dim aq As Double
5912           Range("AQ56", "AQ71").Select
5913           aq = .Max(Selection)
5914           Dim ax As Double
5915           Range("AX56", "AX71").Select
5916           ax = .Max(Selection)

               'then current (as and az)
5917           Dim at As Double
5918           Range("AS56", "AS71").Select
5919           at = .Max(Selection)
5920           Dim ba As Double
5921           Range("AZ56", "AZ71").Select
5922           ba = .Max(Selection)

5923           Dim CurrentScaleMax As Integer
5924           Dim TDHScaleMax As Integer

5925           Dim MaxTDH As Integer
5926           With Application.WorksheetFunction
5927               If aq > ax Then
5928                   MaxTDH = .Ceiling(aq, 25)
5929               Else
5930                   MaxTDH = .Ceiling(ax, 25)
5931               End If
5932           End With

5933           Dim MaxCurrent As Integer
5934           With Application.WorksheetFunction
5935               If at > ba Then
5936                   Select Case at
                           Case Is <= 5
5937                           CurrentScaleMax = 5

5938                       Case Is <= 10
5939                           CurrentScaleMax = 10

5940                       Case Else
5941                           CurrentScaleMax = 25
5942                   End Select

5943                   MaxCurrent = .Ceiling(at, CurrentScaleMax)
5944               Else
5945                  Select Case ba
                           Case Is <= 5
5946                           CurrentScaleMax = 5

5947                       Case Is <= 10
5948                           CurrentScaleMax = 10

5949                       Case Else
5950                           CurrentScaleMax = 25
5951                   End Select

5952                   MaxCurrent = .Ceiling(ba, CurrentScaleMax)
5953               End If
5954           End With

5955           ActiveSheet.ChartObjects("HydRepChart").Activate
5956            Dim ShtName As String
5957            ShtName = "'" & ActiveSheet.Name & "'"

5958           RowStr = 56 + 15
5959            For I = 1 To 8

5960                Select Case I
                        Case 1
5961                        SeriesName = "=""TDH"""
5962                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5963                        YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
5964                        LineType = msoLineSolid
5965                        AxisGroup = 1
5966                        LabelPos = xlLabelPositionRight
5967                        LineColor = vbBlue

5968                    Case 2
5969                        SeriesName = "=""Input Power"""
5970                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5971                        YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
5972                        LineType = msoLineSolid
5973                        AxisGroup = 2
5974                        LabelPos = xlLabelPositionRight
5975                        LineColor = vbRed

5976                    Case 3
5977                        SeriesName = "=""Current"""
5978                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5979                        YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
5980                        LineType = msoLineSolid
5981                        AxisGroup = 2
5982                        LabelPos = xlLabelPositionRight
5983                        LineColor = vbGreen

5984                    Case 4
       '                     SeriesName = "=""Overall Eff"""
       '                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
       '                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
       '                     LineType = msoLineSolid
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionRight
       '                     LineColor = vbCyan

5985                    Case 5
5986                        SeriesName = "=""TDH (Adj)"""
5987                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5988                        YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
5989                        LineType = msoLineDash
5990                        AxisGroup = 1
5991                        LabelPos = xlLabelPositionBelow
5992                        LineColor = vbBlue

5993                    Case 6
5994                        SeriesName = "=""Input Power (Adj)"""
5995                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5996                        YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
5997                        LineType = msoLineDash
5998                        AxisGroup = 2
5999                        LabelPos = xlLabelPositionBelow
6000                        LineColor = vbRed

6001                    Case 7
6002                        SeriesName = "=""Current (Adj)"""
6003                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
6004                        YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
6005                        LineType = msoLineDash
6006                        AxisGroup = 2
6007                        LabelPos = xlLabelPositionBelow
6008                        LineColor = vbGreen

6009                    Case 8
       '                     SeriesName = "=""Overall Eff (Adj)"""
       '                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
       '                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
       '                     LineType = msoLineDash
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionBelow
       '                     LineColor = vbCyan

6010               End Select
6011               LastPoint = 16
6012               ActiveChart.SeriesCollection.NewSeries
6013               ActiveChart.SeriesCollection(I).Name = SeriesName
6014               ActiveChart.SeriesCollection(I).XValues = XVals
6015               ActiveChart.SeriesCollection(I).Values = YVals
6016               ActiveChart.SeriesCollection(I).Select
6017               ActiveChart.SeriesCollection(I).Points(LastPoint).Select
6018               ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
6019               ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
6020               If I < 5 Then
6021                   Selection.ShowSeriesName = True
6022                   Selection.Position = LabelPos
6023               Else
6024                   Selection.ShowSeriesName = False
6025               End If
6026               Selection.ShowValue = False
6027               ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
6028               ActiveChart.SeriesCollection(I).Select
6029               With Selection.Format.line
6030                   .Visible = msoTrue
6031                   .DashStyle = LineType
6032                   .ForeColor.RGB = LineColor
6033               End With


6034               ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
6035               ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
6036               ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
6037           Next I

               'show design point
6038           SeriesName = "=""Design Point"""
6039           XVals = "=" & ShtName & "!$L$63"
6040           YVals = "=" & ShtName & "!$L$64"
6041           LineType = msoLineSolid
6042           AxisGroup = 1
6043           ActiveChart.SeriesCollection.NewSeries
6044           ActiveChart.SeriesCollection(I).Name = SeriesName
6045           ActiveChart.SeriesCollection(I).XValues = XVals
6046           ActiveChart.SeriesCollection(I).Values = YVals
6047           ActiveChart.SeriesCollection(I).Select

6048           Selection.MarkerStyle = 4
6049           Selection.MarkerSize = 7
6050           With Selection.Format.line
6051               .Visible = msoTrue
6052               .Weight = 2.25
6053               .ForeColor.RGB = vbBlack
6054           End With


6055           ActiveChart.Axes(xlValue).Select
6056           ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
6057           ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

6058           ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
6059           ActiveChart.Axes(xlValue).MinimumScale = 0
6060           ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
6061           Selection.TickLabels.NumberFormat = "0"

6062           ActiveChart.Axes(xlValue, xlSecondary).Select
6063           ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
6064           ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True

6065           ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
6066           ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
6067           ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
6068           Selection.TickLabels.NumberFormat = "0"

6069           ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
6070           ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
       '        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
6071           ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
               'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"

6072           Range("A1").Select

               'delete all macros in the excel file

               ' Declare variables to access the macros in the workbook.
6073           Dim objProject As VBIDE.VBProject
6074           Dim objComponent As VBIDE.VBComponent
6075           Dim objCode As VBIDE.CodeModule

               ' Get the project details in the workbook.
6076           Set objProject = xlBook.VBProject

               ' Iterate through each component in the project.
6077           For Each objComponent In objProject.VBComponents

                   ' Delete code modules
6078               Set objCode = objComponent.CodeModule
6079               objCode.DeleteLines 1, objCode.CountOfLines

6080               Set objCode = Nothing
6081               Set objComponent = Nothing
6082           Next

6083           Set objProject = Nothing


6084           xlApp.Visible = True                    'show the sheet

       '        xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
       '        xlApp.Run "AssignButton"
6085       End With

       '    Exit Sub

6086   ErrHandler:
           'User pressed the Cancel button

6087       On Error GoTo notopen
6088       If Not xlApp.ActiveWorkbook Is Nothing Then
6089           ActiveWorkbook.CheckCompatibility = False
6090           xlApp.ActiveWorkbook.Save               'save the workbook
               'xlApp.ActiveWorkbook.Close

6091       End If

6092   notopen:

       '    xlApp.Application.Quit

       '    xlApp.Quit
       '    Set xlApp = Nothing

       '    If CommonDialog1.filename <> "" Then
       '        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
       '    End If

6093   On Error GoTo vbwErrHandler

' <VB WATCH>
6094       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6095       Exit Sub
' <VB WATCH>
6096       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6097       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ExportToExcel"

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
            vbwReportVariable "SaveFileName", SaveFileName
            vbwReportVariable "WorkSheetName", WorkSheetName
            vbwReportVariable "I", I
            vbwReportVariable "iRowNo", iRowNo
            vbwReportVariable "sImp", sImp
            vbwReportVariable "ans", ans
            vbwReportVariable "bCanShowSpeed", bCanShowSpeed
            vbwReportVariable "CantShowReason", CantShowReason
            vbwReportVariable "objWMIService", objWMIService
            vbwReportVariable "colProcesses", colProcesses
            vbwReportVariable "xlTemplateName", xlTemplateName
            vbwReportVariable "sheetName", sheetName
            vbwReportVariable "ACoef", ACoef
            vbwReportVariable "BCoef", BCoef
            vbwReportVariable "CCoef", CCoef
            vbwReportVariable "RundownRev", RundownRev
            vbwReportVariable "RPMvalue", RPMvalue
            vbwReportVariable "DesPress", DesPress
            vbwReportVariable "j", j
            vbwReportVariable "f", f
            vbwReportVariable "AColumnRow", AColumnRow
            vbwReportVariable "PColumnRow", PColumnRow
            vbwReportVariable "SeriesName", SeriesName
            vbwReportVariable "XVals", XVals
            vbwReportVariable "YVals", YVals
            vbwReportVariable "RowNo", RowNo
            vbwReportVariable "RowStr", RowStr
            vbwReportVariable "LastPoint", LastPoint
            vbwReportVariable "LineType", LineType
            vbwReportVariable "AxisGroup", AxisGroup
            vbwReportVariable "LabelPos", LabelPos
            vbwReportVariable "LineColor", LineColor
            vbwReportVariable "aq", aq
            vbwReportVariable "ax", ax
            vbwReportVariable "at", at
            vbwReportVariable "ba", ba
            vbwReportVariable "CurrentScaleMax", CurrentScaleMax
            vbwReportVariable "TDHScaleMax", TDHScaleMax
            vbwReportVariable "MaxTDH", MaxTDH
            vbwReportVariable "MaxCurrent", MaxCurrent
            vbwReportVariable "ShtName", ShtName
            vbwReportVariable "xlTemplate", xlTemplate
            vbwReportVariable "TemplateWS", TemplateWS
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportVariable "S", S
            vbwReportVariable "objProject", objProject
            vbwReportVariable "objComponent", objComponent
            vbwReportVariable "objCode", objCode
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


'Private Sub ExportToExcelOrg()
'
'    Dim SaveFileName As String
'    Dim WorkSheetName As String
'
'    Dim I As Integer
'    Dim iRowNo As Integer
'    Dim sImp As String
'    Dim ans As Integer
'
'    Dim bCanShowSpeed As Boolean
'    Dim CantShowReason As String
'
'
'    Set xlApp = New Excel.Application
''    Set xlApp = CreateObject("Excel.Application")
'
'    CommonDialog1.CancelError = True        'in case the user
'    On Error GoTo ErrHandler                '  chooses the cancel button
'
'    'set up dialog box
'    CommonDialog1.DialogTitle = "Open Excel Files"
'    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
'    CommonDialog1.InitDir = App.Path
''    CommonDialog1.InitDir = "C:\"    'in this directory
'    CommonDialog1.ShowOpen                              'open the file selection dialog box
'
'    If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
'        SaveFileName = CommonDialog1.filename           'get the name of the file
'        If Not IsNull(xlApp.Workbooks) Then xlApp.Workbooks.Close   'if there's a workbook open, close it
'        ' Create the Excel Workbook Object.
'        On Error GoTo 0
'        Set xlBook = xlApp.Workbooks.Add                'add a workbook
'        NewWorkBook                                     'do some stuff for the new workbook
'        xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
'            FileFormat:=xlNormal                        'save the file
'    Else                                                'the file name already exists
'        SaveFileName = CommonDialog1.filename
'        ' Create the Excel Workbook Object.
'        If Not IsNull(xlApp.Workbooks) Then xlApp.Workbooks.Close   'if there's a workbook open, close it
'        Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
'        If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
'            MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
'            Exit Sub
'        Else
'        End If
'    End If
'
'    On Error GoTo 0
'
'    'see if we can export Speed and SG and if we can, ask user if s/he wants it
'    'assume that we can show speed calcs
'
'    bCanShowSpeed = True
'
'      'get the coefficients
'    Dim qy As New ADODB.Command
'    Dim rs As New ADODB.Recordset
'    Dim NumberOfCalculations As Integer
'
'    If (cmbMotor.ListIndex <> -1) Then          'it's a chempump pump
'        If (cmbMotor.ItemData(cmbMotor.ListIndex) < 7) Then  'N1 through N8
'
'            'select the testsetup data for the serial number
'            qy.ActiveConnection = cnPumpData
'
'            If cmbStatorFill.ItemData(cmbStatorFill.ListIndex) = 1 Then 'dry stator fill
'                qy.CommandText = "SELECT * FROM Coefficients WHERE (((Coefficients.MotorKey)=" & cmbMotor.ItemData(cmbMotor.ListIndex) & ") AND ((Coefficients.Filled)='Dry')) ;"
'            Else
'                qy.CommandText = "SELECT * FROM Coefficients WHERE (((Coefficients.MotorKey)=" & cmbMotor.ItemData(cmbMotor.ListIndex) & ") AND ((Coefficients.Filled)='Filled')) ;"
'            End If
'
'            With rs     'open the recordset for the query
'                .CursorLocation = adUseServer
'                .CursorType = adOpenDynamic
'                .Open qy
'            End With
'
'            If Not rs.BOF = True And rs.EOF = True Then
'                bCanShowSpeed = False
'                CantShowReason = "No Coefficients for this Motor / Fill combination"
'            End If
'
'            If bCanShowSpeed Then
'                rsEff.MoveFirst
'                rsEff.Move frmPLCData.UpDown2.value - 1    'get last data
'
'                'make sure there is kw data
'                Do While Not rsEff.BOF
'                    bCanShowSpeed = False
'                    CantShowReason = "No entries with kW>0."
'                    If rsEff.Fields("KW") <> 0 Then
'                        bCanShowSpeed = True
'                        Exit Do
'                    End If
'                    rsEff.MovePrevious
'                Loop
'
'                rs.Filter = "Parameter = 'Speed'"
'                If bCanShowSpeed Then
'                    NumberOfCalculations = CalculateSpeed(rs.Fields("Squared"), rs.Fields("Linear"), rs.Fields("Constant"), rsEff.Fields("KW") * (rsEff.Fields("MotorEfficiency") / 100) / 0.7457, CDbl(Val(txtSpGr.Text)))
'                Else
'                    NumberOfCalculations = -1
'                End If
'                If NumberOfCalculations > 15 Or NumberOfCalculations = 0 Then
'                    bCanShowSpeed = False
'                    CantShowReason = "More than 15 calculations needed."
'                ElseIf NumberOfCalculations = -1 Then
'                    bCanShowSpeed = False
'                    CantShowReason = "No entries with kW>0."
'                End If
'            End If
'        Else
'            bCanShowSpeed = False
'            CantShowReason = "Motor Not N1-N8"
'        End If
'    ElseIf Me.txtTEMCFrameNumber <> "" Then
'
'            'select the testsetup data for the serial number
'            qy.ActiveConnection = cnPumpData
'
'            qy.CommandText = "SELECT * FROM TEMCCoefficients WHERE (((TEMCCoefficients.Frame)=" & txtTEMCFrameNumber.Text & ") ) ;"
'
'            With rs     'open the recordset for the query
'                .CursorLocation = adUseServer
'                .CursorType = adOpenDynamic
'                .Open qy
'            End With
'
'            If (rs.BOF = True And rs.EOF = True) Then
'                bCanShowSpeed = False
'                CantShowReason = "No Coefficients for this TEMC Frame size"
'            End If
'
'            If bCanShowSpeed Then
'                rsEff.MoveFirst
'                rsEff.Move frmPLCData.UpDown2.value - 1    'get last data
'
'                'make sure there is kw data
'                Do While Not rsEff.BOF
'                    bCanShowSpeed = False
'                    CantShowReason = "No entries with kW>0."
'                    If rsEff.Fields("KW") <> 0 Then
'                        bCanShowSpeed = True
'                        Exit Do
'                    End If
'                    rsEff.MovePrevious
'                Loop
'
'                rs.Filter = "Parameter = 'Speed'"
'                If bCanShowSpeed Then
'                    NumberOfCalculations = CalculateSpeed(rs.Fields("Squared"), rs.Fields("Linear"), rs.Fields("Constant"), rsEff.Fields("KW") * (rsEff.Fields("MotorEfficiency") / 100) / 0.7457, CDbl(txtSpGr.Text))
'                Else
'                    NumberOfCalculations = -1
'                End If
'                If NumberOfCalculations > 15 Or NumberOfCalculations = 0 Then
'                    bCanShowSpeed = False
'                    CantShowReason = "More than 15 calculations needed."
'                ElseIf NumberOfCalculations = -1 Then
'                    bCanShowSpeed = False
'                    CantShowReason = "No entries with kW>0."
'                End If
'            End If
'    Else
'        bCanShowSpeed = False
'        CantShowReason = "Motor Not N1-N8"
'    End If
'
'
'    If bCanShowSpeed Then
'        ans = MsgBox("Do you want to show Speed / SG calculations on the spreadsheet?", vbYesNo, "Show Calculations?")
'        If ans = vbNo Then
'            bCanShowSpeed = False
'        End If
'    Else
'        MsgBox "Cannot show Speed / SG calculations since " & CantShowReason, vbOKOnly, "Can't Show Calculations"
'    End If
'
'    'write the data to the spreadsheet
'    With xlApp
'
'    'write header data
'
'        .Range("A1").Select
'        .ActiveCell.FormulaR1C1 = "Serial Number"
'        .Range("D1").Select
'        .ActiveCell.FormulaR1C1 = txtSN
'
'        .Range("A2").Select
'        .ActiveCell.FormulaR1C1 = "Customer"
'        .Range("D2").Select
'        .ActiveCell.FormulaR1C1 = txtShpNo
'
'        .Range("A3").Select
'        .ActiveCell.FormulaR1C1 = "Model"
'        .Range("D3").Select
'        .ActiveCell.FormulaR1C1 = txtModelNo
'
'        .Range("A4").Select
'        .ActiveCell.FormulaR1C1 = "Sales Order"
'        .Range("D4").Select
'        .ActiveCell.FormulaR1C1 = txtSalesOrderNumber
'
'        .Range("A5").Select
'        .ActiveCell.FormulaR1C1 = "Design Flow"
'        .Range("D5").Select
'        .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)
'
'        .Range("A6").Select
'        .ActiveCell.FormulaR1C1 = "Design Head"
'        .Range("D6").Select
'        .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)
'
'        .Range("A7").Select
'        .ActiveCell.FormulaR1C1 = "Barometric Pressure"
'        .Range("D7").Select
'        .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)
'
'        .Range("A8").Select
'        .ActiveCell.FormulaR1C1 = "Suction Gage Height"
'        .Range("D8").Select
'        .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)
'
'        .Range("A9").Select
'        .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
'        .Range("D9").Select
'        .ActiveCell.FormulaR1C1 = Val(txtDischHeight)
'
'        .Range("A10").Select
'        .ActiveCell.FormulaR1C1 = "Run Date"
'        .Range("D10").Select
'        .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)
'
'        .Range("D10:E10").Select
'        With xlApp.Selection
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlBottom
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = False
'        End With
'        xlApp.Selection.Merge
'
'        .Range("G1").Select
'        .ActiveCell.FormulaR1C1 = "RPM"
'        .Range("I1").Select
'        .ActiveCell.FormulaR1C1 = cmbRPM.List(cmbRPM.ListIndex)
'
'        .Range("G2").Select
'        .ActiveCell.FormulaR1C1 = "Sp Gravity"
'        .Range("I2").Select
'        .ActiveCell.FormulaR1C1 = txtSpGr
'
'        .Range("G3").Select
'        .ActiveCell.FormulaR1C1 = "Motor"
'        .Range("I3").Select
'        .ActiveCell.FormulaR1C1 = cmbMotor.List(cmbMotor.ListIndex)
'
'        .Range("G4").Select
'        .ActiveCell.FormulaR1C1 = "Voltage"
'        .Range("I4").Select
'        .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)
'
'        .Range("G5").Select
'        .ActiveCell.FormulaR1C1 = "End Play"
'        .Range("I5").Select
'        .ActiveCell.FormulaR1C1 = Val(txtEndPlay)
'
'        .Range("G6").Select
'        .ActiveCell.FormulaR1C1 = "Design Pressure"
'        .Range("I6").Select
'        .ActiveCell.FormulaR1C1 = cmbDesignPressure.List(cmbDesignPressure.ListIndex)
'
'        .Range("G7").Select
'        .ActiveCell.FormulaR1C1 = "Stator Fill"
'        .Range("I7").Select
'        .ActiveCell.FormulaR1C1 = cmbStatorFill.List(cmbStatorFill.ListIndex)
'
'        .Range("G8").Select
'        .ActiveCell.FormulaR1C1 = "Circulation Path"
'        .Range("I8").Select
'        .ActiveCell.FormulaR1C1 = cmbCirculationPath.List(cmbCirculationPath.ListIndex)
'
'        .Range("G9").Select
'        .ActiveCell.FormulaR1C1 = "Impeller Dia"
'        .Range("I9").Select
'
''        If LenB(txtImpTrim) <> 0 Then
''            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
''        Else
''            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
''        End If
''
'        If chkTrimmed.value = 1 Then
'            If Val(txtImpTrim.Text) <> 0 Then
'                .ActiveCell.FormulaR1C1 = txtImpTrim
'            Else
'                .ActiveCell.FormulaR1C1 = txtImpellerDia
'            End If
'        Else
'            .ActiveCell.FormulaR1C1 = txtImpellerDia
'        End If
'
'
'
'        .Range("K1").Select
'        .ActiveCell.FormulaR1C1 = "KW Mult"
'        .Range("N1").Select
'        .ActiveCell.FormulaR1C1 = Val(txtKWMult)
'
'        .Range("K2").Select
'        .ActiveCell.FormulaR1C1 = "HD Cor"
'        .Range("N2").Select
'        If Val(txtHDCor) = 0 Then
'            .ActiveCell.FormulaR1C1 = 0
'        Else
'            .ActiveCell.FormulaR1C1 = Val(txtHDCor)
'        End If
'
'        .Range("K3").Select
'        .ActiveCell.FormulaR1C1 = "Suction Dia"
'        .Range("N3").Select
'        .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)
'
'        .Range("K4").Select
'        .ActiveCell.FormulaR1C1 = "Discharge Dia"
'        .Range("N4").Select
'        .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)
'
'        .Range("K5").Select
'        .ActiveCell.FormulaR1C1 = "Test Spec"
'        .Range("N5").Select
'        .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)
'
'        .Range("K6").Select
'        .ActiveCell.FormulaR1C1 = "Imp Feathered"
'        .Range("N6").Select
'        If chkFeathered.value = 1 Then
'            .ActiveCell.FormulaR1C1 = "Yes"
'        Else
'            .ActiveCell.FormulaR1C1 = "No"
'        End If
'
'        .Range("K7").Select
'        .ActiveCell.FormulaR1C1 = "Disch Orifice"
'        .Range("N7").Select
'        If chkOrifice.value = 1 Then
'            .ActiveCell.FormulaR1C1 = Val(txtOrifice)
'        Else
'            .ActiveCell.FormulaR1C1 = "None"
'        End If
'
'
'        .Range("K8").Select
'        .ActiveCell.FormulaR1C1 = "Circulation Orifice"
'        .Range("N8").Select
'        If chkCircOrifice.value = 1 Then
'            .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
'        Else
'            .ActiveCell.FormulaR1C1 = "None"
'        End If
'
'        .Range("A12").Select
'        .ActiveCell.FormulaR1C1 = "Other Mods"
'        .Range("C12").Select
'        .ActiveCell.FormulaR1C1 = txtOtherMods
'
'        .Range("A13").Select
'        .ActiveCell.FormulaR1C1 = "Remarks"
'        .Range("C13").Select
'        .ActiveCell.FormulaR1C1 = txtRemarks
'
'        .Range("A14").Select
'        .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
'        .Range("C14").Select
'        .ActiveCell.FormulaR1C1 = txtTestSetupRemarks
'
'        .Range("P1").Select
'        .ActiveCell.FormulaR1C1 = "Suct ID"
'        .Range("R1").Select
'        .ActiveCell.FormulaR1C1 = txtSuctionID
'
'        .Range("P2").Select
'        .ActiveCell.FormulaR1C1 = "Disch ID"
'        .Range("R2").Select
'        .ActiveCell.FormulaR1C1 = txtDischargeID
'
'        .Range("P3").Select
'        .ActiveCell.FormulaR1C1 = "Temp ID"
'        .Range("R3").Select
'        .ActiveCell.FormulaR1C1 = txtTemperatureID
'
'        .Range("P4").Select
'        .ActiveCell.FormulaR1C1 = "Circ Flow ID"
'        .Range("R4").Select
'        .ActiveCell.FormulaR1C1 = txtMagflowID
'
'        .Range("P5").Select
'        .ActiveCell.FormulaR1C1 = "Flow ID"
'        .Range("R5").Select
'        .ActiveCell.FormulaR1C1 = txtFlowmeterID
'
'        .Range("P6").Select
'        .ActiveCell.FormulaR1C1 = "Analyzer ID"
'        .Range("R6").Select
'        .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)
'
'        .Range("P7").Select
'        .ActiveCell.FormulaR1C1 = "Loop ID"
'        .Range("R7").Select
'        .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)
'
'        .Range("P8").Select
'        .ActiveCell.FormulaR1C1 = "Tach ID"
'        .Range("R8").Select
'        .ActiveCell.FormulaR1C1 = cmbTachID.List(cmbTachID.ListIndex)
'
'        .Range("P9").Select
'        .ActiveCell.FormulaR1C1 = "Orifice ID"
'        .Range("R9").Select
'        .ActiveCell.FormulaR1C1 = cmbOrificeNumber.List(cmbOrificeNumber.ListIndex)
'
'    'list the columns starting at row17
'
'        .Range("A17").Select
'        .ActiveCell.FormulaR1C1 = "Flow"
'        .Range("A18").Select
'        .ActiveCell.FormulaR1C1 = "(GPM)"
'
'        .Range("B17").Select
'        .ActiveCell.FormulaR1C1 = "TDH"
'        .Range("B18").Select
'        .ActiveCell.FormulaR1C1 = "(Ft)"
'
'        .Range("C17").Select
'        .ActiveCell.FormulaR1C1 = "KW"
'
'        .Range("D17").Select
'        .ActiveCell.FormulaR1C1 = "Ave"
'        .Range("D18").Select
'        .ActiveCell.FormulaR1C1 = "Volts"
'
'        .Range("E17").Select
'        .ActiveCell.FormulaR1C1 = "Ave"
'        .Range("E18").Select
'        .ActiveCell.FormulaR1C1 = "Amps"
'
'        .Range("F17").Select
'        .ActiveCell.FormulaR1C1 = "Power"
'        .Range("F18").Select
'        .ActiveCell.FormulaR1C1 = "Factor"
'
'        .Range("G17").Select
'        .ActiveCell.FormulaR1C1 = "Overall"
'        .Range("G18").Select
'        .ActiveCell.FormulaR1C1 = "Eff"
'
'        .Range("H17").Select
'        .ActiveCell.FormulaR1C1 = "RPM"
'
'        .Range("I17").Select
'        .ActiveCell.FormulaR1C1 = "Suction"
'        .Range("I18").Select
'        .ActiveCell.FormulaR1C1 = "Temp(F)"
'
'        .Range("J17").Select
'        .ActiveCell.FormulaR1C1 = "Disch"
'        .Range("J18").Select
'        .ActiveCell.FormulaR1C1 = "Pressure"
'
'        .Range("K17").Select
'        .ActiveCell.FormulaR1C1 = "Suction"
'        .Range("K18").Select
'        .ActiveCell.FormulaR1C1 = "Pressure"
'
'        .Range("L17").Select
'        .ActiveCell.FormulaR1C1 = "Vel"
'        .Range("L18").Select
'        .ActiveCell.FormulaR1C1 = "Head"
'
'        .Range("M17").Select
'        .ActiveCell.FormulaR1C1 = "Axial"
'        .Range("M18").Select
'        .ActiveCell.FormulaR1C1 = "Position"
'
'        .Range("N17").Select
'        .ActiveCell.FormulaR1C1 = "Hydraulic"
'        .Range("N18").Select
'        .ActiveCell.FormulaR1C1 = "Efficiency"
'
'        .Range("O17").Select
'        .ActiveCell.FormulaR1C1 = "Circ"
'        .Range("O18").Select
'        .ActiveCell.FormulaR1C1 = "Flow"
'
'        .Range("P17").Select
'        .ActiveCell.FormulaR1C1 = "Motor"
'        .Range("P18").Select
'        .ActiveCell.FormulaR1C1 = "Efficiency"
'
'        .Range("Q17").Select
'        .ActiveCell.FormulaR1C1 = "NPSHa"
'
'        .Range("R17").Select
'        .ActiveCell.FormulaR1C1 = "Phase 1"
'        .Range("R18").Select
'        .ActiveCell.FormulaR1C1 = "Current"
'
'        .Range("S17").Select
'        .ActiveCell.FormulaR1C1 = "Phase 2"
'        .Range("S18").Select
'        .ActiveCell.FormulaR1C1 = "Current"
'
'        .Range("T17").Select
'        .ActiveCell.FormulaR1C1 = "Phase 3"
'        .Range("T18").Select
'        .ActiveCell.FormulaR1C1 = "Current"
'
'        .Range("U17").Select
'        .ActiveCell.FormulaR1C1 = "Phase 1"
'        .Range("U18").Select
'        .ActiveCell.FormulaR1C1 = "Voltage"
'
'        .Range("V17").Select
'        .ActiveCell.FormulaR1C1 = "Phase 2"
'        .Range("V18").Select
'        .ActiveCell.FormulaR1C1 = "Voltage"
'
'        .Range("W17").Select
'        .ActiveCell.FormulaR1C1 = "Phase 3"
'        .Range("W18").Select
'        .ActiveCell.FormulaR1C1 = "Voltage"
'
'        .Range("X17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(0).Text
'
'        .Range("X18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(1).Text
'
'        .Range("Y17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(2).Text
'
'        .Range("Y18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(3).Text
'
'        .Range("Z17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(4).Text
'
'        .Range("Z18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(5).Text
'
'        .Range("AA17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(6).Text
'
'        .Range("AA18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(7).Text
'
'        .Range("AB17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text
'
'        .Range("AB18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text
'
'        .Range("AC17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text
'
'        .Range("AC18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text
'
'        .Range("AD17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text
'
'        .Range("AD18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text
'
'        .Range("AE17").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text
'
'        .Range("AE18").Select
'        .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text
'
'        .Range("AF17").Select
'        .ActiveCell.FormulaR1C1 = "TRG"
'        .Range("AF18").Select
'        .ActiveCell.FormulaR1C1 = "Position"
'        .Range("AG17").Select
'        .ActiveCell.FormulaR1C1 = "Thrust"
'
'        .Range("AH17").Select
'        .ActiveCell.FormulaR1C1 = "F/R"
'
'        .Range("AI17").Select
'        .ActiveCell.FormulaR1C1 = "Moment"
'        .Range("AI18").Select
'        .ActiveCell.FormulaR1C1 = "Arm"
'        .Range("AJ17").Select
'        .ActiveCell.FormulaR1C1 = "Rig"
'        .Range("AJ18").Select
'        .ActiveCell.FormulaR1C1 = "Pressure"
'
'        .Range("AK17").Select
'        .ActiveCell.FormulaR1C1 = "Viscosity"
'        .Range("AL17").Select
'        .ActiveCell.FormulaR1C1 = "Rear"
'        .Range("AL18").Select
'        .ActiveCell.FormulaR1C1 = "Force"
'        .Range("AM17").Select
'        .ActiveCell.FormulaR1C1 = "PV"
'        .Range("AN17").Select
'        .ActiveCell.FormulaR1C1 = "Shaft"
'        .Range("AN18").Select
'        .ActiveCell.FormulaR1C1 = "Power"
'        .Range("AO17").Select
'        .ActiveCell.FormulaR1C1 = "Pct Full"
'        .Range("AO18").Select
'        .ActiveCell.FormulaR1C1 = "Scale"
'
'        .Range("AP17").Select
'        .ActiveCell.FormulaR1C1 = "Remarks"
'
'        'now output the data
'
'        iRowNo = 20
'
'        rsEff.MoveFirst
'        For I = 1 To frmPLCData.UpDown2.value
'            .Range("A" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")
'
'            .Range("B" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")
'
'            .Range("C" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")
'
'            .Range("D" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")
'
'            .Range("E" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")
'
'            .Range("F" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")
'
'            .Range("G" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")
'
'            .Range("H" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")
'
'            .Range("I" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")
'
'            .Range("J" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")
'
'            .Range("K" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")
'
'            .Range("L" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")
'
'            .Range("M" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")
'
'            .Range("N" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")
'
'            .Range("O" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")
'
'            .Range("P" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")
'
'            .Range("Q" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")
'
'            .Range("R" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")
'
'            .Range("S" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")
'
'            .Range("T" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")
'
'            .Range("U" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")
'
'            .Range("V" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")
'
'            .Range("W" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")
'
'            .Range("X" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC1")
'
'            .Range("Y" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC2")
'
'            .Range("Z" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC3")
'
'            .Range("AA" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC4")
'
'            .Range("AB" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")
'
'            .Range("AC" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")
'
'            .Range("AD" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")
'
'            .Range("AE" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")
'
'            .Range("AF" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")
'
'            .Range("AG" & iRowNo).Select
'            If rsEff.Fields("TEMCFrontThrust") = 0 Then
'                If rsEff.Fields("TEMCRearThrust") = 0 Then
'                    .ActiveCell.FormulaR1C1 = " "
'                    .Range("AH" & iRowNo).Select
'                    .ActiveCell.FormulaR1C1 = " "
'                Else
'                    .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
'                    .Range("AH" & iRowNo).Select
'                    .ActiveCell.FormulaR1C1 = "R"
'                End If
'            Else
'                .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
'                .Range("AH" & iRowNo).Select
'                .ActiveCell.FormulaR1C1 = "F"
'            End If
'
'            .Range("AI" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")
'
'            .Range("AJ" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")
'
'            .Range("AK" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")
'
'            .Range("AL" & iRowNo).Select
'            If rsEff.Fields("TEMCForceDirection") = "F" Then
'                .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
'            Else
'                .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
'            End If
'
'            .Range("AM" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")
'
'            .Range("AN" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100
'
'            .Range("AO" & iRowNo).Select
'            If RatedKW = 999 Then
'                .ActiveCell.FormulaR1C1 = ""
'            Else
'                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
'            End If
'
'            .Range("AP" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")
'
'
'            rsEff.MoveNext
'            iRowNo = iRowNo + 1
'        Next I
'
'    If bCanShowSpeed Then   'if we're able to show the speed data
'        'coeff box
'        .Range("B29").Select
'        .ActiveCell.Formula = "=""2nd Order Polynomial Coefficients for "" & $H$3"
'
'        .Range("B31").Select
'        .ActiveCell.Formula = "Squared"
'
'        .Range("B32").Select
'        .ActiveCell.Formula = "Linear"
'
'        .Range("B33").Select
'        .ActiveCell.Formula = "Constant"
'
'        .Range("C30").Select
'        .ActiveCell.Formula = "HP vs RPM"
'
'        .Range("E30").Select
'        .ActiveCell.Formula = "HP vs Amps"
'
'        .Range("G30").Select
'        .ActiveCell.Formula = "HP vs KW in"
'
'        .Range("C31").Select
'        .ActiveCell.Formula = rs.Fields("Squared")
'
'        .Range("C32").Select
'        .ActiveCell.Formula = rs.Fields("Linear")
'
'        .Range("C33").Select
'        .ActiveCell.Formula = rs.Fields("Constant")
'
'        rs.Filter = "Parameter = 'Amps'"
'
'        .Range("E31").Select
'        .ActiveCell.Formula = rs.Fields("Squared")
'
'        .Range("E32").Select
'        .ActiveCell.Formula = rs.Fields("Linear")
'
'        .Range("E33").Select
'        .ActiveCell.Formula = rs.Fields("Constant")
'
'        rs.Filter = "Parameter = 'KWIn'"
'
'        .Range("G31").Select
'        .ActiveCell.Formula = rs.Fields("Squared")
'
'        .Range("G32").Select
'        .ActiveCell.Formula = rs.Fields("Linear")
'
'        .Range("G33").Select
'        .ActiveCell.Formula = rs.Fields("Constant")
'
'        'data
'        .Range("A35").Select
'        .ActiveCell.Formula = "SG = 1"
'
'        .Range("C35").Select
'        .ActiveCell.Formula = "=""SG = "" & $H$2"
'
'        .Range("A36").Select
'        .ActiveCell.Formula = "Shaft HP"
'
'        .Range("B36").Select
'        .ActiveCell.Formula = "Calc Speed"
'
'        .Range("A37").Select
'        .ActiveCell.Formula = "=C20*(P20/100)/0.7457"
'
'        .Range("B37").Select
'        .ActiveCell.Formula = "=$C$31*A37^2+$C$32*A37+$C$33"
'
'        .Range("C36").Select
'        .ActiveCell.Formula = "Shaft HP"
'
'        .Range("D36").Select
'        .ActiveCell.Formula = "Calc Speed"
'
'        .Range("C37").Select
'        .ActiveCell.Formula = "=A37*$H$2"
'
'        .Range("D37").Select
'        .ActiveCell.Formula = "=$C$31*C37^2+$C$32*C37+$C$33"
'
'        .Range("E36").Select
'        .ActiveCell.Formula = "Shaft HP"
'
'        .Range("F36").Select
'        .ActiveCell.Formula = "Calc Speed"
'
'        .Range("E37").Select
'        .ActiveCell.Formula = "=C37*(D37/B37)^3"
'
'        .Range("F37").Select
'        .ActiveCell.Formula = "=$C$31*E37^2+$C$32*E37+$C$33"
'
'        'copy the calculations
'        .Range("E36:F37").Select
'        .Selection.Copy
'
'        'and paste them
'        For I = 1 To NumberOfCalculations - 3
'            .Cells(36, 5 + 2 * I).Select
'            .ActiveSheet.Paste
'        Next I
'
'        .Range(.Cells(37, 1), .Cells(37, 5 + 2 * I)).Select
'        .Selection.AutoFill Destination:=.Range(.Cells(37, 1), .Cells(44, 5 + 2 * I)), Type:=xlFillDefault
'
'        .Range("A46").Select
'        .ActiveCell.Formula = "Speed Delta"
'
'        For I = 1 To NumberOfCalculations - 1
'            .Cells(46, 2 + 2 * (I)).Select
'            .ActiveCell.FormulaR1C1 = "=R[-2]C[-2]-R[-2]C"
'        Next I
'
'        'final numbers
'        .Range("B49").Select
'        .ActiveCell.Formula = "=""Final Calculated Data at SG = "" & $H$2"
'
'        .Range("B50").Select
'        .ActiveCell.Formula = "Flow"
'
'        .Range("C50").Select
'        .ActiveCell.Formula = "Head"
'
'        .Range("D50").Select
'        .ActiveCell.Formula = "Input"
'
'        .Range("E50").Select
'        .ActiveCell.Formula = "Amps"
'
'        .Range("B51").Select
'        .ActiveCell.Formula = "=R[-31]C[-1]*(R[-14]C[" & 2 * (I - 1) & "]/R[-14]C)"
'
'        .Range("C51").Select
'        .ActiveCell.Formula = "=R[-31]C[-1]*(R[-14]C[" & (2 * (I - 1) - 1) & "]/R[-14]C[-1])^2"
'
'        .Range("D51").Select
'        .ActiveCell.Formula = "=R31C7*R[-14]C[" & (2 * (I - 1) - 3) & "]^2+R32C7*R[-14]C[" & (2 * (I - 1) - 3) & "] + R33C7"
'
'        .Range("E51").Select
'        .ActiveCell.Formula = "=R31C5*R[-14]C[" & (2 * (I - 1) - 4) & "]^2+R32C5*R[-14]C[" & (2 * (I - 1) - 4) & "]+R33C5"
'
'        .Range("B51:E51").Select
'        .Selection.AutoFill Destination:=.Range("B51:E58"), Type:=xlFillDefault
'
'        FixCoef
'        FixFormat
'    End If 'bCanShowSpeed
'
'    'export balance holes
'    If boGotBalanceHoles Then
'        If rsBalanceHoles.State = adStateClosed Then
'            rsBalanceHoles.ActiveConnection = cnPumpData
'            rsBalanceHoles.Open
'        End If 'rsBalanceHoles.State = adStateClosed
'
'        If rsBalanceHoles.RecordCount <> 0 Then
'
'            .Range("M11:P11").Merge
'            .Range("M12:P11").Formula = "Balance Hole Data"
'            .Range("M11:P11").HorizontalAlignment = xlCenter
'
'            .Range("M12").Select
'            .ActiveCell.Formula = "Date"
'
'            .Range("N12").Select
'            .ActiveCell.Formula = "Number"
'
'            .Range("O12").Select
'            .ActiveCell.Formula = "Diameter"
'
'            .Range("P12").Select
'            .ActiveCell.Formula = "Bolt Circle"
'
'            iRowNo = 13
'
'            If rsBalanceHoles.RecordCount > 3 Then
'                For I = 1 To rsBalanceHoles.RecordCount - 3
'                    Rows("16:16").Select
'                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'                Next I
'            End If
'
'            rsBalanceHoles.MoveFirst
'            For I = 1 To rsBalanceHoles.RecordCount
'
'                .Range("M" & iRowNo).Select
'                .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
'                .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
'                .Range("N" & iRowNo).Select
'                .ActiveCell = rsBalanceHoles.Fields("Number")
'                .ActiveCell.NumberFormat = "0"
'                .Range("O" & iRowNo).Select
'                If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
'                    .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
'                    .ActiveCell.NumberFormat = "0.0000"
'                Else
'                    .ActiveCell = rsBalanceHoles.Fields("Diameter1")
'                End If
'
'                .Range("P" & iRowNo).Select
'                If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
'                    .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
'                    .ActiveCell.NumberFormat = "0.0000"
'                Else
'                    .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
'                End If
'
'                rsBalanceHoles.MoveNext
'                iRowNo = iRowNo + 1
'            Next I
'            .Range("M12:P" & iRowNo - 1).Select
'            With .Selection.Interior
'                .ColorIndex = 34
'                .Pattern = xlSolid
'            End With
'        End If 'rsBalanceHoles.RecordCount <> 0
'    End If ' boGotBalanceHoles
'
'    End With
'
''    Exit Sub
'
'ErrHandler:
'    'User pressed the Cancel button
'
'    On Error GoTo notopen
'    xlApp.ActiveWorkbook.Save               'save the workbook
''    xlApp.Visible = True                    'show the sheet
'
'notopen:
'
'    xlApp.Application.Quit
'
'    xlApp.Quit
'    Set xlApp = Nothing
'
'    If CommonDialog1.filename <> "" Then
'        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
'    End If
'
'    On Error GoTo 0
'
'    Exit Sub
'End Sub
Function GetWorksheetTabs(filename As String, WorkSheetName As String)
' <VB WATCH>
6098       On Error GoTo vbwErrHandler
6099       Const VBWPROCNAME = "frmPLCData.GetWorksheetTabs"
6100       If vbwProtector.vbwTraceProc Then
6101           Dim vbwProtectorParameterString As String
6102           If vbwProtector.vbwTraceParameters Then
6103               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("filename", filename) & ", "
6104               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("WorkSheetName", WorkSheetName) & ") "
6105           End If
6106           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6107       End If
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

6108       Dim intSheets As Integer    'number of sheets in the workbook
6109       Dim I As Integer
6110       Dim S As String
6111       Dim ans
6112       Dim NameOK As Boolean

6113       intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
6114       S = vbCrLf

6115       For I = 1 To intSheets
6116           S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
6117       Next I

           'tell the user the names so far and ask if he/she wants to add another
6118       ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
6119       If ans = vbNo Then
6120           GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
' <VB WATCH>
6121       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6122           Exit Function
6123       End If

           'get worksheet name from user and check to see that it's not already used

6124       NameOK = False  'start assuming that the name is bad

6125       While Not NameOK    'as long as it's bad, stay in this loop
6126           WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

6127           If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
6128               GetWorksheetTabs = vbNo
' <VB WATCH>
6129       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6130               Exit Function
6131           End If

6132           For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
6133               If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
6134                   MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
6135                   NameOK = False
6136                   Exit For
6137               End If
6138               NameOK = True       'if we make it thru say the name is ok
6139           Next I
6140       Wend

6141       xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
6142       xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
6143       GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
6144       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6145       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetWorksheetTabs"

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
            vbwReportVariable "filename", filename
            vbwReportVariable "WorkSheetName", WorkSheetName
            vbwReportVariable "intSheets", intSheets
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportVariable "ans", ans
            vbwReportVariable "NameOK", NameOK
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function NewWorkBook() As String
' <VB WATCH>
6146       On Error GoTo vbwErrHandler
6147       Const VBWPROCNAME = "frmPLCData.NewWorkBook"
6148       If vbwProtector.vbwTraceProc Then
6149           Dim vbwProtectorParameterString As String
6150           If vbwProtector.vbwTraceParameters Then
6151               vbwProtectorParameterString = "()"
6152           End If
6153           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6154       End If
' </VB WATCH>

6155       Dim WorkSheetName As String

           'we've just added a new workbook, delete sheet1, sheet2, etc
6156       xlApp.DisplayAlerts = False
6157       While xlApp.Worksheets.Count > 1
6158           xlApp.Worksheets(1).Delete          'delete the sheet
6159       Wend
6160       xlApp.DisplayAlerts = True

6161       WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
6162       xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

6163       NewWorkBook = WorkSheetName

' <VB WATCH>
6164       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6165       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NewWorkBook"

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
            vbwReportVariable "WorkSheetName", WorkSheetName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Sub FindMagtrols()
' <VB WATCH>
6166       On Error GoTo vbwErrHandler
6167       Const VBWPROCNAME = "frmPLCData.FindMagtrols"
6168       If vbwProtector.vbwTraceProc Then
6169           Dim vbwProtectorParameterString As String
6170           If vbwProtector.vbwTraceParameters Then
6171               vbwProtectorParameterString = "()"
6172           End If
6173           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6174       End If
' </VB WATCH>
6175       Dim I As Integer
6176       Dim j As Integer
6177       Dim rs As New ADODB.Recordset

6178       Do While cmbMagtrol.ListCount > 0
6179           cmbMagtrol.RemoveItem cmbMagtrol.ListCount - 1
6180       Loop

       '==============
6181       Dim sGPIBAddress As String
6182       Dim sGPIBName As String
6183       rs.Open "GPIBAddresses", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

6184       rs.MoveFirst                                'goto the top
6185       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
6186           sGPIBAddress = rs.Fields("IPAddress")        'get the description
6187           sGPIBName = rs.Fields("GPIBName")                      'get the index number - promary key
6188           j = PingSilent(sGPIBAddress)
6189           If j <> 0 Then
                   'also get the type of magtrol (5300 or 6530) from CheckMagtrolModel
6190               sGPIBName = sGPIBName & CheckMagtrolModel(Val(Right(sGPIBName, 1)))
6191               If iberr = 0 Then
6192                   cmbMagtrol.AddItem sGPIBName
6193                   cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = Val(Mid(sGPIBName, 5, 1))
6194               End If
6195           End If
6196           rs.MoveNext                             'get the next record
6197       Next I
6198       rs.Close
6199       Set rs = Nothing

6200       cmbMagtrol.AddItem "Add Manually"
6201       cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
6202       cmbMagtrol.ListIndex = 0

' <VB WATCH>
6203       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6204       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindMagtrols"

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
            vbwReportVariable "j", j
            vbwReportVariable "sGPIBAddress", sGPIBAddress
            vbwReportVariable "sGPIBName", sGPIBName
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Function CheckMagtrolModel(GPIBNo As Integer) As String
' <VB WATCH>
6205       On Error GoTo vbwErrHandler
6206       Const VBWPROCNAME = "frmPLCData.CheckMagtrolModel"
6207       If vbwProtector.vbwTraceProc Then
6208           Dim vbwProtectorParameterString As String
6209           If vbwProtector.vbwTraceParameters Then
6210               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("GPIBNo", GPIBNo) & ") "
6211           End If
6212           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6213       End If
' </VB WATCH>
6214       Dim I As Integer
6215       Dim strRead As String
6216       Dim sSendStr As String
6217       strRead = Space$(182)

           'if we're talking to a magtrol, close the connection
6218       If iUD <> 0 Then
6219           ibonl iUD, 0
       '        UnregisterGPIBGlobals
6220           iUD = 0
6221       End If

           'open a new connection to the magtrol:
               'primary address = 14
               'secondary address = 0
               'timeout = 3 second
               'eoi mode = 1
               'stop reading when line feed character is received - 0x10
               'and return iUD

6222       ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

6223       If iberr Then
6224           I = 0
       '        Debug.Print GPIBNo & " - i=" & iberr
6225           CheckMagtrolModel = ""
6226       Else    'if no error
               'ask who it is
6227           sSendStr = "*IDN?" & vbCrLf
6228           ibwrt iUD, sSendStr

6229           Sleep (1000)

               'see what the Magtrol says
6230           ibrd iUD, strRead
               '6530 will return a string like 6530 R 1.16"
               '5300 will return measurement data

6231           If Left(strRead, 4) = "6530" Then
6232               CheckMagtrolModel = " - 6530"
6233           ElseIf Left(strRead, 2) = "A=" Then
6234               CheckMagtrolModel = " - 5300"
6235           Else
6236               CheckMagtrolModel = " - Unknown"
6237           End If
       '        Debug.Print GPIBNo & " - " & strRead
6238           If iberr Then
       '            Debug.Print iberr
6239           End If
6240       End If
' <VB WATCH>
6241       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6242       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CheckMagtrolModel"

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
            vbwReportVariable "GPIBNo", GPIBNo
            vbwReportVariable "I", I
            vbwReportVariable "strRead", strRead
            vbwReportVariable "sSendStr", sSendStr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Private Sub CalibrateSoftware()
' <VB WATCH>
6243       On Error GoTo vbwErrHandler
6244       Const VBWPROCNAME = "frmPLCData.CalibrateSoftware"
6245       If vbwProtector.vbwTraceProc Then
6246           Dim vbwProtectorParameterString As String
6247           If vbwProtector.vbwTraceParameters Then
6248               vbwProtectorParameterString = "()"
6249           End If
6250           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6251       End If
' </VB WATCH>
6252           frmCalibrate.Show
               'Calibrating = True

' <VB WATCH>
6253       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6254       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalibrateSoftware"

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

Function ParseTEMCModelNo(cmbComboName As ComboBox, ltr As String)
' <VB WATCH>
6255       On Error GoTo vbwErrHandler
6256       Const VBWPROCNAME = "frmPLCData.ParseTEMCModelNo"
6257       If vbwProtector.vbwTraceProc Then
6258           Dim vbwProtectorParameterString As String
6259           If vbwProtector.vbwTraceParameters Then
6260               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
6261               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ltr", ltr) & ") "
6262           End If
6263           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6264       End If
' </VB WATCH>
6265       Dim I As Integer
6266       Dim iStart As Integer
6267       Dim iStop As Integer
6268       Dim strCompare As String

6269       For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
6270           iStart = InStr(1, cmbComboName.List(I), "[")
6271           iStop = InStr(1, cmbComboName.List(I), "]")
6272           strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
6273           If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
6274               cmbComboName.ListIndex = I                                              'if we do, set the combo box
6275               Exit For                                            'and we're done
6276           End If
       '        cmbComboName.ListIndex = -1                             'else, remove any pointer
6277           cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
6278       Next I

6279       txtModelNo.Text = UCase(txtModelNo.Text)
6280       txtModelNo.SelStart = Len(txtModelNo.Text)
' <VB WATCH>
6281       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6282       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ParseTEMCModelNo"

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
            vbwReportVariable "ltr", ltr
            vbwReportVariable "I", I
            vbwReportVariable "iStart", iStart
            vbwReportVariable "iStop", iStop
            vbwReportVariable "strCompare", strCompare
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function LoadCombo(cmbComboName As ComboBox, sTableName As String)
       'load all of the pump parameter combo boxes from the tables on the database
' <VB WATCH>
6283       On Error GoTo vbwErrHandler
6284       Const VBWPROCNAME = "frmPLCData.LoadCombo"
6285       If vbwProtector.vbwTraceProc Then
6286           Dim vbwProtectorParameterString As String
6287           If vbwProtector.vbwTraceParameters Then
6288               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
6289               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ") "
6290           End If
6291           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6292       End If
' </VB WATCH>

6293       Dim I As Integer
6294       Dim sItem As String
6295       Dim iID As Integer
6296       Dim qy As New ADODB.Command
6297       Dim rs As New ADODB.Recordset

6298       qy.ActiveConnection = cnPumpData
6299       If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
6300           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
6301       Else
6302           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
6303       End If
6304       rs.CursorLocation = adUseClient
6305       rs.CursorType = adOpenStatic

6306       rs.Open qy

6307       On Error GoTo NoField

6308       rs.MoveFirst                                'goto the top

6309       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
6310           sItem = rs.Fields("Description")        'get the description
6311           iID = rs.Fields(0)                      'get the index number - promary key
6312           cmbComboName.AddItem sItem, I                                   'add the description to the combo box
6313           cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
6314           rs.MoveNext                             'get the next record
6315       Next I
6316       rs.Close
6317       cmbComboName.ListIndex = -1
6318   On Error GoTo vbwErrHandler
6319       Set rs = Nothing
6320       Set qy = Nothing
' <VB WATCH>
6321       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6322       Exit Function

6323   NoField:
6324   On Error GoTo vbwErrHandler
6325       Resume Next

' <VB WATCH>
6326       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6327       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LoadCombo"

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
            vbwReportVariable "sTableName", sTableName
            vbwReportVariable "I", I
            vbwReportVariable "sItem", sItem
            vbwReportVariable "iID", iID
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function LoadInstrumentationCombo(cmbComboName As ComboBox, sTableName As String)
       'load all of the pump parameter combo boxes from the tables on the database
' <VB WATCH>
6328       On Error GoTo vbwErrHandler
6329       Const VBWPROCNAME = "frmPLCData.LoadInstrumentationCombo"
6330       If vbwProtector.vbwTraceProc Then
6331           Dim vbwProtectorParameterString As String
6332           If vbwProtector.vbwTraceParameters Then
6333               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
6334               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ") "
6335           End If
6336           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6337       End If
' </VB WATCH>

6338       Dim I As Integer
6339       Dim sItem As String
6340       Dim iID As Integer
6341       Dim qy As New ADODB.Command
6342       Dim rs As New ADODB.Recordset


6343       qy.ActiveConnection = cnPumpData
6344       If sTableName = "AnalyzerNo" Then
6345           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE UseInDropdown = true ORDER BY val(Description)"
6346       Else
6347           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE UseInDropdown = true ORDER BY Description"
6348       End If
6349       rs.CursorLocation = adUseClient
6350       rs.CursorType = adOpenStatic

6351       rs.Open qy
6352       Dim j As Integer

6353       On Error GoTo NoField
6354       rs.MoveFirst                                'goto the top
6355       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
6356           sItem = rs.Fields("Description")        'get the description
6357           iID = rs.Fields(0)                      'get the index number - promary key
6358           cmbComboName.AddItem sItem, I                                   'add the description to the combo box
6359           cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
6360           rs.MoveNext                             'get the next record
6361           j = I + 1
6362       Next I
6363       rs.Close

6364       cmbComboName.AddItem "---- Legacy Items Below ---", j
6365       j = j + 1

6366       qy.CommandText = "SELECT * FROM " & sTableName & " WHERE UseInDropdown = false ORDER BY val(Description)"
6367       rs.Open qy

6368       rs.MoveFirst                                'goto the top
6369       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
6370           sItem = rs.Fields("Description")        'get the description
6371           iID = rs.Fields(0)                      'get the index number - promary key
6372           cmbComboName.AddItem sItem, I + j                                   'add the description to the combo box
6373           cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
6374           rs.MoveNext                             'get the next record
6375       Next I
6376       rs.Close

6377       cmbComboName.ListIndex = -1
6378   On Error GoTo vbwErrHandler
6379       Set rs = Nothing
6380       Set qy = Nothing
' <VB WATCH>
6381       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6382       Exit Function

6383   NoField:
       '    bUseDropdown = False
6384   On Error GoTo vbwErrHandler
6385       Resume Next

' <VB WATCH>
6386       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6387       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LoadInstrumentationCombo"

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
            vbwReportVariable "sTableName", sTableName
            vbwReportVariable "I", I
            vbwReportVariable "sItem", sItem
            vbwReportVariable "iID", iID
            vbwReportVariable "j", j
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
'Function SetGraphMax(Plothead) As Integer
'    Dim I As Integer
'    Dim m As Single
'
'    m = 0
'    For I = 0 To UBound(Plothead, 2)
'        If Plothead(1, I) > m Then
'            m = Plothead(1, I)
'        End If
'    Next I
'    SetGraphMax = 10 * (Int((m / 10) + 0.5) + 1)
''    CWGraph2.Axes(2).Maximum = 10 * (Int((m / 10) + 0.5) + 1)
''    SetGraphMax = CWGraph2.Axes(2).Maximum
''    CWGraphTDH.Axes(2).Maximum = CWGraph2.Axes(2).Maximum
''    CWGraphAmps.Axes(2).Maximum = CWGraph2.Axes(2).Maximum
'End Function
Function SetGraphMax(GraphArray) As Integer
' <VB WATCH>
6388       On Error GoTo vbwErrHandler
6389       Const VBWPROCNAME = "frmPLCData.SetGraphMax"
6390       If vbwProtector.vbwTraceProc Then
6391           Dim vbwProtectorParameterString As String
6392           If vbwProtector.vbwTraceParameters Then
6393               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("GraphArray", GraphArray) & ") "
6394           End If
6395           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6396       End If
' </VB WATCH>

6397       Dim I As Integer
6398       Dim m As Single

6399       m = 0
6400       For I = 0 To UBound(GraphArray, 1)
6401           If GraphArray(I, 1) > m Then
6402               m = GraphArray(I, 1)
6403           End If
6404       Next I
6405       SetGraphMax = m

' <VB WATCH>
6406       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6407       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetGraphMax"

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
            vbwReportVariable "GraphArray", GraphArray
            vbwReportVariable "I", I
            vbwReportVariable "m", m
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function CalculateSpeed(CoefSq As Double, CoefLin As Double, CoefConstant As Double, InputHP As Double, SG As Double) As Integer
' <VB WATCH>
6408       On Error GoTo vbwErrHandler
6409       Const VBWPROCNAME = "frmPLCData.CalculateSpeed"
6410       If vbwProtector.vbwTraceProc Then
6411           Dim vbwProtectorParameterString As String
6412           If vbwProtector.vbwTraceParameters Then
6413               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
6414               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
6415               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
6416               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("InputHP", InputHP) & ", "
6417               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SG", SG) & ") "
6418           End If
6419           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6420       End If
' </VB WATCH>
6421       Dim I As Integer
6422       Dim OldResult As Double
6423       Dim NewResult As Double

6424       CalculateSpeed = 0

6425       If SG > 5 Or SG < 0.01 Then
6426           MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
' <VB WATCH>
6427       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6428           Exit Function
6429       End If

6430       OldResult = 1000
6431       NewResult = 0

6432       I = 1

6433       Do While Abs(NewResult - OldResult) > 0.1
6434           ReDim Preserve results(I)
6435           Select Case I
                   Case 1
6436                   results(I - 1).HP = InputHP
6437               Case 2
6438                   results(I - 1).HP = results(I - 2).HP * SG
6439               Case Else
6440                   results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
6441           End Select
6442           OldResult = NewResult
6443           results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
6444           NewResult = results(I - 1).Speed
6445           If I > 15 Then
6446               If I = 0 Or I > 15 Then
6447                   MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
' <VB WATCH>
6448       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6449                   Exit Function
6450               End If
' <VB WATCH>
6451       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
6452               Exit Function
6453           End If
6454           I = I + 1
6455       Loop
6456       CalculateSpeed = I - 1
' <VB WATCH>
6457       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6458       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalculateSpeed"

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
            vbwReportVariable "CoefSq", CoefSq
            vbwReportVariable "CoefLin", CoefLin
            vbwReportVariable "CoefConstant", CoefConstant
            vbwReportVariable "InputHP", InputHP
            vbwReportVariable "SG", SG
            vbwReportVariable "I", I
            vbwReportVariable "OldResult", OldResult
            vbwReportVariable "NewResult", NewResult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function CalcPoly(CoefSq As Double, CoefLin As Double, CoefConstant As Double, DataIn As Double) As Double
' <VB WATCH>
6459       On Error GoTo vbwErrHandler
6460       Const VBWPROCNAME = "frmPLCData.CalcPoly"
6461       If vbwProtector.vbwTraceProc Then
6462           Dim vbwProtectorParameterString As String
6463           If vbwProtector.vbwTraceParameters Then
6464               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
6465               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
6466               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
6467               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("DataIn", DataIn) & ") "
6468           End If
6469           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6470       End If
' </VB WATCH>
6471       CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
' <VB WATCH>
6472       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6473       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalcPoly"

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
            vbwReportVariable "CoefSq", CoefSq
            vbwReportVariable "CoefLin", CoefLin
            vbwReportVariable "CoefConstant", CoefConstant
            vbwReportVariable "DataIn", DataIn
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Sub FixCoef()
       'format the coefficient box
       '
' <VB WATCH>
6474       On Error GoTo vbwErrHandler
6475       Const VBWPROCNAME = "frmPLCData.FixCoef"
6476       If vbwProtector.vbwTraceProc Then
6477           Dim vbwProtectorParameterString As String
6478           If vbwProtector.vbwTraceParameters Then
6479               vbwProtectorParameterString = "()"
6480           End If
6481           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6482       End If
' </VB WATCH>
6483       With xlApp
6484           .Range("C30:D30").Select
6485           With .Selection
6486               .HorizontalAlignment = xlCenter
6487               .VerticalAlignment = xlBottom
6488               .WrapText = False
6489               .Orientation = 0
6490               .AddIndent = False
6491               .IndentLevel = 0
6492               .ShrinkToFit = False
6493               .ReadingOrder = xlContext
6494               .MergeCells = False
6495           End With
6496           .Selection.Merge
6497           .Range("E30:F30").Select
6498           With .Selection
6499               .HorizontalAlignment = xlCenter
6500               .VerticalAlignment = xlBottom
6501               .WrapText = False
6502               .Orientation = 0
6503               .AddIndent = False
6504               .IndentLevel = 0
6505               .ShrinkToFit = False
6506               .ReadingOrder = xlContext
6507               .MergeCells = False
6508           End With
6509           .Selection.Merge
6510           .Range("G30:H30").Select
6511           With .Selection
6512               .HorizontalAlignment = xlCenter
6513               .VerticalAlignment = xlBottom
6514               .WrapText = False
6515               .Orientation = 0
6516               .AddIndent = False
6517               .IndentLevel = 0
6518               .ShrinkToFit = False
6519               .ReadingOrder = xlContext
6520               .MergeCells = False
6521           End With
6522           .Selection.Merge
6523           .Range("C31:D31").Select
6524           With .Selection
6525               .HorizontalAlignment = xlCenter
6526               .VerticalAlignment = xlBottom
6527               .WrapText = False
6528               .Orientation = 0
6529               .AddIndent = False
6530               .IndentLevel = 0
6531               .ShrinkToFit = False
6532               .ReadingOrder = xlContext
6533               .MergeCells = False
6534           End With
6535           .Selection.Merge
6536           .Range("C32:D32").Select
6537           With .Selection
6538               .HorizontalAlignment = xlCenter
6539               .VerticalAlignment = xlBottom
6540               .WrapText = False
6541               .Orientation = 0
6542               .AddIndent = False
6543               .IndentLevel = 0
6544               .ShrinkToFit = False
6545               .ReadingOrder = xlContext
6546               .MergeCells = False
6547           End With
6548           .Selection.Merge
6549           .Range("C33:D33").Select
6550           With .Selection
6551               .HorizontalAlignment = xlCenter
6552               .VerticalAlignment = xlBottom
6553               .WrapText = False
6554               .Orientation = 0
6555               .AddIndent = False
6556               .IndentLevel = 0
6557               .ShrinkToFit = False
6558               .ReadingOrder = xlContext
6559               .MergeCells = False
6560           End With
6561           .Selection.Merge
6562           .Range("E31:F31").Select
6563           With .Selection
6564               .HorizontalAlignment = xlCenter
6565               .VerticalAlignment = xlBottom
6566               .WrapText = False
6567               .Orientation = 0
6568               .AddIndent = False
6569               .IndentLevel = 0
6570               .ShrinkToFit = False
6571               .ReadingOrder = xlContext
6572               .MergeCells = False
6573           End With
6574           .Selection.Merge
6575           .Range("E32:F32").Select
6576           With .Selection
6577               .HorizontalAlignment = xlCenter
6578               .VerticalAlignment = xlBottom
6579               .WrapText = False
6580               .Orientation = 0
6581               .AddIndent = False
6582               .IndentLevel = 0
6583               .ShrinkToFit = False
6584               .ReadingOrder = xlContext
6585               .MergeCells = False
6586           End With
6587           .Selection.Merge
6588           .Range("E33:F33").Select
6589           With .Selection
6590               .HorizontalAlignment = xlCenter
6591               .VerticalAlignment = xlBottom
6592               .WrapText = False
6593               .Orientation = 0
6594               .AddIndent = False
6595               .IndentLevel = 0
6596               .ShrinkToFit = False
6597               .ReadingOrder = xlContext
6598               .MergeCells = False
6599           End With
6600           .Selection.Merge
6601           .Range("G31:H31").Select
6602           With .Selection
6603               .HorizontalAlignment = xlCenter
6604               .VerticalAlignment = xlBottom
6605               .WrapText = False
6606               .Orientation = 0
6607               .AddIndent = False
6608               .IndentLevel = 0
6609               .ShrinkToFit = False
6610               .ReadingOrder = xlContext
6611               .MergeCells = False
6612           End With
6613           .Selection.Merge
6614           .Range("G32:H32").Select
6615           With .Selection
6616               .HorizontalAlignment = xlCenter
6617               .VerticalAlignment = xlBottom
6618               .WrapText = False
6619               .Orientation = 0
6620               .AddIndent = False
6621               .IndentLevel = 0
6622               .ShrinkToFit = False
6623               .ReadingOrder = xlContext
6624               .MergeCells = False
6625           End With
6626           .Selection.Merge
6627           .Range("G33:H33").Select
6628           With .Selection
6629               .HorizontalAlignment = xlCenter
6630               .VerticalAlignment = xlBottom
6631               .WrapText = False
6632               .Orientation = 0
6633               .AddIndent = False
6634               .IndentLevel = 0
6635               .ShrinkToFit = False
6636               .ReadingOrder = xlContext
6637               .MergeCells = False
6638           End With
6639           .Selection.Merge
6640           .Range("B29:H29").Select
6641           With .Selection
6642               .HorizontalAlignment = xlGeneral
6643               .VerticalAlignment = xlBottom
6644               .WrapText = False
6645               .Orientation = 0
6646               .AddIndent = False
6647               .IndentLevel = 0
6648               .ShrinkToFit = False
6649               .ReadingOrder = xlContext
6650               .MergeCells = True
6651           End With
6652           .Selection.UnMerge
6653           With .Selection
6654               .HorizontalAlignment = xlCenter
6655               .VerticalAlignment = xlBottom
6656               .WrapText = False
6657               .Orientation = 0
6658               .AddIndent = False
6659               .IndentLevel = 0
6660               .ShrinkToFit = False
6661               .ReadingOrder = xlContext
6662               .MergeCells = False
6663           End With
6664           .Selection.Merge
6665           .Range("B29:H33").Select
6666           With .Selection.Interior
6667               .ColorIndex = 34
6668               .Pattern = xlSolid
6669           End With
6670           .Range("B29:H33").Select
6671           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6672           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6673           With .Selection.Borders(xlEdgeLeft)
6674               .LineStyle = xlContinuous
6675               .Weight = xlMedium
6676               .ColorIndex = xlAutomatic
6677           End With
6678           With .Selection.Borders(xlEdgeTop)
6679               .LineStyle = xlContinuous
6680               .Weight = xlMedium
6681               .ColorIndex = xlAutomatic
6682           End With
6683           With .Selection.Borders(xlEdgeBottom)
6684               .LineStyle = xlContinuous
6685               .Weight = xlMedium
6686               .ColorIndex = xlAutomatic
6687           End With
6688           With .Selection.Borders(xlEdgeRight)
6689               .LineStyle = xlContinuous
6690               .Weight = xlMedium
6691               .ColorIndex = xlAutomatic
6692           End With
6693           .Selection.Borders(xlInsideVertical).LineStyle = xlNone
6694           .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
6695           .Range("B29:H29").Select
6696           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6697           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6698           With .Selection.Borders(xlEdgeLeft)
6699               .LineStyle = xlContinuous
6700               .Weight = xlMedium
6701               .ColorIndex = xlAutomatic
6702           End With
6703           With .Selection.Borders(xlEdgeTop)
6704               .LineStyle = xlContinuous
6705               .Weight = xlMedium
6706               .ColorIndex = xlAutomatic
6707           End With
6708           With .Selection.Borders(xlEdgeBottom)
6709               .LineStyle = xlContinuous
6710               .Weight = xlMedium
6711               .ColorIndex = xlAutomatic
6712           End With
6713           With .Selection.Borders(xlEdgeRight)
6714               .LineStyle = xlContinuous
6715               .Weight = xlMedium
6716               .ColorIndex = xlAutomatic
6717           End With
6718           .Selection.Borders(xlInsideVertical).LineStyle = xlNone
6719           .Range("B30:H30").Select
6720           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6721           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6722           With .Selection.Borders(xlEdgeLeft)
6723               .LineStyle = xlContinuous
6724               .Weight = xlMedium
6725               .ColorIndex = xlAutomatic
6726           End With
6727           With .Selection.Borders(xlEdgeTop)
6728               .LineStyle = xlContinuous
6729               .Weight = xlMedium
6730               .ColorIndex = xlAutomatic
6731           End With
6732           With .Selection.Borders(xlEdgeBottom)
6733               .LineStyle = xlContinuous
6734               .Weight = xlMedium
6735               .ColorIndex = xlAutomatic
6736           End With
6737           With .Selection.Borders(xlEdgeRight)
6738               .LineStyle = xlContinuous
6739               .Weight = xlMedium
6740               .ColorIndex = xlAutomatic
6741           End With
6742           .Selection.Borders(xlInsideVertical).LineStyle = xlNone
6743           .Range("B30:B33").Select
6744           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6745           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6746           With .Selection.Borders(xlEdgeLeft)
6747               .LineStyle = xlContinuous
6748               .Weight = xlMedium
6749               .ColorIndex = xlAutomatic
6750           End With
6751           With .Selection.Borders(xlEdgeTop)
6752               .LineStyle = xlContinuous
6753               .Weight = xlMedium
6754               .ColorIndex = xlAutomatic
6755           End With
6756           With .Selection.Borders(xlEdgeBottom)
6757               .LineStyle = xlContinuous
6758               .Weight = xlMedium
6759               .ColorIndex = xlAutomatic
6760           End With
6761           With .Selection.Borders(xlEdgeRight)
6762               .LineStyle = xlContinuous
6763               .Weight = xlThin
6764               .ColorIndex = xlAutomatic
6765           End With
6766           .Range("C30:D33").Select
6767           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6768           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6769           With .Selection.Borders(xlEdgeLeft)
6770               .LineStyle = xlContinuous
6771               .Weight = xlThin
6772               .ColorIndex = xlAutomatic
6773           End With
6774           With .Selection.Borders(xlEdgeTop)
6775               .LineStyle = xlContinuous
6776               .Weight = xlMedium
6777               .ColorIndex = xlAutomatic
6778           End With
6779           With .Selection.Borders(xlEdgeBottom)
6780               .LineStyle = xlContinuous
6781               .Weight = xlMedium
6782               .ColorIndex = xlAutomatic
6783           End With
6784           With .Selection.Borders(xlEdgeRight)
6785               .LineStyle = xlContinuous
6786               .Weight = xlThin
6787               .ColorIndex = xlAutomatic
6788           End With
6789           .Selection.Borders(xlInsideVertical).LineStyle = xlNone
6790           .Range("E30:F33").Select
6791           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6792           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6793           With .Selection.Borders(xlEdgeLeft)
6794               .LineStyle = xlContinuous
6795               .Weight = xlThin
6796               .ColorIndex = xlAutomatic
6797           End With
6798           With .Selection.Borders(xlEdgeTop)
6799               .LineStyle = xlContinuous
6800               .Weight = xlMedium
6801               .ColorIndex = xlAutomatic
6802           End With
6803           With .Selection.Borders(xlEdgeBottom)
6804               .LineStyle = xlContinuous
6805               .Weight = xlMedium
6806               .ColorIndex = xlAutomatic
6807           End With
6808           With .Selection.Borders(xlEdgeRight)
6809               .LineStyle = xlContinuous
6810               .Weight = xlThin
6811               .ColorIndex = xlAutomatic
6812           End With
6813           .Selection.Borders(xlInsideVertical).LineStyle = xlNone
6814           .Range("J29").Select
6815       End With
' <VB WATCH>
6816       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6817       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FixCoef"

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
Sub FixFormat()
       '
       '   format the final data
       '
' <VB WATCH>
6818       On Error GoTo vbwErrHandler
6819       Const VBWPROCNAME = "frmPLCData.FixFormat"
6820       If vbwProtector.vbwTraceProc Then
6821           Dim vbwProtectorParameterString As String
6822           If vbwProtector.vbwTraceParameters Then
6823               vbwProtectorParameterString = "()"
6824           End If
6825           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
6826       End If
' </VB WATCH>
6827       With xlApp
6828           .Range("B49:E58").Select
6829           With .Selection.Interior
6830               .ColorIndex = 6
6831               .Pattern = xlSolid
6832           End With
6833           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6834           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6835           With .Selection.Borders(xlEdgeLeft)
6836               .LineStyle = xlContinuous
6837               .Weight = xlMedium
6838               .ColorIndex = xlAutomatic
6839           End With
6840           With .Selection.Borders(xlEdgeTop)
6841               .LineStyle = xlContinuous
6842               .Weight = xlMedium
6843               .ColorIndex = xlAutomatic
6844           End With
6845           With .Selection.Borders(xlEdgeBottom)
6846               .LineStyle = xlContinuous
6847               .Weight = xlMedium
6848               .ColorIndex = xlAutomatic
6849           End With
6850           With .Selection.Borders(xlEdgeRight)
6851               .LineStyle = xlContinuous
6852               .Weight = xlMedium
6853               .ColorIndex = xlAutomatic
6854           End With
6855           .Selection.Borders(xlInsideVertical).LineStyle = xlNone
6856           .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
6857           .Range("B49:E49").Select
6858           .Selection.Merge
6859           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6860           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6861           With .Selection.Borders(xlEdgeLeft)
6862               .LineStyle = xlContinuous
6863               .Weight = xlMedium
6864               .ColorIndex = xlAutomatic
6865           End With
6866           With .Selection.Borders(xlEdgeTop)
6867               .LineStyle = xlContinuous
6868               .Weight = xlMedium
6869               .ColorIndex = xlAutomatic
6870           End With
6871           With .Selection.Borders(xlEdgeBottom)
6872               .LineStyle = xlContinuous
6873               .Weight = xlMedium
6874               .ColorIndex = xlAutomatic
6875           End With
6876           With .Selection.Borders(xlEdgeRight)
6877               .LineStyle = xlContinuous
6878               .Weight = xlMedium
6879               .ColorIndex = xlAutomatic
6880           End With
6881           .Selection.Borders(xlInsideVertical).LineStyle = xlNone
6882           .Range("B50:B58").Select
6883           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6884           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6885           With .Selection.Borders(xlEdgeLeft)
6886               .LineStyle = xlContinuous
6887               .Weight = xlMedium
6888               .ColorIndex = xlAutomatic
6889           End With
6890           With .Selection.Borders(xlEdgeTop)
6891               .LineStyle = xlContinuous
6892               .Weight = xlMedium
6893               .ColorIndex = xlAutomatic
6894           End With
6895           With .Selection.Borders(xlEdgeBottom)
6896               .LineStyle = xlContinuous
6897               .Weight = xlMedium
6898               .ColorIndex = xlAutomatic
6899           End With
6900           With .Selection.Borders(xlEdgeRight)
6901               .LineStyle = xlContinuous
6902               .Weight = xlMedium
6903               .ColorIndex = xlAutomatic
6904           End With
6905           .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
6906           .Range("C50:C58").Select
6907           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6908           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6909           With .Selection.Borders(xlEdgeLeft)
6910               .LineStyle = xlContinuous
6911               .Weight = xlMedium
6912               .ColorIndex = xlAutomatic
6913           End With
6914           With .Selection.Borders(xlEdgeTop)
6915               .LineStyle = xlContinuous
6916               .Weight = xlMedium
6917               .ColorIndex = xlAutomatic
6918           End With
6919           With .Selection.Borders(xlEdgeBottom)
6920               .LineStyle = xlContinuous
6921               .Weight = xlMedium
6922               .ColorIndex = xlAutomatic
6923           End With
6924           With .Selection.Borders(xlEdgeRight)
6925               .LineStyle = xlContinuous
6926               .Weight = xlMedium
6927               .ColorIndex = xlAutomatic
6928           End With
6929           .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
6930           .Range("D50:D58").Select
6931           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6932           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6933           With .Selection.Borders(xlEdgeLeft)
6934               .LineStyle = xlContinuous
6935               .Weight = xlMedium
6936               .ColorIndex = xlAutomatic
6937           End With
6938           With .Selection.Borders(xlEdgeTop)
6939               .LineStyle = xlContinuous
6940               .Weight = xlMedium
6941               .ColorIndex = xlAutomatic
6942           End With
6943           With .Selection.Borders(xlEdgeBottom)
6944               .LineStyle = xlContinuous
6945               .Weight = xlMedium
6946               .ColorIndex = xlAutomatic
6947           End With
6948           With .Selection.Borders(xlEdgeRight)
6949               .LineStyle = xlContinuous
6950               .Weight = xlMedium
6951               .ColorIndex = xlAutomatic
6952           End With
6953           .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
6954           .Range("B50:E50").Select
6955           .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
6956           .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
6957           With .Selection.Borders(xlEdgeLeft)
6958               .LineStyle = xlContinuous
6959               .Weight = xlMedium
6960               .ColorIndex = xlAutomatic
6961           End With
6962           With .Selection.Borders(xlEdgeTop)
6963               .LineStyle = xlContinuous
6964               .Weight = xlMedium
6965               .ColorIndex = xlAutomatic
6966           End With
6967           With .Selection.Borders(xlEdgeBottom)
6968               .LineStyle = xlContinuous
6969               .Weight = xlMedium
6970               .ColorIndex = xlAutomatic
6971           End With
6972           With .Selection.Borders(xlEdgeRight)
6973               .LineStyle = xlContinuous
6974               .Weight = xlMedium
6975               .ColorIndex = xlAutomatic
6976           End With
6977           .Range("B49:E58").Select
6978           .Selection.Font.Bold = True

6979           .Range("B51:E58").Select
6980           .Selection.NumberFormat = "0.00"
6981           .Range("B49:E58").Select
6982           With .Selection
6983               .HorizontalAlignment = xlCenter
6984               .VerticalAlignment = xlBottom
6985               .WrapText = False
6986               .Orientation = 0
6987               .AddIndent = False
6988               .IndentLevel = 0
6989               .ShrinkToFit = False
6990               .ReadingOrder = xlContext
6991               .MergeCells = False
6992           End With
6993           .Range("B49:E49").Select
6994           .Selection.Merge
6995       End With
' <VB WATCH>
6996       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
6997       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FixFormat"

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

Sub GetBalanceHoleData(SerialNumber As String, TestDate As String)
' <VB WATCH>
6998       On Error GoTo vbwErrHandler
6999       Const VBWPROCNAME = "frmPLCData.GetBalanceHoleData"
7000       If vbwProtector.vbwTraceProc Then
7001           Dim vbwProtectorParameterString As String
7002           If vbwProtector.vbwTraceParameters Then
7003               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SerialNumber", SerialNumber) & ", "
7004               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("TestDate", TestDate) & ") "
7005           End If
7006           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
7007       End If
' </VB WATCH>
7008       If rsBalanceHoles.State = adStateOpen Then
7009           rsBalanceHoles.Close
7010       End If
7011       qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
               "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
               "FROM BalanceHoles " & _
               "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
               "ORDER BY [Date], Val([BoltCircle]);"

7012       rsBalanceHoles.Open qyBalanceHoles
7013       rsBalanceHoles.Filter = ""

7014       Set dgBalanceHoles.DataSource = rsBalanceHoles

7015       Dim c As Column
7016       For Each c In dgBalanceHoles.Columns
7017           Select Case c.DataField
               Case "BalanceHoleID"
7018               c.Visible = False
7019           Case "SerialNo"
7020               c.Visible = False
7021           Case "Date"
7022               c.Visible = True
7023               c.Alignment = dbgCenter
7024               c.Width = 2000
7025           Case "Number"
7026               c.Visible = True
7027               c.Alignment = dbgCenter
7028               c.Width = 700
7029           Case "Diameter"
7030               c.Visible = False
7031           Case "Diameter1"
7032               c.Caption = "Diameter"
7033               c.Visible = True
7034               c.Alignment = dbgCenter
7035               c.Width = 700
7036           Case "BoltCircle1"
7037               c.Caption = "Bolt Circle"
7038               c.Visible = True
7039               c.Alignment = dbgCenter
7040               c.Width = 800
7041           Case "BoltCircle"
7042               c.Visible = False
7043           Case Else ' hide all other columns.
7044               c.Visible = False
7045           End Select
7046       Next c

' <VB WATCH>
7047       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
7048       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetBalanceHoleData"

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
            vbwReportVariable "SerialNumber", SerialNumber
            vbwReportVariable "TestDate", TestDate
            vbwReportVariable "c", c
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub ReportToExcel()
' <VB WATCH>
7049       On Error GoTo vbwErrHandler
7050       Const VBWPROCNAME = "frmPLCData.ReportToExcel"
7051       If vbwProtector.vbwTraceProc Then
7052           Dim vbwProtectorParameterString As String
7053           If vbwProtector.vbwTraceParameters Then
7054               vbwProtectorParameterString = "()"
7055           End If
7056           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
7057       End If
' </VB WATCH>

7058       frmReportOptions.Show 1

7059       Dim PosRPM As Integer
7060       Dim PosAxPos As Integer
7061       Dim PosCircFlow As Integer
7062       Dim PosVib As Integer
7063       Dim PosRem As Integer
7064       Dim PosTRG As Integer

7065       PosTRG = frmReportOptions.chkTRG.value * 12
7066       PosRPM = frmReportOptions.chkSelectRPM.value * 12 + frmReportOptions.chkTRG.value
7067       PosAxPos = frmReportOptions.chkSelectAxPos.value * 12 + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value
7068       PosCircFlow = frmReportOptions.chkSelectCircFlow.value * 12 + frmReportOptions.chkSelectAxPos.value + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value
7069       PosVib = frmReportOptions.chkVibration.value * 12 + frmReportOptions.chkSelectCircFlow.value + frmReportOptions.chkSelectAxPos.value + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value
7070       PosRem = 12 + frmReportOptions.chkVibration.value * 2 + frmReportOptions.chkSelectCircFlow.value + frmReportOptions.chkSelectAxPos.value + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value

7071       Dim SaveFileName As String
7072       Dim WorkSheetName As String

7073       Dim I As Integer
7074       Dim iRowNo As Integer
7075       Dim sImp As String
7076       Dim ans As Integer

           'excel
7077       Dim ReportWorkbookName As String
       '    ReportWorkbookName = "C:\Users\MRosenbaum.CHEMPUMP\Desktop\HydraulicTestReportTemplate.xls"
7078       ReportWorkbookName = "\\tei-main-01\F\EN\GROUPS\SHARED\Software\Rundown Test Sheet Templates\HydraulicTestReportTemplate.xls"

7079       Dim SaveReportFileName As String
7080       Dim TemplateWorkSheetName As String
7081       TemplateWorkSheetName = "TestReport"

7082       Dim oXLApp As Excel.Application
7083       Dim oXLBook As Excel.Workbook
7084       Dim oXLNewBook As Excel.Workbook
7085       Dim oXLSheet As Excel.Worksheet
7086       Dim oXLSheetToCopy As Excel.Worksheet

           'open excel
7087       Set oXLApp = New Excel.Application

           'open the template as readonly
7088       Set oXLBook = oXLApp.Workbooks.Open(ReportWorkbookName, ReadOnly:=True)

           'open the report sheet
7089       Set oXLSheet = oXLBook.Worksheets(TemplateWorkSheetName)


7090       oXLApp.Visible = False

           'get the name for the saved report file
7091       CommonDialog1.CancelError = True        'in case the user
7092       On Error GoTo CancelErrHandler                '  chooses the cancel button

7093       CommonDialog1.DialogTitle = "Save Excel Hydraulic Report Files"
7094       CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
7095       CommonDialog1.InitDir = App.Path
7096       CommonDialog1.ShowOpen                     'open the file selection dialog box

7097       If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
7098       Else                                                'the file name already exists
7099           ans = MsgBox(CommonDialog1.filename & " already exists.  Overwrite?", vbYesNo, "File Exists")
7100           If ans = vbYes Then
7101           Else
7102               MsgBox "Exiting routine.  Please reenter and select valid Report File Name", vbOKOnly, "Exiting . . ."
' <VB WATCH>
7103       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
7104               Exit Sub
7105           End If

7106       End If

7107       SaveReportFileName = CommonDialog1.filename

7108       Set oXLNewBook = oXLApp.Workbooks.Add

7109       Set oXLSheetToCopy = oXLSheet
7110       oXLSheetToCopy.Copy oXLNewBook.Sheets(1)

7111       oXLApp.DisplayAlerts = False
7112       While oXLApp.Worksheets.Count > 1
7113           If oXLApp.Worksheets(oXLApp.Worksheets.Count).Name <> oXLSheet.Name Then
7114               oXLApp.Worksheets(oXLApp.Worksheets.Count).Delete           'delete the sheet
7115           End If
7116       Wend
7117       oXLApp.DisplayAlerts = True


7118       Set oXLSheet = Nothing
7119       oXLBook.Close savechanges:=False
7120       Set oXLBook = Nothing
7121       Set oXLSheet = oXLNewBook.Worksheets(TemplateWorkSheetName)

7122       Dim oXLVariant As Variant
7123       oXLVariant = oXLSheet.Range("A1:T44").value



7124       Dim XA(44, 20) As String
7125       Dim ir As Integer
7126       Dim ic As Integer

7127       For ir = 0 To 41
7128           For ic = 0 To 19
7129               XA(ir, ic) = oXLVariant(ir + 1, ic + 1)
7130           Next ic
7131       Next ir

           'write the data to the spreadsheet
7132       With oXLApp

           'write header data
7133           XA(1, 1) = "Run Date:"
7134           XA(1, 3) = CStr(cmbTestDate.List(cmbTestDate.ListIndex))
7135           XA(1, 15) = "Instrumentation / Setup"
7136           XA(2, 1) = "Serial Number:"
7137           XA(2, 3) = txtSN.Text
7138           XA(2, 7) = "Customer:"
7139           XA(2, 9) = Me.txtShpNo.Text
7140           XA(3, 1) = "Model:"
7141           XA(3, 3) = txtModelNo.Text
7142           XA(3, 13) = "Suction:"
7143           XA(3, 15) = txtSuctionID.Text
7144           XA(3, 17) = "Loop:"
7145           XA(3, 19) = cmbLoopNumber.List(cmbLoopNumber.ListIndex)
7146           XA(4, 13) = "Discharge:"
7147           XA(4, 15) = txtDischargeID.Text
7148           XA(4, 17) = "Orifice:"
7149           XA(4, 19) = cmbOrificeNumber.List(cmbOrificeNumber.ListIndex)
7150           XA(5, 1) = "Sales Order:"
7151           XA(5, 3) = txtSalesOrderNumber.Text
7152           XA(5, 5) = "Fluid:"
7153           XA(5, 6) = txtLiquid.Text
7154           XA(5, 9) = "Motor:"
7155           XA(5, 11) = cmbMotor.List(cmbMotor.ListIndex)
7156           XA(5, 13) = "Temperature:"
7157           XA(5, 15) = txtTemperatureID.Text
7158           XA(5, 17) = "Circ Flow:"
7159           XA(5, 19) = txtMagflowID.Text
7160           XA(6, 1) = "RMA:"
7161           XA(6, 3) = txtRMA.Text
7162           XA(6, 5) = "S. G.:"
7163           XA(6, 7) = txtSpGr.Text
7164           XA(6, 9) = "Voltage:"
7165           XA(6, 11) = cmbVoltage.List(cmbVoltage.ListIndex)
7166           XA(6, 13) = "Flow:"
7167           XA(6, 15) = txtFlowmeterID.Text
7168           XA(6, 17) = "PLC:"
7169           XA(6, 19) = cmbPLCNo.List(cmbPLCNo.ListIndex)
7170           XA(7, 5) = "Viscosity (cP):"
7171           XA(7, 7) = txtViscosity.Text
7172           XA(7, 9) = "Frequency (Hz):"
7173           XA(7, 11) = cmbFrequency.List(cmbFrequency.ListIndex)
7174           XA(7, 13) = "Power Analyzer:"
7175           XA(7, 15) = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)
7176           XA(7, 17) = "Tach:"
7177           XA(7, 19) = cmbTachID.List(cmbTachID.ListIndex)
7178           XA(8, 1) = "Test Spec:"
7179           XA(8, 3) = cmbTestSpec.List(cmbTestSpec.ListIndex)
7180           XA(8, 5) = "Temperature:"
7181           XA(8, 7) = Me.txtLiquidTemperature.Text
7182           XA(8, 9) = "Nominal RPM:"
7183           XA(8, 11) = cmbRPM.List(cmbRPM.ListIndex)
7184           XA(9, 13) = "Suction Pipe Dia (in):"
7185           XA(9, 16) = cmbSuctDia.List(cmbSuctDia.ListIndex)
7186           XA(10, 1) = "Design Point"
7187           XA(10, 5) = "Impeller Dia:"
7188           If chkTrimmed.value = 1 Then
7189               If Val(txtImpTrim.Text) <> 0 Then
7190                   XA(10, 7) = txtImpTrim.Text
7191               Else
7192                   XA(10, 7) = txtImpellerDia.Text
7193               End If
7194           Else
7195               XA(10, 7) = txtImpellerDia.Text
7196           End If
7197           XA(10, 9) = "Stator Fill:"
7198           XA(10, 11) = cmbStatorFill.List(cmbStatorFill.ListIndex)
7199           XA(10, 13) = "Suction Gage Height (in):"
7200           XA(10, 16) = txtSuctHeight.Text
7201           XA(11, 1) = "Flow Rate (GPM):"
7202           XA(11, 3) = txtDesignFlow.Text
7203           XA(11, 5) = "Design Pressure (psi):"
7204           XA(11, 7) = cmbDesignPressure.List(cmbDesignPressure.ListIndex)
7205           XA(11, 9) = "Full Load Current (A):"
7206           XA(11, 13) = "Discharge Pipe Dia (in):"
7207           XA(11, 16) = cmbDischDia.List(cmbDischDia.ListIndex)
7208           XA(12, 1) = "TDH (ft):"
7209           XA(12, 3) = txtDesignTDH.Text
7210           XA(12, 5) = "Circulation Path:"
7211           XA(12, 7) = cmbCirculationPath.List(cmbCirculationPath.ListIndex)
7212           XA(12, 9) = "Insulation Class:"
7213           XA(12, 13) = "Discharge Gage Height (in):"
7214           XA(12, 16) = txtDischHeight.Text
7215           XA(16, 1) = "Flow"
7216           XA(16, 2) = "TDH"
7217           XA(16, 3) = "KW"
7218           XA(16, 4) = "Ave"
7219           XA(16, 5) = "Ave"
7220           XA(16, 6) = "Power"
7221           XA(16, 7) = "Overall"
7222           XA(16, 8) = "Suction"
7223           XA(16, 9) = "Disch"
7224           XA(16, 10) = "Suction"
7225           XA(16, 11) = "Vel"

7226           XA(17, 1) = "(GPM)"
7227           XA(17, 2) = "(Ft)"
7228           XA(17, 4) = "Volts"
7229           XA(17, 5) = "Amps"
7230           XA(17, 6) = "Factor"
7231           XA(17, 7) = "Eff"
7232           XA(17, 8) = "Temp(F)"
7233           XA(17, 9) = "Pressure"
7234           XA(17, 10) = "Pressure"
7235           XA(17, 11) = "Head"

7236           XA(18, 11) = "(ft)"

               'variable data from user selection
7237           If PosTRG >= 12 Then
7238               XA(16, PosTRG) = "TRG"
7239               XA(17, PosTRG) = "Position"
7240           End If

7241           If PosRPM >= 12 Then
7242               XA(16, PosRPM) = "RPM"
7243           End If

7244           If PosVib >= 12 Then
7245               XA(16, PosVib) = "Vibration"
7246               XA(17, PosVib) = "Data X"
7247               XA(18, PosVib) = "(in/sec)"
7248               XA(16, PosVib + 1) = "Vibration"
7249               XA(17, PosVib + 1) = "Data Y"
7250               XA(18, PosVib + 1) = "(in/sec)"
7251           End If

7252           If PosAxPos >= 12 Then
7253               XA(16, PosAxPos) = "Axial"
7254               XA(17, PosAxPos) = "Position"
7255               XA(18, PosAxPos) = "(in)"
7256           End If

7257           If PosCircFlow >= 12 Then
7258               XA(16, PosCircFlow) = "Circ Flow"
7259               XA(17, PosCircFlow) = "(GPM)"
7260           End If

7261           XA(16, PosRem) = "Remarks"

7262           Dim j As Integer
7263           rsEff.MoveFirst
7264           For j = 1 To frmPLCData.UpDown2.value
7265               XA(18 + j, 1) = rsEff.Fields("Flow")
7266               XA(18 + j, 2) = Format(rsEff.Fields("TDH"), "##.00")
7267               XA(18 + j, 3) = Format(rsEff.Fields("KW"), "##.00")
7268               XA(18 + j, 4) = Format(rsEff.Fields("Volts"), "##.00")
7269               XA(18 + j, 5) = Format(rsEff.Fields("Amps"), "##.00")
7270               XA(18 + j, 6) = Format(rsEff.Fields("PowerFactor"), "##.00")
7271               XA(18 + j, 7) = Format(rsEff.Fields("OverallEfficiency"), "##.00")
7272               XA(18 + j, 8) = Format(rsEff.Fields("Temperature"), "##.00")
7273               XA(18 + j, 9) = Format(rsEff.Fields("DischPress"), "##.00")
7274               XA(18 + j, 10) = Format(rsEff.Fields("SuctPress"), "##.00")
7275               XA(18 + j, 11) = Format(rsEff.Fields("VelocityHead"), "##.00")

7276               If PosTRG >= 12 And Not IsNull(rsEff.Fields("TEMCTRG")) Then
7277                   XA(18 + j, PosTRG) = rsEff.Fields("TEMCTRG")
7278               End If

7279               If PosRPM >= 12 And Not IsNull(rsEff.Fields("RPM")) Then
7280                   XA(18 + j, PosRPM) = rsEff.Fields("RPM")
7281               End If

7282               If PosVib >= 12 And Not IsNull(rsEff.Fields("VibrationX")) And Not IsNull(rsEff.Fields("VibrationY")) Then
7283                   XA(18 + j, PosVib) = rsEff.Fields("VibrationX")
7284                   XA(18 + j, PosVib + 1) = rsEff.Fields("VibrationY")
7285               End If

7286               If PosAxPos >= 12 And Not IsNull(rsEff.Fields("Pos")) Then
7287                   XA(18 + j, PosAxPos) = rsEff.Fields("Pos")
7288               End If

7289               If PosCircFlow >= 12 And Not IsNull(rsEff.Fields("CircFlow")) Then
7290                   XA(18 + j, PosCircFlow) = rsEff.Fields("CircFlow")
7291               End If

7292               If Not IsNull(rsEff.Fields("Remarks")) Then
7293                   XA(18 + j, PosRem) = rsEff.Fields("Remarks")
7294               End If

7295               rsEff.MoveNext
7296           Next j

7297           XA(28, 3) = "Thrust Balance Settings"
7298           If Me.chkFeathered.value = True Then
7299               XA(29, 10) = "Impeller has been feathered."
7300           Else
7301               XA(29, 10) = ""
7302           End If
7303           XA(30, 10) = "Discharge Orifice Size (in):"
7304           XA(30, 14) = Me.txtOrifice.Text
7305           XA(31, 10) = "Circulation Flow Orifice Size (in):"
7306           XA(31, 14) = Me.txtCircOrifice.Text
7307           XA(38, 1) = "Pump Remarks:"
7308           XA(38, 3) = Me.txtRemarks.Text
7309           XA(40, 1) = "Test Setup Remarks:"
7310           XA(40, 3) = Me.txtTestSetupRemarks.Text
7311           XA(42, 1) = "Other Modifications:"
7312           XA(42, 3) = Me.txtOtherMods.Text

7313           If boGotBalanceHoles Then
7314               If rsBalanceHoles.State = adStateClosed Then
7315                   rsBalanceHoles.ActiveConnection = cnPumpData
7316                   rsBalanceHoles.Open
7317               End If

7318               If rsBalanceHoles.RecordCount <> 0 Then
7319                   rsBalanceHoles.MoveFirst
7320                   For I = 1 To rsBalanceHoles.RecordCount
7321                       XA(29 + I, 1) = rsBalanceHoles.Fields("Date")
7322                       XA(29 + I, 4) = rsBalanceHoles.Fields("Number")
7323                       XA(29 + I, 5) = rsBalanceHoles.Fields("Diameter1")
7324                       XA(29 + I, 6) = rsBalanceHoles.Fields("BoltCircle1")
7325                       rsBalanceHoles.MoveNext
7326                   Next I
7327               Else
7328               End If
7329           End If

7330           XA(29, 7) = "End Play(in):"
7331           XA(29, 9) = Me.txtEndPlay.Text
7332           XA(31, 7) = "G-Gap:"
7333           XA(31, 9) = Me.txtGGap.Text

7334           .Range("A1:T44").value = XA

7335       End With



7336       oXLNewBook.CheckCompatibility = False
7337       oXLNewBook.DoNotPromptForConvert = True
7338       oXLApp.DisplayAlerts = False
7339       oXLNewBook.SaveAs CommonDialog1.filename, FileFormat:=xlWorkbookNormal
7340       oXLNewBook.Close savechanges:=False
7341       oXLApp.DisplayAlerts = True

7342   CancelErrHandler:

        '   oXLApp.Visible = True

7343       Set oXLSheet = Nothing
7344       Set oXLNewBook = Nothing
7345       Set oXLApp = Nothing

        '   oXLApp.Quit
7346   On Error GoTo vbwErrHandler

7347       If CommonDialog1.filename <> "" Then
7348           MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
7349       End If

' <VB WATCH>
7350       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
7351       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ReportToExcel"

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
            vbwReportVariable "PosRPM", PosRPM
            vbwReportVariable "PosAxPos", PosAxPos
            vbwReportVariable "PosCircFlow", PosCircFlow
            vbwReportVariable "PosVib", PosVib
            vbwReportVariable "PosRem", PosRem
            vbwReportVariable "PosTRG", PosTRG
            vbwReportVariable "SaveFileName", SaveFileName
            vbwReportVariable "WorkSheetName", WorkSheetName
            vbwReportVariable "I", I
            vbwReportVariable "iRowNo", iRowNo
            vbwReportVariable "sImp", sImp
            vbwReportVariable "ans", ans
            vbwReportVariable "ReportWorkbookName", ReportWorkbookName
            vbwReportVariable "SaveReportFileName", SaveReportFileName
            vbwReportVariable "TemplateWorkSheetName", TemplateWorkSheetName
            vbwReportVariable "oXLVariant", oXLVariant
            vbwReportVariable "XA", XA
            vbwReportVariable "ir", ir
            vbwReportVariable "ic", ic
            vbwReportVariable "j", j
            vbwReportVariable "oXLApp", oXLApp
            vbwReportVariable "oXLBook", oXLBook
            vbwReportVariable "oXLNewBook", oXLNewBook
            vbwReportVariable "oXLSheet", oXLSheet
            vbwReportVariable "oXLSheetToCopy", oXLSheetToCopy
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
' Procedure added by VB Watch 'ID
Private Sub Form_Initialize() 'ID
    vbwInitializeProtector ' Initialize VB Watch 'ID
End Sub 'ID
' </VB WATCH>
' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
    vbwReportVariable "debugging", debugging
    vbwReportVariable "sDataBaseName", sDataBaseName
    vbwReportVariable "ParentDirectoryName", ParentDirectoryName
    vbwReportVariable "vResponse", vResponse
    vbwReportVariable "sData", sData
    vbwReportVariable "iUD", iUD
    vbwReportVariable "vPlot", vPlot
    vbwReportVariable "boUsingHP", boUsingHP
    vbwReportVariable "boFoundPump", boFoundPump
    vbwReportVariable "boPumpIsApproved", boPumpIsApproved
    vbwReportVariable "boTestDateIsApproved", boTestDateIsApproved
    vbwReportVariable "boFoundTestSetup", boFoundTestSetup
    vbwReportVariable "boFoundTestData", boFoundTestData
    vbwReportVariable "boUsingEpicor", boUsingEpicor
    vbwReportVariable "boPLCOperating", boPLCOperating
    vbwReportVariable "boMagtrolOperating", boMagtrolOperating
    vbwReportVariable "boGotBalanceHoles", boGotBalanceHoles
    vbwReportVariable "FromStoredData", FromStoredData
    vbwReportVariable "HeadFlow", HeadFlow
    vbwReportVariable "EffFlow", EffFlow
    vbwReportVariable "KWFlow", KWFlow
    vbwReportVariable "AmpsFlow", AmpsFlow
    vbwReportVariable "FlowHead", FlowHead
    vbwReportVariable "RatedKW", RatedKW
    vbwReportVariable "blnEnabled", blnEnabled
    vbwReportVariable "EpicorConnectionString", EpicorConnectionString
    vbwReportVariable "rsPumpData", rsPumpData
    vbwReportVariable "rsTestSetup", rsTestSetup
    vbwReportVariable "rsTestData", rsTestData
    vbwReportVariable "rsEff", rsEff
    vbwReportVariable "rsEffDisp", rsEffDisp
    vbwReportVariable "rsBalanceHoles", rsBalanceHoles
    vbwReportVariable "rsPumpParameters", rsPumpParameters
    vbwReportVariable "qyPumpData", qyPumpData
    vbwReportVariable "qyTestSetup", qyTestSetup
    vbwReportVariable "qyBalanceHoles", qyBalanceHoles
    vbwReportVariable "xlApp", xlApp
    vbwReportVariable "xlBook", xlBook
End Sub
' </VB WATCH>
