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
      Tab(1).Control(0)=   "lbltab2(0)"
      Tab(1).Control(1)=   "lbltab2(1)"
      Tab(1).Control(2)=   "lbltab2(65)"
      Tab(1).Control(3)=   "lbltab2(88)"
      Tab(1).Control(4)=   "cmbTestSpec"
      Tab(1).Control(5)=   "cmdEnterTestSetupData"
      Tab(1).Control(6)=   "txtWho"
      Tab(1).Control(7)=   "cmdAddNewTestDate"
      Tab(1).Control(8)=   "txtTestSetupRemarks"
      Tab(1).Control(9)=   "frmInstrumentTags"
      Tab(1).Control(10)=   "frmLoopAndXducer"
      Tab(1).Control(11)=   "frmElecData"
      Tab(1).Control(12)=   "frmThrustBalMods"
      Tab(1).Control(13)=   "frmPerfMods"
      Tab(1).Control(14)=   "frmOtherFiles"
      Tab(1).Control(15)=   "CommonDialog1"
      Tab(1).Control(16)=   "cmdDeleteTestDate"
      Tab(1).Control(17)=   "cmdApproveTestDate"
      Tab(1).Control(18)=   "frmTAndI"
      Tab(1).Control(19)=   "Command1"
      Tab(1).Control(20)=   "txtRMA"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Test Data"
      TabPicture(2)   =   "Main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbltab2(55)"
      Tab(2).Control(1)=   "lbltab2(58)"
      Tab(2).Control(2)=   "lbltab2(59)"
      Tab(2).Control(3)=   "Line1"
      Tab(2).Control(4)=   "Line2"
      Tab(2).Control(5)=   "lbltab2(63)"
      Tab(2).Control(6)=   "lbltab2(64)"
      Tab(2).Control(7)=   "lbltab2(53)"
      Tab(2).Control(8)=   "lbltab2(54)"
      Tab(2).Control(9)=   "shpGetPLCData"
      Tab(2).Control(10)=   "DataGrid1"
      Tab(2).Control(11)=   "cmbPLCLoop"
      Tab(2).Control(12)=   "frmAI"
      Tab(2).Control(13)=   "frmThermocouples"
      Tab(2).Control(14)=   "fmrMiscTestData"
      Tab(2).Control(15)=   "cmdEnterTestData"
      Tab(2).Control(16)=   "DataGrid2"
      Tab(2).Control(17)=   "frmPumpData"
      Tab(2).Control(18)=   "txtNPSHa"
      Tab(2).Control(19)=   "cmdReport"
      Tab(2).Control(20)=   "txtTDH"
      Tab(2).Control(21)=   "Command2"
      Tab(2).Control(22)=   "frmPLCMisc"
      Tab(2).Control(23)=   "frmMagtrol"
      Tab(2).Control(24)=   "UpDown1"
      Tab(2).Control(25)=   "UpDown2"
      Tab(2).Control(26)=   "txtUpDn1"
      Tab(2).Control(27)=   "txtUpDn2"
      Tab(2).Control(28)=   "MSChart1"
      Tab(2).Control(29)=   "frmReport"
      Tab(2).Control(30)=   "MSChart2"
      Tab(2).ControlCount=   31
      TabCaption(3)   =   "Charts"
      TabPicture(3)   =   "Main.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSChart3"
      Tab(3).Control(1)=   "MSChart4"
      Tab(3).Control(2)=   "MSChart5"
      Tab(3).Control(3)=   "MSChart6"
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
         BuddyDispid     =   196640
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
         BuddyDispid     =   196641
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
         Height          =   2055
         Left            =   240
         TabIndex        =   373
         Top             =   2160
         Width           =   14535
         Begin VB.TextBox txtSpHeat 
            Height          =   315
            Left            =   9120
            TabIndex        =   426
            Top             =   960
            Width           =   1335
         End
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
            Top             =   1320
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
            Top             =   1560
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
            Caption         =   "Specific Heat:"
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
            Left            =   7680
            TabIndex        =   427
            Top             =   990
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
            Top             =   1590
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

'v1.3.1 - MHR - 5/2/18
'   calculated rpm = 0 if frequency <>0
'   look for right 3 digits of temc frame number in calculated rpm if normal lookup fails
'   find number of points to plot
'   added SpecificHeat to Epicor retrieval and export to excel
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

Private Sub chkAddedDiodes_Click()
    'if the AddedDiodes box is checked, show the number of diodes
    If chkAddedDiodes.value = 1 Then
        lblNoOfDiodes.Visible = True
        txtNoOfDiodes.Visible = True
    Else
        lblNoOfDiodes.Visible = False
        txtNoOfDiodes.Visible = False
    End If
End Sub

Private Sub chkBalanceHoles_Click()
    'if the balance holes box is checked, show the datagrid
    If chkBalanceHoles.value = 1 Then
        dgBalanceHoles.Visible = True
    Else
        dgBalanceHoles.Visible = False
    End If
    If LenB(frmPLCData.txtSN.Text) = 0 Or LenB(cmbTestDate.Text) = 0 Then
        dgBalanceHoles.Visible = False
    End If
End Sub

Private Sub chkCircOrifice_Click()
        'if the CircOrifice box is checked, show the size
    If chkCircOrifice.value = 1 Then
        lblCircOrifice.Visible = True
        txtCircOrifice.Visible = True
    Else
        lblCircOrifice.Visible = False
        txtCircOrifice.Visible = False
    End If
End Sub


Private Sub chkNPSH_Click()
    'if the NPSH file box is checked, show the file name
    If chkNPSH.value = 1 Then
        txtNPSHFile.Visible = True
    Else
        txtNPSHFile.Visible = False
    End If
End Sub

Private Sub chkOrifice_Click()
    'if the orifice box is checked, show the size
    If chkOrifice.value = 1 Then
        lblOrifice.Visible = True
        txtOrifice.Visible = True
    Else
        lblOrifice.Visible = False
        txtOrifice.Visible = False
    End If
End Sub

Private Sub chkPictures_Click()
    'if the pictures box is checked, show the file name
    If chkPictures.value = 1 Then
        txtPicturesFile.Visible = True
    Else
        txtPicturesFile.Visible = False
    End If
End Sub

Private Sub chkTrimmed_Click()
    'if the trimmed box is checked, show the impeller size
    If chkTrimmed.value = 1 Then
        lblImpTrim.Visible = True
        txtImpTrim.Visible = True
    Else
        lblImpTrim.Visible = False
        txtImpTrim.Visible = False
    End If
End Sub

Private Sub chkVibration_Click()
    'if the vibration box is checked, show the file name
    If chkVibration.value = 1 Then
        txtVibrationFile.Visible = True
    Else
        txtVibrationFile.Visible = False
    End If
End Sub

Private Sub cmbAnalyzerNo_Click()
    Exit Sub
    Dim LI As Integer
    LI = cmbAnalyzerNo.ListIndex

    Dim I As Integer
    Dim SepNo As Integer
    For I = 0 To cmbAnalyzerNo.ListCount - 1
        If Left$(cmbAnalyzerNo.List(I), 4) = "----" Then
            SepNo = I
            Exit For
        End If
    Next
    If FromStoredData = False Then
        If LI >= SepNo Then
            cmbAnalyzerNo.ListIndex = 0
        End If
    End If
End Sub

Private Sub cmbFlowMeter_Click()
    Exit Sub
    Dim LI As Integer
    LI = cmbFlowMeter.ListIndex

    Dim I As Integer
    Dim SepNo As Integer
    For I = 0 To cmbFlowMeter.ListCount - 1
        If Left$(cmbFlowMeter.List(I), 4) = "----" Then
            SepNo = I
            Exit For
        End If
    Next
    If FromStoredData = False Then
        If LI >= SepNo Then
            cmbFlowMeter.ListIndex = 0
        End If
    End If
  
End Sub

Private Sub cmbFrequency_Click()
    If cmbFrequency.Text = "VFD" Then
        txtVFDFreq.Visible = True
        lbltab2(86).Visible = True
    Else
        txtVFDFreq.Visible = False
        lbltab2(86).Visible = False
    End If
End Sub



Private Sub cmbMagtrol_Click()
    Dim I As Integer
    Dim sSendStr As String
    Dim sGPIBName As String

    I = cmbMagtrol.ItemData(cmbMagtrol.ListIndex)
    sGPIBName = "GPIB" & I

    If I = 99 Then      'manual entry
        boMagtrolOperating = False
        EnableMagtrolFields
        Exit Sub
    Else
        boMagtrolOperating = True
    End If


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
    If iUD <> 0 Then
        ibonl iUD, 0
    End If

    'open a new connection to the magtrol:
        'primary address = 14
        'secondary address = 0
        'timeout = 1 second
        'eoi mode = 1
        'stop reading when line feed character is received - 0x10
        'and return iUD

    ibdev I, 14, 0, 11, 1, &H140A, iUD

    If iberr Then   'if we have an error
        I = 0
    Else
        If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
            'tell the magtrol that we want full data
            sSendStr = "FULL" & vbCrLf
            ibwrt iUD, sSendStr
            'tell the magtrol that we don't want to wait for data
            sSendStr = "OPEN" & vbCrLf
            ibwrt iUD, sSendStr
        Else
        End If
    End If
  
End Sub


Private Sub cmbPLCLoop_Click()
    'Change the PLC that we're looking at

    Dim RetVal As String

    'manual data entry selection
    If cmbPLCLoop.ListIndex = cmbPLCLoop.ListCount - 1 Then 'no plc
        boPLCOperating = False
        EnablePLCFields
        If DeviceOpen = True Then
            RetVal = DisconnectPLC()
        End If
        Exit Sub
    End If

    If DeviceOpen = True Then
        RetVal = DisconnectPLC()
    End If

    RetVal = ConnectToPLC(cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex))
    If RetVal <> 0 Then
        MsgBox ("Can't connect to PLC - " & Description(cmbPLCLoop.ListIndex))
        boPLCOperating = False
        EnablePLCFields
    Else
        boPLCOperating = True
        tDevice = cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex)
        DisablePLCFields
    End If
End Sub

Private Sub cmbPLCNo_Click()
    Exit Sub
   Dim LI As Integer
    LI = cmbPLCNo.ListIndex

    Dim I As Integer
    Dim SepNo As Integer
    For I = 0 To cmbPLCNo.ListCount - 1
        If Left$(cmbPLCNo.List(I), 4) = "----" Then
            SepNo = I
            Exit For
        End If
    Next
    If FromStoredData = False Then
        If LI >= SepNo Then
            cmbPLCNo.ListIndex = 0
        End If
    End If
End Sub

Private Sub cmbTachID_Change()
        Exit Sub
    Dim LI As Integer
    LI = cmbTachID.ListIndex

    Dim I As Integer
    Dim SepNo As Integer
    For I = 0 To cmbTachID.ListCount - 1
        If Left$(cmbTachID.List(I), 4) = "----" Then
            SepNo = I
            Exit For
        End If
    Next
    If FromStoredData = False Then
        If LI >= SepNo Then
            cmbTachID.ListIndex = 0
        End If
    End If
End Sub

Private Sub cmbTestDate_Click()
    'select a test date to show

    Dim sName As String
    Dim sParam As String
    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bSk As Boolean
    Dim sBC As Single
    Dim NOK() As Long

    cmdModifyBalanceHoleData.Visible = False


    If Not boFoundTestSetup Then    'if we don't have any TestSetup data written
        boFoundTestData = False
        Exit Sub
    End If


    'select the testsetup data for the serial number
    qyTestSetup.ActiveConnection = cnPumpData
    qyTestSetup.CommandText = "SELECT * " & _
           "From TempTestSetupData " & _
           "Where (((TempTestSetupData.SerialNumber) = '" & txtSN.Text & "') AND TempTestSetupData.Date = #" & cmbTestDate.List(cmbTestDate.ListIndex) & "#) " & _
           "ORDER BY TempTestSetupData.Date;"

    '"SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
           txtSN.Text & "'))ORDER BY TempTestSetupData.Date;"

    If rsTestSetup.State = adStateOpen Then
        rsTestSetup.Close
    End If

    With rsTestSetup     'open the recordset for the query
'        .Index = "FindData"
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open qyTestSetup
    End With

    'move to the selected date
    rsTestSetup.MoveFirst
'
    'show the correct combo box entries for this record
    SetComboTestSetup cmbOrificeNumber, "OrificeNumber", "OrificeNumber", rsTestSetup
    SetComboTestSetup cmbTestSpec, "TestSpec", "TestSpecification", rsTestSetup
    SetComboTestSetup cmbLoopNumber, "LoopNumber", "LoopNumber", rsTestSetup
    SetComboTestSetup cmbSuctDia, "SuctDiam", "SuctionDiameter", rsTestSetup
    SetComboTestSetup cmbDischDia, "DischDiam", "DischargeDiameter", rsTestSetup
    SetComboTestSetup cmbTachID, "TachID", "TachID", rsTestSetup
    SetComboTestSetup cmbAnalyzerNo, "AnalyzerNo", "AnalyzerNo", rsTestSetup
    SetComboTestSetup cmbVoltage, "Voltage", "Voltage", rsTestSetup
    SetComboTestSetup cmbFrequency, "Frequency", "Frequency", rsTestSetup
    SetComboTestSetup cmbMounting, "Mounting", "Mounting", rsTestSetup
    SetComboTestSetup cmbPLCNo, "PLCNo", "PLCNo", rsTestSetup

' use this for flowmeter dropdown
'    SetComboTestSetup cmbFlowMeter, "FlowmeterID", "Flowmeter", rsTestSetup


    'show the correct data in the text boxes
    sName = "FlowmeterID"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString   '""
    End If
    txtFlowmeterID.Text = sParam

    sName = "SuctionID"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtSuctionID.Text = sParam

    sName = "DischID"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtDischargeID.Text = sParam

    sName = "TemperatureID"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtTemperatureID.Text = sParam

    sName = "MagflowID"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtMagflowID.Text = sParam

    sName = "HDCor"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtHDCor.Text = sParam

    sName = "KWMult"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtKWMult.Text = sParam

    sName = "Who"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtWho.Text = sParam

    sName = "RMA"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtRMA.Text = sParam

    sName = "Remarks"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtTestSetupRemarks.Text = sParam

    sName = "VFDFrequency"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtVFDFreq.Text = sParam

    sName = "SuctionGageHeight"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtSuctHeight.Text = sParam

    sName = "DischargeGageHeight"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtDischHeight.Text = sParam

    sName = "EndPlay"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtEndPlay.Text = sParam

    sName = "GGAP"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtGGap.Text = sParam

    sName = "OtherMods"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtOtherMods.Text = sParam

    If rsTestSetup.Fields("ImpFeathered") Then
        chkFeathered.value = 1
    Else
        chkFeathered.value = 0
    End If

    If Val(rsTestSetup.Fields("ImpTrimmed")) = 0 Then
        chkTrimmed.value = 0
        txtImpTrim.Visible = False
        txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
    Else
        chkTrimmed.value = 1
        txtImpTrim.Visible = True
        txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
    End If

    If Val(rsTestSetup.Fields("PumpDischOrifice")) = 0 Then
        chkOrifice.value = 0
        txtOrifice.Visible = False
    Else
        chkOrifice.value = 1
        txtOrifice.Visible = True
        txtOrifice.Text = rsTestSetup.Fields("PumpDischOrifice")
    End If

    If Val(rsTestSetup.Fields("CircFlowOrifice")) = 0 Then
        chkCircOrifice.value = 0
        txtCircOrifice.Visible = False
    Else
        chkCircOrifice.value = 1
        txtCircOrifice.Visible = True
        txtCircOrifice.Text = rsTestSetup.Fields("CircFlowOrifice")
    End If

    If IsNull(rsTestSetup.Fields("NoOfTRGDiodes")) Then
        Me.chkAddedDiodes.value = 0
        Me.txtNoOfDiodes.Visible = False
    Else
        chkAddedDiodes.value = 1
        txtNoOfDiodes.Visible = True
        txtNoOfDiodes.Text = rsTestSetup.Fields("NoOfTRGDiodes")
    End If

     If Not IsNull(rsTestSetup.Fields("NoOfTRGDiodes")) Then
        If Val(rsTestSetup.Fields("NoOfTRGDiodes")) = 0 Then
            Me.chkAddedDiodes.value = 0
            Me.txtNoOfDiodes.Visible = False
        Else
            chkAddedDiodes.value = 1
            txtNoOfDiodes.Visible = True
            txtNoOfDiodes.Text = rsTestSetup.Fields("NoOfTRGDiodes")
        End If
    End If

   If (IsNull(rsTestSetup.Fields("NPSHFile"))) Or (LenB(rsTestSetup.Fields("NPSHFile")) = 0) Then
        chkNPSH.value = 0
        txtNPSHFile.Visible = False
    Else
        chkNPSH.value = 1
        txtNPSHFile.Visible = True
        txtNPSHFile.Text = rsTestSetup.Fields("NPSHFile")
    End If

    If (IsNull(rsTestSetup.Fields("PictureFile"))) Or (LenB(rsTestSetup.Fields("PictureFile")) = 0) Then
        chkPictures.value = 0
        txtPicturesFile.Visible = False
    Else
        chkPictures.value = 1
        txtPicturesFile.Visible = True
        txtPicturesFile.Text = rsTestSetup.Fields("PictureFile")
    End If

    If (IsNull(rsTestSetup.Fields("VibrationFile"))) Or (LenB(rsTestSetup.Fields("VibrationFile")) = 0) Then
        chkVibration.value = 0
        txtVibrationFile.Visible = False
    Else
        chkVibration.value = 1
        txtVibrationFile.Visible = True
        txtVibrationFile.Text = rsTestSetup.Fields("VibrationFile")
    End If


    'for TEMC Inspection Report
    sName = "InsulationMeggerVolts"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(0).Text = sParam

    sName = "InsulationMegOhms"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(1).Text = sParam

    sName = "DielectricVolts"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(2).Text = sParam

    sName = "DielectricTime"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(3).Text = sParam

    sName = "HydrostaticValue"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(4).Text = sParam

    sName = "HydrostaticTime"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(5).Text = sParam

    sName = "PneumaticValue"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(6).Text = sParam

    sName = "PneumaticTime"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(7).Text = sParam

    For I = 0 To cmbTestAndInspection(0).ListCount - 1
        If cmbTestAndInspection(0).Text = rsTestSetup.Fields("HydrostaticUnits") Then
                cmbTestAndInspection(0).ListIndex = I
                Exit For
        End If
        cmbTestAndInspection(0).ListIndex = -1
    Next I


    For I = 0 To cmbTestAndInspection(1).ListCount - 1
        If cmbTestAndInspection(1).Text = rsTestSetup.Fields("PneumaticUnits") Then
                cmbTestAndInspection(1).ListIndex = I
                Exit For
        End If
        cmbTestAndInspection(1).ListIndex = -1
    Next I

    TestAndInspectionGood(0).value = Abs(rsTestSetup!insulationgood)
    TestAndInspectionGood(1).value = Abs(rsTestSetup!DielectricGood)
    TestAndInspectionGood(2).value = Abs(rsTestSetup!HydrostaticGood)
    TestAndInspectionGood(3).value = Abs(rsTestSetup!PneumaticGood)
    TestAndInspectionGood(4).value = Abs(rsTestSetup!GeneralAppearanceGood)
    TestAndInspectionGood(5).value = Abs(rsTestSetup!OutlineDimensionsGood)
    TestAndInspectionGood(6).value = Abs(rsTestSetup!MotorNoLoadTestGood)
    TestAndInspectionGood(7).value = Abs(rsTestSetup!MotorLockedRotorTestGood)
    TestAndInspectionGood(8).value = Abs(rsTestSetup!HydrostaticTestGood)
    TestAndInspectionGood(9).value = Abs(rsTestSetup!HydraulicTestGood)
    TestAndInspectionGood(10).value = Abs(rsTestSetup!NPSHTestGood)
    TestAndInspectionGood(11).value = Abs(rsTestSetup!CleanPurgeSealGood)
    TestAndInspectionGood(12).value = Abs(rsTestSetup!PaintCheckGood)
    TestAndInspectionGood(13).value = Abs(rsTestSetup!NameplateGood)
    TestAndInspectionGood(14).value = Abs(rsTestSetup!SupervisorApproval)

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
    GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

    If rsBalanceHoles.RecordCount = 0 Then
        chkBalanceHoles.value = 0
        dgBalanceHoles.Visible = False
        boGotBalanceHoles = False
    Else
        boGotBalanceHoles = True
        ReDim NOK(rsBalanceHoles.RecordCount)
        rsBalanceHoles.MoveLast
        For I = 1 To rsBalanceHoles.RecordCount
            NOK(I) = 0
        Next I

        For j = 1 To rsBalanceHoles.RecordCount - 1
            rsBalanceHoles.MoveFirst
            rsBalanceHoles.Move rsBalanceHoles.RecordCount - j
            sBC = rsBalanceHoles.Fields("BoltCircle")
            bSk = False
            For k = 1 To rsBalanceHoles.RecordCount
                If NOK(k) = rsBalanceHoles.Fields(0) Then
                    bSk = True
                End If
            Next k
            If Not bSk Then
                For I = rsBalanceHoles.RecordCount - j To 1 Step -1
                    rsBalanceHoles.MovePrevious
                    If rsBalanceHoles.Fields("BoltCircle") = sBC Then
                        NOK(I) = rsBalanceHoles.Fields(0)
                    End If
                Next I
            End If
        Next j

        Dim sFilt
        sFilt = ""
        For I = 1 To rsBalanceHoles.RecordCount
            If NOK(I) <> 0 Then
                sFilt = sFilt & "(BalanceHoleID <> " & NOK(I) & ") AND "
'                sFilt = sFilt & "(" & rsBalanceHoles.Filter & " AND BalanceHoleID <> " & NOK(I) & ") AND "
            End If
        Next I

        If Len(sFilt) > 4 Then
            sFilt = Left(sFilt, Len(sFilt) - 4)
            rsBalanceHoles.Filter = sFilt
        End If

        chkBalanceHoles.value = 1
        dgBalanceHoles.Visible = True
'        Set dgBalanceHoles.DataSource = rsBalanceHoles
    End If
'
    'set the test date filter for the test data
    rsTestData.Filter = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"

    If rsTestData.RecordCount = 0 Then
        boFoundTestData = False
        AddTestData
        EnableTestDataControls
        MsgBox "No Test Data Exists for this Serial Number"
    Else
        boFoundTestData = True
        DisableTestDataControls                         'if it's in the real database, don't allow changes here
    End If

    If Not boTestDateIsApproved Then    'data approved?
        EnableTestDataControls
    End If

    If rsTestSetup.Fields("Approved") = True Then
        DisableTestDataControls                         'if it's in the real database, don't allow changes here
        lblTestDateApproved.Visible = True
        MsgBox ("Found pump.  Data cannot be modified.")
        If boCanApprove Then
            cmdApproveTestDate.Caption = "Unapprove this Test Date"
        End If
    Else
        EnableTestDataControls                          'it's in the temp database, allow changes
        lblTestDateApproved.Visible = False
        If boPumpIsApproved = True Then
            MsgBox ("Found pump.  Pump data cannot be modified, but test setup data and test data can be modified.")
        Else
            MsgBox ("Found pump.  Pump data, test setup data, and test data can be modified.")
        End If
        If boCanApprove Then
            If rsPumpData.Fields("Approved") = True Then
                cmdApproveTestDate.Enabled = True
                cmdApproveTestDate.Caption = "Approve this Test Date"
            Else
                cmdApproveTestDate.Caption = "You Must Approve Pump First"
                cmdApproveTestDate.Enabled = False
            End If
        End If
    End If

    rsEff.MoveFirst
    rsTestData.MoveFirst

    For I = 1 To rsTestData.RecordCount
        DoEfficiencyCalcs
        rsEff.MoveNext
        rsTestData.MoveNext
    Next I

    'get a recordset to display
    If rsEffDisp.State = adStateOpen Then
        rsEffDisp.Close
    End If

    Dim qyEffDisp As New ADODB.Command
    qyEffDisp.ActiveConnection = cnEffData
    qyEffDisp.CommandText = "SELECT Flow, TDH, KW, Volts, Amps, OverallEfficiency FROM Efficiency;"

    With rsEffDisp     'open the recordset for the query
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open qyEffDisp
    End With


   ' fix the datagrid
   Set DataGrid1.DataSource = rsTestData
   Set DataGrid2.DataSource = rsEffDisp

   Dim c As Column
   For Each c In DataGrid1.Columns
      Select Case c.DataField
      Case "TestDataID"     'Hide some columns
         c.Visible = False
      Case "SerialNumber"
         c.Visible = False
      Case "Date"
         c.Visible = False
      Case Else             ' Show all other columns.
         c.Visible = True
         c.Alignment = dbgRight
      End Select
    Next c

'DataGrid2.Columns(0).NumberFormat = "###0.00"
'DataGrid2.Columns(1).NumberFormat = "N2"
'DataGrid2.Columns(2).NumberFormat = "N2"
'DataGrid2.Columns(3).NumberFormat = "N2"
'DataGrid2.Columns(4).NumberFormat = "N2"
'DataGrid2.Columns(5).NumberFormat = "N2"

    For Each c In DataGrid2.Columns
        c.Alignment = dbgCenter
        c.Width = 750
        c.NumberFormat = "###0.00"
        Select Case c.ColIndex
            Case 0
                c.Caption = "Flow"
                c.NumberFormat = "###0.00"
            Case 1
                c.Caption = "TDH"
                c.NumberFormat = "##0.00"
            Case 2
                c.Caption = "Input Pwr"
                c.NumberFormat = "##0.00"
                c.Width = 850
                c.Visible = True
            Case 3
                c.Caption = "Voltage"
                c.NumberFormat = "##0.00"
            Case 4
                c.Caption = "Current"
                c.NumberFormat = "##0.00"
            Case 5
                c.Caption = "Overall Eff"
                c.NumberFormat = "##0.00"
                c.Width = 850
'            Case 7
'                c.Caption = "NPSHr"
'                c.NumberFormat = "#0.00"
            Case Else
                'c.Visible = False
        End Select
    Next c
    FixPointsToPlot

    txtUpDn1.Text = 1
'    UpDown2.value = rsTestData.RecordCount

'unlock the text boxes
    For I = 0 To 7
        txtTitle(I).Locked = False
    Next I

    For I = 20 To 27
        txtTitle(I).Locked = False
    Next I

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
    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

    qy.ActiveConnection = cnPumpData

    'see if we have an entry in the table
    qy.CommandText = "SELECT * FROM AITitles " & _
               "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
               "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#)); "

    With rs     'open the recordset for the query
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open qy
    End With

    If Not (rs.BOF = True And rs.EOF = True) Then   'update titles
        rs.MoveFirst
        Do While Not rs.EOF
            txtTitle(rs.Fields("Channel")).Text = rs.Fields("Title")
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set qy = Nothing
End Sub

Private Sub cmdAddNewBalanceHoles_Click()
    Dim strInput As String
    Dim I As Integer
    Dim sNumber As Integer
    Dim sDia As Single
    Dim sBC As Single

    'get the data for the balance holes
    strInput = InputBox("Enter Number of Holes")
    If strInput <> "" Then
        sNumber = CInt(strInput)
    Else
        GoTo CancelPressed
    End If

    strInput = InputBox("Enter Decimal Value of Hole Diameter or Slot (For Example, 0.675) ")
    If strInput <> "" Then
        If UCase(strInput) = "SLOT" Then
            strInput = 99
        End If
        sDia = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    strInput = InputBox("Enter Decimal Value of Bolt Circle or Unknown (For Example, 4.525)")
    If strInput <> "" Then
        If UCase(strInput) = "UNKNOWN" Then
            strInput = 99
        End If
        sBC = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    If rsBalanceHoles.State <> adStateOpen Then
        rsBalanceHoles.Open
    End If

    rsBalanceHoles.AddNew
    rsBalanceHoles!SerialNo = txtSN.Text
    rsBalanceHoles!Date = cmbTestDate.Text
    rsBalanceHoles!Number = sNumber
    rsBalanceHoles!diameter = sDia
    rsBalanceHoles!boltcircle = sBC

    rsBalanceHoles.Update
'    rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"
'    rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "' AND Date <= #" & cmbTestDate.Text & "#"

    GetBalanceHoleData txtSN.Text, cmbTestDate.Text
'    rsBalanceHoles.Requery
    rsBalanceHoles.MoveLast
    dgBalanceHoles.Refresh
    chkBalanceHoles.value = 1


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
    Exit Sub

CancelPressed:
    MsgBox "No New Balance Hole Data Entered", vbOKOnly
End Sub

Private Sub cmdAddNewTestDate_Click()
    'add a new test date/time
    Dim I As Integer

    For I = 1 To cmbTestDate.ListCount      'see if we already have today's date entered
        If cmbTestDate.List(I) = Date Then
            MsgBox "There is already an entry for today.  You can only have one entry for each Serial Number and a given date.  You may want to modify the Serial Number.", vbOKOnly
            Exit Sub
        End If
    Next I

    'we didn't find today's date entered, allow data entry
    boFoundTestSetup = False

    EnableTestSetupDataControls
    cmdEnterTestSetupData_Click
    cmdAddNewBalanceHoles.Visible = True
    txtWho.Text = LogInInitials
    MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
End Sub

Private Sub cmdApprovePump_Click()
    'allow the pump data to be approved
    rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
    rsPumpData.Update
    rsPumpData.Requery
    lblPumpApproved.Visible = rsPumpData.Fields("Approved")
    If rsPumpData.Fields("Approved") = True Then
        cmdApprovePump.Caption = "Unapprove This Pump"
        cmdApproveTestDate.Enabled = True
        If rsTestSetup.Fields("Approved") = True Then
            cmdApproveTestDate.Caption = "Unapprove This Test Date"
        Else
            cmdApproveTestDate.Caption = "Approve This Test Date"
        End If
    Else
        cmdApprovePump.Caption = "Approve This Pump"
        cmdApproveTestDate.Caption = "You Must Approve Pump First"
        cmdApproveTestDate.Enabled = False
    End If
End Sub

Private Sub cmdApproveTestDate_Click()
    'allow the test setup data to be approved
    rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
    rsTestSetup.Update
    rsTestSetup.Requery
    lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
    If rsTestSetup.Fields("Approved") = True Then
        cmdApproveTestDate.Caption = "Unapprove This Test Date"
    Else
        cmdApproveTestDate.Caption = "Approve This Test Date"
    End If
End Sub

Private Sub cmdCalibrate_Click()
    Dim ans As Integer
    Dim I As Integer

    ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
    If ans = vbNo Then
        Calibrating = False
        Exit Sub
    Else
        CalibrateSoftware
    End If
End Sub

Private Sub cmdClearPumpData_Click()
    BlankData
End Sub

Private Sub cmdDeletePump_Click()
    'delete this pump
    Dim Answer As Integer
    Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
    If Answer = vbYes Then
        rsPumpData.Delete
        rsPumpData.Update
        cmdFindPump_Click
    End If
End Sub

Private Sub cmdDeleteTestDate_Click()
    'delete this test date
    Dim Answer As Integer
    Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
    If Answer = vbYes Then
        rsTestSetup.Delete
        rsTestSetup.Update
        cmdFindPump_Click
    End If
End Sub

Private Sub cmdEnterPumpData_Click()
    'store the data on the screen to the pump (pumpdata)
    Dim d As Integer
    Dim sSearch As String
    Dim ans As Integer
    Dim boWriteDataWritten As Boolean


    'check for a serial number
    If LenB(txtSN.Text) = 0 Then
        MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
        Exit Sub
    End If

    'check to make sure most entries are filled in
    If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
        MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If
    If LenB(txtSalesOrderNumber.Text) = 0 Then
        If InStr(1, txtSN.Text, "-") <> 0 Then
            txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
        End If
    End If
    If LenB(txtSalesOrderNumber.Text) = 0 Then
        MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
        cmbStatorFill.ListIndex = 0
    End If

    If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If


    If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

'check TEMC dropdowns

    If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
        MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If


    If Not boFoundPump Then     'if we havent found a pump in the database, add it
        rsPumpData.AddNew
        boWriteDataWritten = False
    Else    'else, find the entry
        sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
        rsPumpData.MoveFirst
        rsPumpData.Find sSearch, , adSearchForward, 1
        boWriteDataWritten = True
    End If

    If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
        ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
        If ans = vbNo Then
            rsPumpData!DataWritten = True
            rsPumpData.Update   'update datawritten
            Exit Sub
        End If
    End If

    rsPumpData!SerialNumber = frmPLCData.txtSN.Text
    If LenB(frmPLCData.txtModelNo.Text) <> 0 Then
        rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
    End If
    rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text

    If LenB(frmPLCData.txtShpNo.Text) <> 0 Then
        rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
    End If

    If LenB(frmPLCData.txtBilNo.Text) <> 0 Then
        rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
    End If

    rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
    If Len(frmPLCData.txtViscosity) <> 0 Then
        rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
    End If

    If LenB(txtSpGr.Text) <> 0 Then
        If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
            MsgBox "Specific Gravity must be a number."
            Exit Sub
        End If
        rsPumpData!SpGr = frmPLCData.txtSpGr.Text
    End If
    If LenB(txtImpellerDia.Text) <> 0 Then
        If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
            MsgBox "Impeller Diameter must be a number."
            Exit Sub
        End If
        rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
    End If

    If LenB(txtLiquid.Text) <> 0 Then
        rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
    End If
    If LenB(txtDesignFlow.Text) <> 0 Then
        rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
    End If
    If LenB(txtDesignTDH.Text) <> 0 Then
        rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
    End If
    If LenB(txtRemarks.Text) <> 0 Then
        rsPumpData!Remarks = txtRemarks.Text
    End If

    If optMfr(0).value = True Then
        d = cmbMotor.ItemData(cmbMotor.ListIndex)
        rsPumpData!Motor = d
        d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
        rsPumpData!StatorFill = d
         d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
        rsPumpData!DesignPressure = d
        d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
        rsPumpData!CirculationPath = d
        d = cmbRPM.ItemData(cmbRPM.ListIndex)
        rsPumpData!RPM = d
        d = cmbModel.ItemData(cmbModel.ListIndex)
        rsPumpData!Model = d
        d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
        rsPumpData!ModelGroup = d
    End If
'   TEMC fields
    If optMfr(0).value = False Then
        d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
        rsPumpData!TEMCAdapter = d

        d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
        rsPumpData!TEMCAdditions = d

        d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
        rsPumpData!TEMCcirculation = d

        d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
        rsPumpData!TEMCDesignpressure = d

        d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
        rsPumpData!TEMCDivisionType = d

        d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
        rsPumpData!TEMCImpellerType = d

        d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
        rsPumpData!TEMCInsulation = d

        d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
        rsPumpData!TEMCJacketGasket = d

        d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
        rsPumpData!TEMCMaterials = d

        d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
        rsPumpData!TEMCModel = d

        d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
        rsPumpData!TEMCNominalImpSize = d

        d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
        rsPumpData!TEMCNominalDischargeSize = d

        d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
        rsPumpData!TEMCNominalSuctionSize = d

        d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
        rsPumpData!TEMCOtherMotor = d

        d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
        rsPumpData!TEMCPumpStages = d

        d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
        rsPumpData!TEMCTRG = d

        d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
        rsPumpData!TEMCVoltage = d

        If LenB(txtTEMCFrameNumber.Text) <> 0 Then
            rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
        End If
    End If

    rsPumpData!ChempumpPump = optMfr(0).value

    rsPumpData!Approved = False

'added from TEMC Inspection Report
    If Len(txtJobNum.Text) <> 0 Then
        rsPumpData!JobNumber = txtJobNum.Text
    End If

    If Len(txtNoPhases.Text) <> 0 Then
        rsPumpData!Phases = txtNoPhases.Text
    End If

    If Len(txtExpClass.Text) <> 0 Then
        rsPumpData!ExpClass = txtExpClass.Text
    End If

    If Len(txtThermalClass.Text) <> 0 Then
        rsPumpData!ThermalClass = txtThermalClass.Text
    End If

    If LenB(txtNPSHr.Text) <> 0 Then
        rsPumpData!NPSHr = Val(txtNPSHr.Text)
    End If

    If LenB(txtLiquidTemperature.Text) <> 0 Then
        rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
    End If

    If LenB(txtRatedInputPower.Text) <> 0 Then
        rsPumpData!RatedOutput = Val(txtRatedInputPower.Text)
    End If

    If LenB(txtAmps.Text) <> 0 Then
        rsPumpData!FLCurrent = Val(txtAmps.Text)
    End If



    If boWriteDataWritten Then
        rsPumpData!DataWritten = True
    Else
        rsPumpData!DataWritten = False
    End If

    'write the data into the database
    rsPumpData.Update
    boFoundPump = True

    'enter a new test date if it's a new entry
    If Not boWriteDataWritten Then


        cmdAddNewTestDate_Click
    End If
End Sub
Private Sub cmdEnterTestData_Click()
    ' save the data on the screen to test data at the selected run
    Dim sSearch As String
    Dim ans As Integer

    'if we didn't find the test setup, can't enter test data
    If Not boFoundTestSetup Then
        MsgBox "You must enter Test Setup Data before entering the Test Data"
        Exit Sub
    End If

    'if we don't find data in the test database, add records
    If boFoundTestData = False Then     'add 8 records for 8 tests
        AddTestData
        rsTestData.MoveFirst
    Else        'find the data in the database
        sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
        rsTestData.MoveFirst
        rsTestData.Filter = sSearch
    End If

    'find the desired record from the form
    rsTestData.MoveFirst
    rsTestData.Move UpDown1.value - 1

    If rsTestData!DataWritten = True Then
        ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
        If ans = vbNo Then
            Exit Sub
        End If
    End If

    rsEff.MoveFirst
    rsEff.Move UpDown1.value - 1

    If LenB(txtV1.Text) <> 0 Then
        rsTestData!VoltageA = Val(txtV1.Text)
    End If

    If LenB(txtV2.Text) <> 0 Then
        rsTestData!VoltageB = Val(txtV2.Text)
    End If

    If LenB(txtV3.Text) <> 0 Then
        rsTestData!VoltageC = Val(txtV3.Text)
    End If

    If LenB(txtI1.Text) <> 0 Then
        rsTestData!CurrentA = Val(txtI1.Text)
    End If

    If LenB(txtI2.Text) <> 0 Then
        rsTestData!CurrentB = Val(txtI2.Text)
    End If

    If LenB(txtI3.Text) <> 0 Then
        rsTestData!CurrentC = Val(txtI3.Text)
    End If

    If LenB(txtP1.Text) <> 0 Then
        rsTestData!PowerA = Val(txtP1.Text)
    End If

    If LenB(txtP2.Text) <> 0 Then
        rsTestData!PowerB = Val(txtP2.Text)
    End If

    If LenB(txtP3.Text) <> 0 Then
        rsTestData!PowerC = Val(txtP3.Text)
    End If

    If LenB(txtKW.Text) <> 0 Then
        rsTestData!TotalPower = Val(txtKW.Text)
    End If

    rsTestData!Flow = Val(txtFlowDisplay.Text)
    rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
    rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
    rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

    rsTestData!TC1 = Val(txtTC1Display.Text)
    rsTestData!TC2 = Val(txtTC2Display.Text)
    rsTestData!TC3 = Val(txtTC3Display.Text)
    rsTestData!TC4 = Val(txtTC4Display.Text)

    rsTestData!CircFlow = Val(txtAI1Display.Text)
    rsTestData!RBHTemp = Val(txtAI2Display.Text)
    rsTestData!RBHPress = Val(txtAI3Display.Text)
    rsTestData!AI4 = Val(txtAI4Display.Text)

    rsTestData!ValvePosition = Val(txtValvePosition.Text)
    rsTestData!SetPoint = Val(txtSetPoint.Text)

    If LenB(txtThrustBal.Text) <> 0 Then
        rsTestData!ThrustBalance = txtThrustBal.Text
    End If

    If LenB(txtVibAx.Text) <> 0 Then
        rsTestData!VibrationX = txtVibAx.Text
    End If

    If LenB(txtVibRad.Text) <> 0 Then
        rsTestData!VibrationY = txtVibRad.Text
    End If

    If LenB(txtTEMCTRGReading.Text) <> 0 Then
        rsTestData!TEMCTRG = txtTEMCTRGReading.Text
    Else
        rsTestData!TEMCTRG = 0
    End If

    If LenB(txtRPM.Text) <> 0 Then
        rsTestData!RPM = txtRPM.Text
    End If

    If LenB(txtTestRemarks.Text) <> 0 Then
        rsTestData!Remarks = txtTestRemarks.Text
    Else
        rsTestData!Remarks = " "
    End If

    If LenB(txtTEMCTRGReading.Text) <> 0 Then
        rsTestData!TEMCTRG = txtTEMCTRGReading.Text
    End If

    If LenB(txtTEMCFrontThrust.Text) <> 0 Then
        rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
    End If

    If LenB(txtTEMCRearThrust.Text) <> 0 Then
        rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
    End If

    If LenB(txtTEMCMomentArm.Text) <> 0 Then
        rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
    End If

    If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
        rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
    End If

    If LenB(txtTEMCViscosity.Text) <> 0 Then
        rsTestData!TEMCViscosity = txtTEMCViscosity.Text
    End If

    If LenB(txtNPSHa.Text) <> 0 Then
        rsTestData!NPSHa = txtNPSHa.Text
    End If

    rsTestData!Approved = False

    rsTestData!DataWritten = True

    'update the database
    rsTestData.Update

    DoEfficiencyCalcs
    rsEff.Update

    'update the form
    DataGrid1.Refresh
    DataGrid2.Refresh

    FixPointsToPlot
  
End Sub
Private Sub cmdEnterTestSetupData_Click()
    'save the data on the screen to testsetupdata
    Dim I As Integer
    Dim d As Integer
    Dim sSearch As String
    Dim ans As Integer
    Dim boWriteDataWritten As Boolean

    'check for a serial number
    If LenB(txtSN.Text) = 0 Then
        MsgBox "You must have a Serial Number to enter data."
        Exit Sub
    End If

    If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
        rsTestSetup.AddNew
        cmbTestDate.AddItem Now
        cmbTestDate.ListIndex = cmbTestDate.NewIndex
        cmdAddNewBalanceHoles.Visible = True
        boFoundTestSetup = True
        boWriteDataWritten = False
        rsTestSetup!DataWritten = False
    Else    'find the record and display
        sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
        rsTestSetup.MoveFirst
        rsTestSetup.Filter = sSearch
        If Not boCanApprove Then
'            cmdAddNewBalanceHoles.Visible = False
        End If
        boWriteDataWritten = True
    End If

    If rsTestSetup!DataWritten = True Then
        ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
        If ans = vbNo Then
            Exit Sub
        End If
    End If

    rsTestSetup!SerialNumber = txtSN
    rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

    If LenB(txtFlowmeterID.Text) <> 0 Then
        rsTestSetup!flowmeterid = txtFlowmeterID
    Else
        rsTestSetup!flowmeterid = vbNullString
    End If

'    I = cmbFlowMeter.ListIndex
'    If I = -1 Then
'        d = -1
'    Else
'        d = cmbFlowMeter.ItemData(I)
'        rsTestSetup!flowmeterid = d
'    End If



    If LenB(txtSuctionID.Text) <> 0 Then
        rsTestSetup!suctionid = txtSuctionID
    Else
        rsTestSetup!suctionid = vbNullString
    End If
    If LenB(txtDischargeID.Text) <> 0 Then
        rsTestSetup!dischid = txtDischargeID
    Else
        rsTestSetup!dischid = vbNullString
    End If
    If LenB(txtTemperatureID.Text) <> 0 Then
        rsTestSetup!temperatureid = txtTemperatureID
    Else
        rsTestSetup!temperatureid = vbNullString
    End If
    If LenB(txtMagflowID.Text) <> 0 Then
        rsTestSetup!magflowid = txtMagflowID
    Else
        rsTestSetup!magflowid = vbNullString
    End If
    If LenB(txtHDCor.Text) <> 0 Then
        rsTestSetup!HDCor = txtHDCor
    Else
        rsTestSetup!HDCor = 0
    End If
    If LenB(txtKWMult.Text) <> 0 Then
        rsTestSetup!kwmult = txtKWMult
    Else
        rsTestSetup!kwmult = 1
    End If
    If LenB(txtWho.Text) <> 0 Then
        rsTestSetup!who = txtWho
    Else
        rsTestSetup!who = vbNullString
    End If
    If LenB(txtRMA.Text) <> 0 Then
        rsTestSetup!RMA = txtRMA
    Else
        rsTestSetup!RMA = vbNullString
    End If
    If LenB(frmPLCData.txtDischHeight) <> 0 Then
        rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
    Else
        rsTestSetup!DischargeGageHeight = 0
    End If
    If LenB(frmPLCData.txtSuctHeight) <> 0 Then
        rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
    Else
        rsTestSetup!SuctionGageHeight = 0
    End If
    If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
        rsTestSetup!Remarks = txtTestSetupRemarks.Text
    Else
        rsTestSetup!Remarks = vbNullString
    End If
    If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
        rsTestSetup!VFDFrequency = txtVFDFreq.Text
    Else
        rsTestSetup!VFDFrequency = 0
    End If

    I = cmbOrificeNumber.ListIndex
    If I = -1 Then
        d = 18      'entry for None
    Else
        d = cmbOrificeNumber.ItemData(I)
    End If
    rsTestSetup!orificenumber = d

    If LenB(txtEndPlay.Text) <> 0 Then
        rsTestSetup!EndPlay = Val(frmPLCData.txtEndPlay.Text)
    Else
        rsTestSetup!EndPlay = 0
    End If

    If LenB(txtGGap.Text) <> 0 Then
        rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
    Else
        rsTestSetup!GGAP = 0
    End If

    If LenB(txtOtherMods.Text) <> 0 Then
        rsTestSetup!OtherMods = txtOtherMods.Text
    Else
        rsTestSetup!OtherMods = vbNullString
    End If

    rsTestSetup!Approved = False

    I = cmbLoopNumber.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbLoopNumber.ItemData(I)
        rsTestSetup!loopnumber = d
    End If

    I = cmbSuctDia.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbSuctDia.ItemData(I)
        rsTestSetup!SuctDiam = d
    End If

    I = cmbDischDia.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbDischDia.ItemData(I)
        rsTestSetup!DischDiam = d
    End If

    I = cmbTachID.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbTachID.ItemData(I)
        rsTestSetup!tachid = d
    End If

    I = cmbAnalyzerNo.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbAnalyzerNo.ItemData(I)
        rsTestSetup!analyzerno = d
    End If

    I = cmbTestSpec.ListIndex
    If I = -1 Then
        d = 0
    Else
        d = cmbTestSpec.ItemData(I)
    End If
    rsTestSetup!testspec = d

    I = cmbVoltage.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbVoltage.ItemData(I)
        rsTestSetup!Voltage = d
    End If

    I = cmbFrequency.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbFrequency.ItemData(I)
        rsTestSetup!Frequency = d
    End If

    I = cmbMounting.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbMounting.ItemData(I)
        rsTestSetup!Mounting = d
    End If

    I = cmbPLCNo.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbPLCNo.ItemData(I)
        rsTestSetup!PLCNo = d
    End If

    rsTestSetup!ImpFeathered = chkFeathered.value

    If chkTrimmed.value = 1 Then
        rsTestSetup!ImpTrimmed = Val(txtImpTrim)
    Else
        rsTestSetup!ImpTrimmed = 0
    End If
    chkTrimmed_Click

    If chkOrifice.value = 1 Then
        rsTestSetup!PumpDischOrifice = Val(txtOrifice)
    Else
        rsTestSetup!PumpDischOrifice = 0
    End If
    chkOrifice_Click

    If chkCircOrifice.value = 1 Then
        rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
    Else
        rsTestSetup!CircFlowOrifice = 0
    End If
    chkCircOrifice_Click

    If Me.chkAddedDiodes.value = 1 Then
        rsTestSetup!NoOfTRGDiodes = Val(Me.txtNoOfDiodes.Text)
    Else
        rsTestSetup!NoOfTRGDiodes = 0
    End If
    chkAddedDiodes_Click

    chkBalanceHoles_Click

    If chkNPSH.value = 1 Then
        txtNPSHFile.Visible = True
        rsTestSetup!NPSHFile = txtNPSHFile
    Else
        rsTestSetup!NPSHFile = vbNullString
        txtNPSHFile.Visible = False
    End If

    If chkPictures.value = 1 Then
        txtPicturesFile.Visible = True
        rsTestSetup!PictureFile = txtPicturesFile
    Else
        rsTestSetup!PictureFile = vbNullString
        txtPicturesFile.Visible = False
    End If

    If chkVibration.value = 1 Then
        txtVibrationFile.Visible = True
        rsTestSetup!VibrationFile = txtVibrationFile
    Else
        rsTestSetup!VibrationFile = vbNullString
        txtVibrationFile.Visible = False
    End If

    If boWriteDataWritten Then
        rsTestSetup!DataWritten = True
    Else
        rsTestSetup!DataWritten = False
    End If

    'for TEMC Inspection Report
    If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
        rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
    Else
        rsTestSetup!InsulationMeggerVolts = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
        rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
    Else
        rsTestSetup!InsulationMegOhms = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
        rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
    Else
        rsTestSetup!DielectricVolts = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
        rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
    Else
        rsTestSetup!DielectricTime = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
        rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
    Else
        rsTestSetup!HydrostaticValue = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
        rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
    Else
        rsTestSetup!HydrostaticTime = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
        rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
    Else
        rsTestSetup!PneumaticValue = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
        rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
    Else
        rsTestSetup!PneumaticTime = ""
    End If

    I = cmbTestAndInspection(0).ListIndex
    If I = -1 Then
        rsTestSetup!HydrostaticUnits = ""
    Else
        rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
    End If


    I = cmbTestAndInspection(1).ListIndex
    If I = -1 Then
        rsTestSetup!PneumaticUnits = ""
    Else
        rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
    End If

    'use abs to convert from 1 and 0 to boolean
    rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
    rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
    rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
    rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
    rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
    rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
    rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
    rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
    rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
    rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
    rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
    rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
    rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
    rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
    rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)



    'update the database
    rsTestSetup.Update

    If boFoundTestData = False Then     'add 8 records for 8 tests
        AddTestData
    End If

    rsTestSetup.Filter = vbNullString
End Sub
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdFindMagtrols_Click()
    FindMagtrols
End Sub

Private Sub cmdFindPump_Click()
    ' find the pump whose sn is shown

'    Dim m As SNRecord
'    m = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)

    Dim sAns As String
    Dim sSO As String
    Dim sParam As String
    Dim sName As String

    Dim I As Integer

    'set TC and AI labels with default values
    txtTitle(0).Text = "TC 1"
    txtTitle(1).Text = "(F)"
    txtTitle(2).Text = "TC 2"
    txtTitle(3).Text = "(F)"
    txtTitle(4).Text = "TC 3"
    txtTitle(5).Text = "(F)"
    txtTitle(6).Text = "TC 4"
    txtTitle(7).Text = "(F)"
    txtTitle(20).Text = "Circ Flow"
    txtTitle(21).Text = "(GPM)"
    txtTitle(22).Text = "RBH Temp"
    txtTitle(23).Text = "(F)"
    txtTitle(24).Text = "RBH Press"
    txtTitle(25).Text = "(psig)"
    txtTitle(26).Text = "AI 4"
    txtTitle(27).Text = ""


    For I = 0 To 7
        lblAutoMan(I).Caption = "Auto"
    Next I

    txtFlowDisplay.Enabled = False
    txtSuctionDisplay.Enabled = False
    txtDischargeDisplay.Enabled = False
    txtTemperatureDisplay.Enabled = False
    txtAI1Display.Enabled = False
    txtAI2Display.Enabled = False
    txtAI3Display.Enabled = False
    txtAI4Display.Enabled = False


    cmdFindPump.Default = False

    'set all found booleans to false
    boUsingHP = False
    boFoundPump = False
    boPumpIsApproved = False
    boFoundTestSetup = False
    boFoundTestData = False


    'get rid of all test dates in combo box
    For I = cmbTestDate.ListCount - 1 To 0 Step -1
        cmbTestDate.RemoveItem 0
    Next I

    rsTestData.Filter = "SerialNumber = ''"

    DataGrid2.ClearFields
    ClearEff

    If rsPumpData.State = adStateOpen Then
        rsPumpData.Close
    End If

    'find the pump listed in the Serial Number text box
    qyPumpData.ActiveConnection = cnPumpData
    qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                      txtSN.Text & "'))"
    rsPumpData.CursorType = adOpenStatic
    rsPumpData.CursorLocation = adUseClient
    rsPumpData.Index = "SerialNumber"
    rsPumpData.Open qyPumpData

    If rsPumpData.BOF = True And rsPumpData.EOF = True Then
        'if the bof=eof, we have an empty recordset
        boFoundPump = False
    Else
        'we found it
        boFoundPump = True
    End If

'    If InStr(1, txtSN.Text, "-") = 0 Then
'        sAns = MsgBox("There is no dash in the Serial Number.  Please add a dash and try again.", vbOKOnly, "No dash in Serial Number")
'        Exit Sub
'    End If


    If boFoundPump = False Then
        'not found in either database, try HP?
        sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
        If sAns = vbNo Then     'new pump - don't get data from HP
            boUsingEpicor = False
        Else
            boUsingEpicor = True
            boUsingHP = False
        End If
'        If boUsingEpicor = False Then
'            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
'            If sAns = vbNo Then     'new pump - don't get data from HP
                 boUsingHP = False
'            Else
'                boUsingHP = True
'            End If
'        End If
        EnablePumpDataControls
        EnableTestSetupDataControls
        EnableTestDataControls
'        BlankData               'clear any data on the screen
        cmdAddNewBalanceHoles.Visible = True

    End If

    If boFoundPump = True Then    'found the pump
        If rsPumpData.Fields("Approved") = True Then
            DisablePumpDataControls                         'if it's in the real database, don't allow changes here
            boPumpIsApproved = True
            lblPumpApproved.Visible = True
            If boCanApprove Then
                cmdApprovePump.Caption = "Unapprove this pump"
            End If
            frmPLCData.cmdApproveTestDate.Enabled = True
        Else
            EnablePumpDataControls                          'it's in the temp database, allow changes
            boPumpIsApproved = False
            boTestDateIsApproved = False
            lblPumpApproved.Visible = False
            If boCanApprove Then
                cmdApprovePump.Caption = "Approve this pump"
            End If
            cmdApproveTestDate.Caption = "You Must Approve Pump First"
            frmPLCData.cmdApproveTestDate.Enabled = False
        End If

        'found the pump, show the data
        txtModelNo.Text = rsPumpData.Fields("ModelNumber")
        frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

        If rsPumpData.Fields("ChempumpPump") = True Then
            SetCombo cmbMotor, "Motor", rsPumpData
            SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
            SetCombo cmbRPM, "RPM", rsPumpData
            SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
            SetCombo cmbStatorFill, "StatorFill", rsPumpData
            SetCombo cmbModel, "Model", rsPumpData
            SetCombo cmbModelGroup, "ModelGroup", rsPumpData
            RatedKW = 999
        End If

        'set the TEMC data
        If rsPumpData.Fields("ChempumpPump") = False Then
            SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
            SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
            SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
            SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
            SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
            SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
            SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
            SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
            SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
            SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
            SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
            SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
            SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
            SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
            SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
            SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
            SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
        End If

        'write ship to and bill to info
        If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
            txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
        Else
            txtShpNo.Text = vbNullString
        End If

        If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
            txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
        Else
            txtBilNo.Text = vbNullString
        End If

        sName = "ImpellerDia"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtImpellerDia.Text = sParam

        sName = "DesignFlow"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtDesignFlow.Text = sParam

        sName = "DesignTDH"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtDesignTDH.Text = sParam

        sName = "SpGr"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtSpGr.Text = sParam

        sName = "Remarks"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtRemarks.Text = sParam

        sName = "SalesOrderNumber"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtSalesOrderNumber.Text = sParam

        sName = "ApplicationFluid"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtLiquid.Text = sParam

        sName = "NPSHFile"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtNPSHFileLocation.Text = sParam

        sName = "ApplicationViscosity"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = Format(rsPumpData.Fields(sName), "#0.00")
        Else
            sParam = vbNullString
        End If
        txtViscosity.Text = sParam

'added from TEMC Inspection Report
        sName = "JobNumber"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = ""
        End If
        txtJobNum.Text = sParam

        sName = "Phases"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtNoPhases.Text = sParam

        sName = "ThermalClass"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtThermalClass.Text = sParam

        sName = "ExpClass"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtExpClass.Text = sParam

        sName = "NPSHr"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtNPSHr.Text = sParam

        sName = "LiquidTemperature"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtLiquidTemperature.Text = sParam

        sName = "RatedOutput"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtRatedInputPower.Text = sParam

        sName = "FLCurrent"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtAmps.Text = sParam

        sName = "TEMCFrameNumber"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtTEMCFrameNumber.Text = sParam

        optMfr(0).value = rsPumpData.Fields("ChempumpPump")
        optMfr(1).value = Not optMfr(0).value

        'select the testsetup data
        qyTestSetup.ActiveConnection = cnPumpData
        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                      txtSN.Text & "')) ORDER BY Date"
'        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

        With rsTestSetup
            If .State = adStateOpen Then
                .Close
            End If
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .Index = "FindData"
            .Open qyTestSetup
        End With


        'add the selection of dates to the Test Date combo box
        If rsTestSetup.RecordCount <> 0 Then
            For I = 0 To cmbTestDate.ListCount - 1
                cmbTestDate.RemoveItem 0
            Next I
            rsTestSetup.MoveFirst
            For I = 1 To rsTestSetup.RecordCount
                cmbTestDate.AddItem rsTestSetup.Fields("Date")
                rsTestSetup.MoveNext
            Next I
            rsTestSetup.MoveFirst
            boFoundTestSetup = True

            If rsTestSetup.Fields("Approved") = True Then
                DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
                boTestDateIsApproved = True
                lblTestDateApproved.Visible = True
                If boCanApprove Then
                    cmdApproveTestDate.Caption = "Unapprove this Test Date"
                End If
            Else
                EnableTestSetupDataControls                          'it's in the temp database, allow changes
                lblTestDateApproved.Visible = False
                If boCanApprove Then
                    cmdApproveTestDate.Caption = "Approve this Test Date"
                End If
            End If
            cmbTestDate.ListIndex = 0
        Else
            MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
            boFoundTestSetup = False        'didn't find any data
            boFoundTestData = False
            cmbTestDate.AddItem Date        'load with today
            cmbTestDate.ListIndex = 0       'show the entry
            EnableTestSetupDataControls
            txtTestRemarks.Text = ""
            txtVibAx.Text = ""
            txtVibRad.Text = ""
            txtThrustBal.Text = ""
            txtTEMCTRGReading.Text = ""
            txtTEMCFrontThrust.Text = ""
            txtTEMCRearThrust.Text = ""
            Exit Sub
        End If

        If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
        End If
        Exit Sub
    End If

    'hp interface stuff
    If boUsingHP = True Then
        If InStr(1, txtSN.Text, "-") = 0 Then
            MsgBox "Please check the Serial Number.  There doesn't seem to be a -", vbOKOnly
            Exit Sub
        Else
            sSO = Left$(txtSN.Text, 7)       'look for the sales order
            If Len(sSO) <> 7 Then
                MsgBox "Please check the Serial Number.  There doesn't seem to be 7 digits before the -", vbOKOnly
                Exit Sub
            End If
        End If

        If Not cnHPOpen Then
            frmConnectingToPLC.Show
            DoEvents
            With cnHP
                .ConnectionString = sHPDataBaseName
                .CommandTimeout = 10
                .Open
            End With
            cnHPOpen = True
            frmConnectingToPLC.Hide
            DoEvents
        End If

        txtSalesOrderNumber.Text = sSO

        GetDetail sSO

        SearchSalesOrder txtSalesOrderNumber.Text

        frmPLCData.txtShpNo.Text = strShipTo
        frmPLCData.txtBilNo.Text = strBillTo

        rsHPDetail.Filter = 0

        If rsHPDetail.BOF = True And rsHPDetail.EOF = True Then
            MsgBox "Not found on HP.  You may enter data manually."
        Else
            frmPLCData.txtModelNo = strModelNo(1, intLineNo)
            frmPLCData.txtDesignTDH.Text = strTDH(1, intLineNo)
            frmPLCData.txtSpGr.Text = strSpGr(1, intLineNo)
            frmPLCData.txtImpellerDia.Text = strImpellers(1, intLineNo)
            frmPLCData.txtDesignFlow.Text = strCapacity(1, intLineNo)

            For I = 0 To cmbStatorFill.ListCount - 1
                If InStr(1, UCase$(strStatorFill(1, intLineNo)), UCase$(cmbStatorFill.List(I))) <> 0 Then
                    cmbStatorFill.ListIndex = I
                    Exit For
                End If
            Next I

            For I = 0 To cmbDesignPressure.ListCount - 1
                If InStr(1, strDesignPress(1, intLineNo), cmbDesignPressure.List(I)) <> 0 Then
                    cmbDesignPressure.ListIndex = I
                    Exit For
                End If
            Next I

            I = InStr(strVoltage(1, intLineNo), "VOLT")
            sName = strFindTheNumber(strVoltage(1, intLineNo), I)

            For I = 0 To cmbVoltage.ListCount - 1
                If InStr(1, sName, cmbVoltage.List(I)) <> 0 Then
                    cmbVoltage.ListIndex = I
                    Exit For
                End If
            Next I

            I = InStr(strVoltage(1, intLineNo), "CY")
            sName = strFindTheNumber(strVoltage(1, intLineNo), I)

            For I = 0 To cmbFrequency.ListCount - 1
                If InStr(1, cmbFrequency.List(I), sName) <> 0 Then
                    cmbFrequency.ListIndex = I
                    Exit For
                End If
            Next I

            For I = 0 To cmbRPM.ListCount - 1
                If InStr(1, strRPM(1, intLineNo), cmbRPM.List(I)) <> 0 Then
                    cmbRPM.ListIndex = I
                    Exit For
                End If
            Next I

            For I = 0 To cmbSuctDia.ListCount - 1
                If InStr(1, strSuctFlg(1, intLineNo), cmbSuctDia.List(I)) <> 0 Then
                    cmbSuctDia.ListIndex = I
                    Exit For
                End If
            Next I

            For I = 0 To cmbDischDia.ListCount - 1
                If InStr(1, strDischFlg(1, intLineNo), cmbDischDia.List(I)) <> 0 Then
                    cmbDischDia.ListIndex = I
                    Exit For
                End If
            Next I

            For I = 0 To cmbTestSpec.ListCount - 1
                If InStr(1, strTestProcedure(1, intLineNo), cmbTestSpec.List(I)) <> 0 Then
                    cmbTestSpec.ListIndex = I
                    Exit For
                End If
            Next I

            rsHPDetail.MoveFirst
            Load FrmSODetails
            FrmSODetails.Show
            FrmSODetails.txtSOData.Text = vbNullString
        End If

        Dim intLastLineNo As Integer
        Dim vFilter As Variant

        intLastLineNo = 0

        If rsHPLineNo.State = adStateOpen Then
            Do While Not rsHPDetail.EOF
                If Int(Val(rsHPDetail.Fields(1))) <> intLastLineNo Then
                    intLastLineNo = Val(rsHPDetail.Fields(1))
                    FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & vbCrLf & "Line No. = " & intLastLineNo & " Quan = "
                    vFilter = "LINE = '" & str$(intLastLineNo) & "'"
                    rsHPLineNo.Filter = vFilter
                    If rsHPLineNo.BOF = True And rsHPLineNo.EOF = True Then
                        FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & vbCrLf
                    Else
                        FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & rsHPLineNo.Fields(2) & vbCrLf
                    End If
                    rsHPLineNo.Filter = 0
                End If
                FrmSODetails.txtSOData.Text = FrmSODetails.txtSOData.Text & "   " & rsHPDetail.Fields(2) & vbCrLf
                rsHPDetail.MoveNext
            Loop
        End If
    End If

    If boUsingEpicor = True Then
        Dim MyRecord As SNRecord
'            I = InStr(1, txtSN.Text, "-")
'            If I > 0 Then
            MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
'            End If
        If MyRecord.SONumber = "" Then
            MsgBox ("Not found in Epicor")
            Exit Sub
        End If
        txtSalesOrderNumber.Text = MyRecord.SONumber
        txtLineNumber.Text = MyRecord.SOLine
        txtBilNo.Text = MyRecord.Customer
        If MyRecord.ShipTo = "" Then
            txtShpNo.Text = MyRecord.Customer
        Else
            txtShpNo.Text = MyRecord.ShipTo
        End If
        txtModelNo.Text = MyRecord.ModelNo
        txtModelNo_Change
        txtDesignTDH.Text = MyRecord.TDH
        txtSpGr.Text = MyRecord.SpGr
        txtImpellerDia.Text = MyRecord.ImpellerDiameter
        txtDesignFlow.Text = MyRecord.Flow
        txtNoPhases.Text = MyRecord.Phases
        txtNPSHr.Text = MyRecord.NPSHr
        txtRatedInputPower.Text = MyRecord.RatedInputPower
        txtAmps.Text = MyRecord.FLCurrent
        txtThermalClass.Text = MyRecord.ThermalClass
        txtViscosity.Text = MyRecord.Viscosity
        txtExpClass.Text = MyRecord.ExpClass
        txtLiquidTemperature.Text = MyRecord.LiquidTemp
        txtLiquid.Text = MyRecord.Fluid
        txtJobNum.Text = MyRecord.JobNumber
        txtSpHeat.Text = MyRecord.SpecHeat

        For I = 0 To cmbStatorFill.ListCount - 1
            If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
                cmbStatorFill.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbCirculationPath.ListCount - 1
            If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
                cmbCirculationPath.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbDesignPressure.ListCount - 1
            If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
                cmbDesignPressure.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbVoltage.ListCount - 1
            If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
                cmbVoltage.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbFrequency.ListCount - 1
            If InStr(1, MyRecord.Frequency, sName) <> 0 Then
                cmbFrequency.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbRPM.ListCount - 1
            If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
                cmbRPM.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbSuctDia.ListCount - 1
            If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
                cmbSuctDia.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbDischDia.ListCount - 1
            If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
                cmbDischDia.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbTestSpec.ListCount - 1
            If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
                cmbTestSpec.ListIndex = I
                Exit For
            End If
        Next I

        For I = 0 To cmbMotor.ListCount - 1
            If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
                cmbMotor.ListIndex = I
                Exit For
            End If
        Next I


    End If
  
End Sub

Private Sub cmdModifyBalanceHoleData_Click()
    Dim strInput As String
    Dim I As Integer
    Dim sNumber As Integer
    Dim sDia As String
    Dim sBC As String

    cmdModifyBalanceHoleData.Visible = False

    If dgBalanceHoles.SelBookmarks.Count = 0 Then
        cmdModifyBalanceHoleData.Visible = False
        Exit Sub
    End If

    rsBalanceHoles.MoveFirst
    rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

    sNumber = rsBalanceHoles!Number
    If rsBalanceHoles!diameter = 99 Then
        sDia = "Slot"
    Else
        sDia = str(rsBalanceHoles!diameter)
    End If
    If rsBalanceHoles!boltcircle = 99 Then
        sBC = "Unknown"
    Else
        sBC = str(rsBalanceHoles!boltcircle)
    End If


    'get the data for the balance holes
    strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
    If strInput = "" Then
        GoTo DeleteIt
    End If
    sNumber = CInt(strInput)
    If Val(sNumber) = 0 Then
        GoTo DeleteIt
    End If

    strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
    If strInput <> "" Then
        If UCase(strInput) = "SLOT" Then
            strInput = 99
        End If
        sDia = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
    If strInput <> "" Then
        If UCase(strInput) = "UNKNOWN" Then
            strInput = 99
        End If
        sBC = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    rsBalanceHoles!Number = sNumber
    rsBalanceHoles!diameter = sDia
    rsBalanceHoles!boltcircle = sBC

    rsBalanceHoles.Update
    'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

    GetBalanceHoleData txtSN.Text, cmbTestDate.Text
'    rsBalanceHoles.Requery
    rsBalanceHoles.MoveLast
    dgBalanceHoles.Refresh
    chkBalanceHoles.value = 1
    rsBalanceHoles.MoveFirst

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
    Exit Sub

CancelPressed:
    MsgBox "No New Balance Hole Data Entered", vbOKOnly

DeleteIt:
    If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
        rsBalanceHoles.Delete
        rsBalanceHoles.Update
        GetBalanceHoleData txtSN.Text, cmbTestDate.Text
'        rsBalanceHoles.Requery
        If rsBalanceHoles.RecordCount > 0 Then
            rsBalanceHoles.MoveLast
        End If
        dgBalanceHoles.Refresh
        chkBalanceHoles.value = 1
        rsBalanceHoles.MoveFirst
    End If

  
End Sub

Private Sub cmdReport_Click()
    'view/print a report
    Dim I As Integer

    frmReport.Visible = True
    For I = 0 To optReport.Count - 1
        optReport(I).value = False
    Next I
  
End Sub

Private Sub cmdSearchForPump_Click()
    LoadCombo frmSearch.cmbSearchModel, "Model"
    frmSearch.Show
End Sub

Private Sub cmdWriteSP_Click()
    'write the sp to the plc
    Dim rc As String
    Dim S As String

    'write the set point data to the PLC
        bWrite = True
        S = Right$("0000" & txtWriteSPData, 4)
        S = Right$(S, 2) & Left$(S, 2)
        rc = StringToByteArray(S, ByteBuffer)

        DataLength = HexConvert(ByteBuffer, 2)
        DataAddress = StringToHexInt("2005")

        rc = GetData

        bWrite = False
End Sub



Private Sub Command1_Click()
'    Dim frmem As New InteropDBWithButtons.Form1
'    frmem.ConString = cnPumpData.ConnectionString
'    frmem.Caption = "Email Database Maintenance"
'    frmem.Show 1
End Sub

Private Sub Command2_Click()
    ReportToExcel
End Sub







Private Sub updown1_change()
    Dim sName As String

    If Not rsTestData.BOF Then
        rsTestData.MoveFirst
    End If

    If Not rsTestData.BOF Or Not rsTestData.EOF Then
        rsTestData.Move UpDown1.value - 1
    End If

    sName = "VibrationX"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtVibAx.Text = rsTestData.Fields(sName)
    Else
'        txtVibAx.Text = vbNullString
    End If

    sName = "VibrationY"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtVibRad.Text = rsTestData.Fields(sName)
    Else
'        txtVibRad.Text = vbNullString
    End If

    sName = "Remarks"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTestRemarks.Text = rsTestData.Fields(sName)
    Else
'        txtTestRemarks.Text = vbNullString
    End If

    sName = "ThrustBalance"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtThrustBal.Text = rsTestData.Fields(sName)
    Else
'        txtThrustBal.Text = vbNullString
    End If

    sName = "TEMCTRG"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCTRGReading.Text = rsTestData.Fields(sName)
    Else
        txtTEMCTRGReading.Text = 0
'        txtTEMCTRGReading.Text = vbNullString
    End If

    sName = "TEMCFrontThrust"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCFrontThrust.Text = vbNullString
    End If

    sName = "TEMCRearThrust"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCRearThrust.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCRearThrust.Text = vbNullString
    End If
    sName = "TEMCMomentArm"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCMomentArm.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCMomentArm.Text = vbNullString
    End If
    sName = "TEMCThrustRigPressure"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCThrustRigPressure.Text = vbNullString
    End If
    sName = "TEMCViscosity"
    If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
        txtTEMCViscosity.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCViscosity.Text = vbNullString
    End If

    CalculateTEMCForce

    rsEff.MoveFirst
    rsEff.Move UpDown1.value - 1
End Sub
Sub CalculateTEMCForce()
    Dim NoOfPoles As Integer
    Dim Frequency As Integer
    Dim Additions As String
    Dim Frame As String
    Dim VOverA As Double
    Dim Force As Double

    'show calculated values
    If Val(txtTEMCFrontThrust.Text) = 0 Then
        If Val(txtTEMCRearThrust.Text) = 0 Then
        'no thrust entered
            lblTEMCFrontRear.Visible = False
            txtTEMCCalcForce.Text = " "
        Else
            'rear thrust
            txtTEMCCalcForce.Text = Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5
            lblTEMCFrontRear.Caption = "REAR"
            lblTEMCFrontRear.Visible = True
        End If
    Else
        'front thrust
        txtTEMCCalcForce.Text = Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5
        lblTEMCFrontRear.Caption = "FRONT"
        lblTEMCFrontRear.Visible = True
    End If

    If Val(txtTEMCCalcForce.Text) < 0 Then
        txtTEMCCalcForce.Text = -txtTEMCCalcForce
        lblTEMCFrontRear.Caption = "FRONT"
    End If

    'see how many poles we have, it's the next to last number in the frame size
    If Len(txtTEMCFrameNumber) > 2 Then
        NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
    End If

    If cmbTEMCAdditions.ListIndex <> -1 Then
        Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
        If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
            Frequency = 60
        ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
            Frequency = 50
        Else
            Frequency = 0
        End If
    End If

    If Len(txtTEMCFrameNumber.Text) = 3 Then
        Frame = Left$(txtTEMCFrameNumber, 2) & "0"
    Else
        Frame = txtTEMCFrameNumber.Text
        If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
            Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
        Else
        End If
    End If
    Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
    If Frequency = 60 Then
        Force = Force / 1.2
    End If
    If Val(txtTEMCViscosity.Text) > 1# Then
        If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If

    If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
        If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If

    If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
        If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If

    If (Val(txtTEMCViscosity.Text) <= 0.3) Then
        If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If
    If NoOfPoles <> 0 Then
        VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles / 2)
    End If
    If Frequency = 60 Then
        VOverA = VOverA * 1.2
    End If

    txtTEMCPVValue.Text = Val(txtTEMCCalcForce.Text) * VOverA

    If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
        txtTEMCPVValue.Text = ""
        txtTEMCCalcForce.Text = ""
        lblTEMCPassFail.Visible = False
    End If
  
End Sub
Private Sub UpDown2_change()
    Dim Plothead(7, 1) As Single
    Dim PlotEff(7, 1) As Single
    Dim PlotKW(7, 1) As Single
    Dim PlotAmps(7, 1) As Single

    Dim j As Integer

    For j = 0 To UpDown2.value - 1
        Plothead(j, 0) = HeadFlow(0, j)
        Plothead(j, 1) = HeadFlow(1, j)

        PlotEff(j, 0) = EffFlow(0, j)
        PlotEff(j, 1) = EffFlow(1, j)
        PlotKW(j, 0) = KWFlow(0, j)
        PlotKW(j, 1) = KWFlow(1, j)
        PlotAmps(j, 0) = AmpsFlow(0, j)
        PlotAmps(j, 1) = AmpsFlow(1, j)
    Next j

    MSChart1 = Plothead
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(Plothead) / 10) + 0.5) + 1)
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

    MSChart3 = PlotAmps
    MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(PlotAmps) / 10) + 0.5) + 1)
    MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
    MSChart3.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

    MSChart4 = PlotKW
    MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(PlotKW) / 10) + 0.5) + 1)
    MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
    MSChart4.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

    MSChart5 = PlotEff
    MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(PlotEff) / 10) + 0.5) + 1)
    MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
    MSChart5.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

    MSChart6 = Plothead
    MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(Plothead) / 10) + 0.5) + 1)
    MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
    MSChart6.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5
  
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    DoEfficiencyCalcs
End Sub

Private Sub dgBalanceHoles_SelChange(Cancel As Integer)
    If dgBalanceHoles.SelBookmarks.Count = 0 Then
        cmdModifyBalanceHoleData.Visible = False
    Else
        cmdModifyBalanceHoleData.Visible = True
    End If
End Sub
Private Sub Form_Load()
    Dim RetVal As String
    Dim sSendStr As String
    Dim I As Integer
    Dim j As Integer
    Dim sTableName As String
    Dim WhichServer As String
    Dim WhichDatabase As String

    debugging = 0   'assume not debugging
    WhichServer = "Production"     'change to production server
    WhichDatabase = "Production"

    If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
        I = MsgBox("Debug?", vbYesNo)
        If I = vbYes Then
            debugging = 1
            WhichServer = "Production"
            WhichDatabase = "Production"
        Else
        End If
    End If

    If debugging Then
        GoTo temp
    End If
    'see if the mdb file is where it's supposed to be
    If Dir(sDevelopmentDatabase) = "" Then
        MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
        End
    End If

    'get the database info from the new mdb file
    Dim cnDevelopment As New ADODB.Connection
    Dim qyDevelopment As New ADODB.Command
    Dim rsDevelopment As New ADODB.Recordset

    On Error GoTo CannotConnect

    With cnDevelopment
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDevelopmentDatabase & ";Persist Security Info=False;Jet OLEDB:Database Password=Access7277word;"
        .ConnectionTimeout = 10
        .Open
    End With

On Error GoTo 0
    GoTo Connected

CannotConnect:
    MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
    End

Connected:

    'we're connected, get the data for the Epicor SQL server
    qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
    qyDevelopment.ActiveConnection = cnDevelopment

    rsDevelopment.CursorLocation = adUseClient
    rsDevelopment.CursorType = adOpenStatic
    rsDevelopment.LockType = adLockOptimistic

    On Error GoTo NoServerData

    rsDevelopment.Open qyDevelopment

On Error GoTo 0
    GoTo GotServerData

NoServerData:

    MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
    End

GotServerData:

    If rsDevelopment.RecordCount <> 1 Then
        GoTo NoServerData
    End If

    'construct Epicor connection string
    EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                           "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                           "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                           "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                           "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


    'make sure we can open the SQL database

    On Error GoTo CannotOpenEpicorSQLServer

    Dim cnTestEpicor As New ADODB.Connection
    cnTestEpicor.ConnectionString = EpicorConnectionString
    cnTestEpicor.Open
    cnTestEpicor.Close
    Set cnTestEpicor = Nothing
On Error GoTo 0

    GoTo FoundEpicorSQLServer

CannotOpenEpicorSQLServer:
    MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
    End

FoundEpicorSQLServer:
    'get data on rundown database
    rsDevelopment.Close
    qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PumpRundown'"

    On Error GoTo NoRundownDatabase

    rsDevelopment.Open qyDevelopment

    GoTo FoundRundownDatabase

NoRundownDatabase:
    MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
    End

FoundRundownDatabase:
    If rsDevelopment.RecordCount <> 1 Then
        GoTo NoRundownDatabase
        End
    End If

temp:

    If debugging Then
        ParentDirectoryName = "C:\databases"
        sDataBaseName = "c:\databases\PumpData 2k.mdb"
    Else
        ParentDirectoryName = "\\TEI-MAIN-01\f\groups\shared\databases"
        sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

'        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
    End If

    'see if we can open the pump rundown database
    On Error GoTo NoRundownDatabase
    With cnPumpData
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False"
        .ConnectionTimeout = 10
        .Open
    End With
On Error GoTo 0


    If debugging = 0 Then
        Printer.Orientation = vbPRORLandscape
    End If

    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    boFoundPump = False

    Me.Show

    Dim k As Integer
    For k = 0 To 20
        vPlot(k, 0) = 0
        vPlot(k, 1) = 0
    Next k

    With MSChart1
        .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
        .Plot.Axis(VtChAxisIdY).AxisTitle = "TDH (Ft)"
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
        .Plot.SeriesCollection.Item(1).Pen.Width = 15
        .Plot.AutoLayout = False
        .Plot.LocationRect.Max.X = 5700
        .Plot.LocationRect.Max.Y = 3000
        .Plot.LocationRect.Min.X = -100
        .Plot.LocationRect.Min.Y = -100
    End With
    With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
        .Visible = True
        .Size = 60
        .Style = VtMarkerStyleCircle
        .FillColor.Automatic = False
        .FillColor.Set 0, 0, 255
    End With

    With MSChart2
'
        .Plot.AutoLayout = False
        .Plot.LocationRect.Max.X = 2000
        .Plot.LocationRect.Max.Y = 1500
        .Plot.LocationRect.Min.X = 0
        .Plot.LocationRect.Min.Y = 0

    End With

    With MSChart3
        .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
        .Plot.Axis(VtChAxisIdY).AxisTitle = "Current (Amps)"
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
        .Plot.SeriesCollection.Item(1).Pen.Width = 15
        .Plot.AutoLayout = False
        .Plot.LocationRect.Max.X = 5700
        .Plot.LocationRect.Max.Y = 3000
        .Plot.LocationRect.Min.X = -100
        .Plot.LocationRect.Min.Y = -100
    End With
    With MSChart3.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
        .Visible = True
        .Size = 60
        .Style = VtMarkerStyleCircle
        .FillColor.Automatic = False
        .FillColor.Set 0, 0, 255
    End With
    With MSChart4
        .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
        .Plot.Axis(VtChAxisIdY).AxisTitle = "kW"
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
        .Plot.SeriesCollection.Item(1).Pen.Width = 15
        .Plot.AutoLayout = False
        .Plot.LocationRect.Max.X = 5700
        .Plot.LocationRect.Max.Y = 3000
        .Plot.LocationRect.Min.X = -100
        .Plot.LocationRect.Min.Y = -100
    End With
    With MSChart4.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
        .Visible = True
        .Size = 60
        .Style = VtMarkerStyleCircle
        .FillColor.Automatic = False
        .FillColor.Set 0, 0, 255
    End With
    With MSChart5
        .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
        .Plot.Axis(VtChAxisIdY).AxisTitle = "Efficiency (%)"
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
        .Plot.SeriesCollection.Item(1).Pen.Width = 15
        .Plot.AutoLayout = False
        .Plot.LocationRect.Max.X = 5700
        .Plot.LocationRect.Max.Y = 3000
        .Plot.LocationRect.Min.X = -100
        .Plot.LocationRect.Min.Y = -100
    End With
    With MSChart5.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
        .Visible = True
        .Size = 60
        .Style = VtMarkerStyleCircle
        .FillColor.Automatic = False
        .FillColor.Set 0, 0, 255
    End With

    With MSChart6
        .Plot.Axis(VtChAxisIdX).AxisTitle = "Flow (GPM)"
        .Plot.Axis(VtChAxisIdY).AxisTitle = "TDH (Ft)"
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdX).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Size = 10
        .Plot.Axis(VtChAxisIdY).AxisTitle.VtFont.Style = VtFontStyleBold
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
        .Plot.SeriesCollection.Item(1).Pen.Width = 15
        .Plot.AutoLayout = False
        .Plot.LocationRect.Max.X = 5700
        .Plot.LocationRect.Max.Y = 3000
        .Plot.LocationRect.Min.X = -100
        .Plot.LocationRect.Min.Y = -100
    End With
    With MSChart6.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
        .Visible = True
        .Size = 60
        .Style = VtMarkerStyleCircle
        .FillColor.Automatic = False
        .FillColor.Set 0, 0, 255
    End With

    'assure that the timers are off
    frmPLCData.tmrGetDDE.Enabled = False

    frmPLCData.tmrStartUp.Enabled = False

    'initialize the PLC network
    RetVal = NetWorkInitialize()
    If RetVal <> 0 Then
        MsgBox ("Can't Initialize Network. Exiting...")
        End
    End If

    If debugging = 0 Then
        'load array of plcs
        I = 0
        Open rsDevelopment.Fields("ServerName") & "plcaddresses.txt" For Input As 1
        While Not EOF(1)
            Input #1, Description(I)
            For j = 0 To 125
                Input #1, aDevices(I).Address(j)
            Next j
            Input #1, j
            I = I + 1
        Wend
        Close #1

        DeviceCount = I

        If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
'            Dim k As Integer
            For k = 0 To DeviceCount - 1
                If InStr(Description(k), "MWSC") <> 0 Then
                    Exit For
                End If
            Next k
            Description(DeviceCount) = Description(0)
            Description(0) = Description(k)
            Description(k) = Description(DeviceCount)

            aDevices(DeviceCount) = aDevices(0)
            aDevices(0) = aDevices(k)
            aDevices(k) = aDevices(DeviceCount)

        End If

        Dim PLCAddress As String
        For I = 0 To DeviceCount - 1
            PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
            RetVal = PingSilent(PLCAddress)
            If RetVal <> 0 Then
                frmPLCData.cmbPLCLoop.AddItem Description(I)
                frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
            End If
        Next I
    End If

    frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

    'turn on the PLC led

    frmPLCData.cmbPLCLoop.ListIndex = 0
    frmPLCData.tmrGetDDE.Enabled = True

    'hook up to the various databases

    DataEnvironment2.Connection1.ConnectionString = cnPumpData
    DataEnvironment3.Connection1.ConnectionString = cnPumpData

    With cnEffData
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
        .Open
    End With

    'open some recordsets
    rsPumpData.Index = "SerialNumber"
    rsTestSetup.Index = "FindData"
    rsTestData.Index = "PrimaryKey"
    rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
    rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
    rsTestData.Filter = "SerialNumber = ''"
    rsTestData.CursorLocation = adUseClient
    rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
    rsEff.CursorLocation = adUseClient
    rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

    qyBalanceHoles.ActiveConnection = cnPumpData
    rsBalanceHoles.CursorLocation = adUseClient
    rsBalanceHoles.CursorType = adOpenStatic
    rsBalanceHoles.LockType = adLockOptimistic
'    qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 FROM BalanceHoles;"
'    rsBalanceHoles.Open qyBalanceHoles
'    rsBalanceHoles.Filter = "SerialNo = ''"



    If debugging <> 1 Then
        FindMagtrols
    Else
        cmbMagtrol.AddItem "Add Manually"
        cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
        cmbMagtrol.ListIndex = 0
    End If
    optKW(1).value = True
    optKW_Click (1)


    'blank out data grid
    Set DataGrid1.DataSource = rsTestData

    'load the combo boxes
    LoadCombo cmbStatorFill, "StatorFill"
    LoadCombo cmbCirculationPath, "CirculationPath"
    LoadCombo cmbVoltage, "Voltage"
    LoadCombo cmbFrequency, "Frequency"
    LoadCombo cmbMotor, "Motor"
    LoadCombo cmbDesignPressure, "DesignPressure"
    LoadCombo cmbRPM, "RPM"
    LoadCombo cmbOrificeNumber, "OrificeNumber"
    LoadCombo cmbTestSpec, "TestSpecification"
    LoadCombo cmbLoopNumber, "LoopNumber"
    LoadCombo cmbSuctDia, "SuctionDiameter"
    LoadCombo cmbDischDia, "DischargeDiameter"
    LoadCombo cmbModel, "Model"
    LoadCombo cmbModelGroup, "ModelGroup"
    LoadCombo cmbMounting, "Mounting"
    LoadCombo cmbTachID, "TachID"
    LoadCombo cmbAnalyzerNo, "AnalyzerNo"
    LoadCombo cmbPLCNo, "PLCNo"
    LoadCombo cmbFlowMeter, "Flowmeter"
'    LoadInstrumentationCombo cmbTachID, "TachID"
'    LoadInstrumentationCombo cmbAnalyzerNo, "AnalyzerNo"
'    LoadInstrumentationCombo cmbPLCNo, "PLCNo"
'    LoadInstrumentationCombo cmbFlowMeter, "Flowmeter"

    'load the TEMC combo boxes, too
    LoadCombo cmbTEMCAdapter, "TEMCAdapter"
    LoadCombo cmbTEMCAdditions, "TEMCAdditions"
    LoadCombo cmbTEMCCirculation, "TEMCCirculation"
    LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
    LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
    LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
    LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
    LoadCombo cmbTEMCInsulation, "TEMCInsulation"
    LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
    LoadCombo cmbTEMCMaterials, "TEMCMaterials"
    LoadCombo cmbTEMCModel, "TEMCModel"
    LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
    LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
    LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
    LoadCombo cmbTEMCVoltage, "TEMCVoltage"
    LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
    LoadCombo cmbTEMCTRG, "TEMCTRG"

    LoadCombo frmSearch.cmbSearchModel, "Model"

    'fill memory arrays for dlookups
    FillArrays

    'choose the first tab
    frmPLCData.SSTab1.Tab = 0

    'set the grid column names
    Dim c As Column
    For Each c In DataGrid1.Columns
        Select Case c.DataField
        Case "TestDataID"
            c.Visible = False
        Case "SerialNumber"
            c.Visible = False
        Case "Date"
            c.Visible = False
        Case Else ' Show all other columns.
            c.Visible = True
            c.Alignment = dbgRight
        End Select
    Next c

    Set dgBalanceHoles.DataSource = rsBalanceHoles

    For Each c In dgBalanceHoles.Columns
        Select Case c.DataField
        Case "BalanceHoleID"
            c.Visible = False
        Case "SerialNo"
            c.Visible = False
        Case "Date"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 2000
        Case "Number"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "Diameter"
            c.Visible = False
        Case "Diameter1"
            c.Caption = "Diameter"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "BoltCircle1"
            c.Caption = "Bolt Circle"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 800
        Case "BoltCircle"
            c.Visible = False
        Case "SetNo"
            c.Visible = False
        Case Else ' Show all other columns.
            c.Visible = False
        End Select
    Next c

    BlankData

'    If debugging <> 1 Then
        'get user initials
        frmLogin.Show
'    End If

    'setup eff.mdb file
    DataEnvironment1.Connection2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
'    DataEnvironment3.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"

    'set dropdown width
    SendMessage cmbAnalyzerNo.hwnd, &H160, SendMessage(cmbAnalyzerNo.hwnd, &H15F, 0, 0) + 100, 0
    SendMessage cmbPLCNo.hwnd, &H160, SendMessage(cmbPLCNo.hwnd, &H15F, 0, 0) + 100, 0
    SendMessage cmbTachID.hwnd, &H160, SendMessage(cmbTachID.hwnd, &H15F, 0, 0) + 100, 0
    SendMessage cmbFlowMeter.hwnd, &H160, SendMessage(cmbFlowMeter.hwnd, &H15F, 0, 0) + 100, 0

    FromStoredData = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label15_Click()
    frmDiagram.Show
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

    Dim blnEnabled As Boolean

    If lblAutoMan(Index).Caption = "Auto" Then
        lblAutoMan(Index).Caption = "Man"
        blnEnabled = True
    Else
        lblAutoMan(Index).Caption = "Auto"
        blnEnabled = False
    End If

    Select Case Index
        Case 0
            txtFlowDisplay.Enabled = blnEnabled
        Case 1
            txtSuctionDisplay.Enabled = blnEnabled
        Case 2
            txtDischargeDisplay.Enabled = blnEnabled
        Case 3
            txtTemperatureDisplay.Enabled = blnEnabled
        Case 4
            txtAI1Display.Enabled = blnEnabled
        Case 5
            txtAI2Display.Enabled = blnEnabled
        Case 6
            txtAI3Display.Enabled = blnEnabled
        Case 7
            txtAI4Display.Enabled = blnEnabled
    End Select
  
End Sub



Private Sub txtNPSHFileLocation_Click()
    Dim sTempDir As String
    On Error Resume Next
    sTempDir = CurDir    'Remember the current active directory
    CommonDialog2.DialogTitle = "Select a directory" 'titlebar
    CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
    CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
    CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
    CommonDialog2.CancelError = True 'allow escape key/cancel
    CommonDialog2.ShowSave   'show the dialog screen

    If Err <> 32755 Then    ' User didn't chose Cancel.
        'Me.SDir.Text = CurDir
    End If

'    ChDir sTempDir  'restore path to what it was at entering

Me.txtNPSHFileLocation.Text = CommonDialog2.filename
  
End Sub



Private Sub txtTitle_LostFocus(Index As Integer)

    ChangeTitles Index
  
End Sub
Private Sub ChangeTitles(ChannelNo As Integer)
    Dim I As Integer
    Dim S As String

    If txtTitle(ChannelNo).Locked = True Then
        Exit Sub
    End If

    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

    qy.ActiveConnection = cnPumpData

    'see if we have an entry in the table
    qy.CommandText = "SELECT * FROM AITitles " & _
               "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
               "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
               "AND ((AITitles.Channel)=" & ChannelNo & "));"

    With rs     'open the recordset for the query
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open qy
    End With

    If (rs.BOF = True And rs.EOF = True) Then  'new record
        rs.AddNew
        rs.Fields("SerialNo") = txtSN.Text
        rs.Fields("Date") = cmbTestDate.Text
        rs.Fields("Channel") = CByte(ChannelNo)
        rs.Fields("Title") = txtTitle(ChannelNo).Text
        rs.Update
    Else    'we have an entry, modify it
        rs.Fields("SerialNo") = txtSN.Text
        rs.Fields("Date") = cmbTestDate.Text
        rs.Fields("Channel") = CByte(ChannelNo)
        rs.Fields("Title") = txtTitle(ChannelNo).Text
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
    Set qy = Nothing
  
End Sub

Private Sub optKW_Click(Index As Integer)
    Select Case Index
        Case 0  'add 3 powers
            txtKW.Enabled = False
        Case 1  'enter kw
            txtKW.Enabled = True
        Case 2  'use analog in 4
            txtKW.Enabled = False
    End Select
End Sub

Private Sub optMfr_Click(Index As Integer)
    frmTEMC.Visible = optMfr(1).value
    frmChempump.Visible = optMfr(0).value
    frmTEMCData.Visible = optMfr(1).value
    txtModelNo_Change
End Sub

Private Sub optReport_Click(Index As Integer)
    'choose a report to view/print

    'see if we have balance hole data for this pump
    Dim strBH As String
    Dim I As Integer

    If Index = 6 Then       'cancel pressed
        frmReport.Visible = False
        Exit Sub
    End If

    If Index <> 7 Then
        If boGotBalanceHoles Then
            If rsBalanceHoles.State = adStateClosed Then
                rsBalanceHoles.ActiveConnection = cnPumpData
                rsBalanceHoles.Open
            End If

            If rsBalanceHoles.RecordCount <> 0 Then
                rsBalanceHoles.MoveFirst
                strBH = Space(20) & "Balance Hole Data" & vbCrLf & "Date" & Space(24) & "Number" & Space(6) & "Dia" & Space(7) & "BC"
                For I = 1 To rsBalanceHoles.RecordCount
                    strBH = strBH & vbCrLf & _
                           Left$(rsBalanceHoles.Fields("Date") & Space(30), 30) & _
                           Left$(rsBalanceHoles.Fields("Number") & Space(20), 10) & _
                           Left$(rsBalanceHoles.Fields("Diameter1") & Space(20), 10) & _
                           rsBalanceHoles.Fields("BoltCircle1")
                    rsBalanceHoles.MoveNext
                Next I
            Else
            End If
        Else
'            MsgBox "There are no balance holes for this test date", vbOKOnly, "No Balance Holes"
'            frmReport.Visible = False
'            Exit Sub
        End If
    End If


    Dim strVersion As String
    Dim isChart As Boolean

    isChart = False

    strVersion = "V" & App.Major & "." & App.Minor & "." & App.Revision

    Dim dr As DataReport
    frmReport.Visible = False
    optReport(Index).value = False

    Select Case Index
        Case 0  'Customer no circ flow
            Set dr = drCustNoCirc
        Case 1  'customer circ flow
            Set dr = drCustCirc
            frmReportOptions.Show 1

            Dim PosRPM As Integer
            Dim PosAxPos As Integer
            Dim PosCircFlow As Integer
            Dim PosVib As Integer
            Dim PosRem As Integer
            Dim PosTRG As Integer

                    PosTRG = frmReportOptions.chkTRG.value * 7920
                    PosRPM = frmReportOptions.chkSelectRPM.value * (7920 + frmReportOptions.chkTRG.value * 720)
                    PosAxPos = frmReportOptions.chkSelectAxPos.value * (7920 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)
                    PosCircFlow = frmReportOptions.chkSelectCircFlow.value * (7920 + frmReportOptions.chkSelectAxPos.value * 720 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)
                    PosVib = frmReportOptions.chkVibration.value * (7920 + frmReportOptions.chkSelectCircFlow.value * 720 + frmReportOptions.chkSelectAxPos.value * 720 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)
                    PosRem = 7920 + (frmReportOptions.chkVibration.value * 2 * 720 + frmReportOptions.chkSelectCircFlow.value * 720 + frmReportOptions.chkSelectAxPos.value * 720 + frmReportOptions.chkSelectRPM.value * 720 + frmReportOptions.chkTRG.value * 720)

                    drCustCirc.Sections(2).Controls("labelTRG").Left = PosTRG
                    drCustCirc.Sections(3).Controls("textTRG").Left = PosTRG
                    drCustCirc.Sections(2).Controls("labelRemarks").Left = PosRem
                    drCustCirc.Sections(3).Controls("textRemarks").Left = PosRem
                    drCustCirc.Sections(2).Controls("labelRPM").Left = PosRPM
                    drCustCirc.Sections(3).Controls("textRPM").Left = PosRPM
                    drCustCirc.Sections(2).Controls("labelAxPos").Left = PosAxPos
                    drCustCirc.Sections(3).Controls("textAxPos").Left = PosAxPos
                    drCustCirc.Sections(2).Controls("labelCircFlow").Left = PosCircFlow
                    drCustCirc.Sections(3).Controls("textCircflow").Left = PosCircFlow
                    drCustCirc.Sections(2).Controls("labelSelectVibX").Left = PosVib
                    drCustCirc.Sections(3).Controls("textVibX").Left = PosVib
                    drCustCirc.Sections(2).Controls("labelSelectVibY").Left = PosVib + 720
                    drCustCirc.Sections(3).Controls("textVibY").Left = PosVib + 720

                    drCustCirc.Sections(2).Controls("labelTRG").Visible = PosTRG
                    drCustCirc.Sections(3).Controls("textTRG").Visible = PosTRG
                    drCustCirc.Sections(2).Controls("labelRPM").Visible = PosRPM
                    drCustCirc.Sections(3).Controls("textRPM").Visible = PosRPM
                    drCustCirc.Sections(2).Controls("labelAxPos").Visible = PosAxPos
                    drCustCirc.Sections(3).Controls("textAxPos").Visible = PosAxPos
                    drCustCirc.Sections(2).Controls("labelCircFlow").Visible = PosCircFlow
                    drCustCirc.Sections(3).Controls("textCircflow").Visible = PosCircFlow
                    drCustCirc.Sections(2).Controls("labelSelectVibX").Visible = PosVib
                    drCustCirc.Sections(3).Controls("textVibX").Visible = PosVib
                    drCustCirc.Sections(2).Controls("labelSelectVibY").Visible = PosVib
                    drCustCirc.Sections(3).Controls("textVibY").Visible = PosVib


        Case 2  'customer vibration
            Set dr = drCustVib
        Case 3  'internal
            Set dr = drInternal
            drInternal2.Sections(1).Controls("lblCustomer").Caption = txtShpNo
            drInternal2.Sections(1).Controls("lblBillTo").Caption = txtBilNo
            drInternal.Sections(1).Controls("lblBillTo").Caption = txtBilNo
            drInternal2.Sections(1).Controls("lblmodel").Caption = txtModelNo
            drInternal2.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
            drInternal2.Sections(1).Controls("lblSN").Caption = txtSN
            drInternal2.Sections(1).Controls("lblToday").Caption = Now
            drInternal2.Sections(1).Controls("lblVersion").Caption = strVersion
            drInternal2.Orientation = rptOrientLandscape
            drInternal2.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
            drInternal2.Sections(2).Controls("lblTC1").Caption = txtTitle(0).Text
            drInternal2.Sections(2).Controls("lblTC1A").Caption = txtTitle(1).Text
            drInternal2.Sections(2).Controls("lblTC2").Caption = txtTitle(2).Text
            drInternal2.Sections(2).Controls("lblTC2A").Caption = txtTitle(3).Text
            drInternal2.Sections(2).Controls("lblTC3").Caption = txtTitle(4).Text
            drInternal2.Sections(2).Controls("lblTC3A").Caption = txtTitle(5).Text
            drInternal2.Sections(2).Controls("lblTC4").Caption = txtTitle(6).Text
            drInternal2.Sections(2).Controls("lblTC4A").Caption = txtTitle(7).Text
            drInternal2.Sections(2).Controls("lblAI1").Caption = txtTitle(20).Text
            drInternal2.Sections(2).Controls("lblAI1A").Caption = txtTitle(21).Text
            drInternal2.Sections(2).Controls("lblAI2").Caption = txtTitle(22).Text
            drInternal2.Sections(2).Controls("lblAI2A").Caption = txtTitle(23).Text
            drInternal2.Sections(2).Controls("lblAI3").Caption = txtTitle(24).Text
            drInternal2.Sections(2).Controls("lblAI3A").Caption = txtTitle(25).Text
            drInternal2.Sections(2).Controls("lblAI4").Caption = txtTitle(26).Text
            drInternal2.Sections(2).Controls("lblAI4A").Caption = txtTitle(27).Text

        Case 5  'charts
            Set dr = drChart
            isChart = True
            drChart.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
            drChart.Sections(1).Controls("lblCustomer").Caption = txtShpNo
            drChart.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
            drChart.Sections(1).Controls("lblmodel").Caption = txtModelNo
            drChart.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
            drChart.Sections(1).Controls("lblSN").Caption = txtSN
            drChart.Sections(1).Controls("lblToday").Caption = Now
            drChart.Sections(1).Controls("lblVersion").Caption = strVersion

'            Set drChart.Sections(1).Controls("Image1").Picture = CWGraphKw.ControlImage
'            Set drChart.Sections(1).Controls("Image2").Picture = CWGraphEff.ControlImage
'            Set drChart.Sections(1).Controls("Image3").Picture = CWGraphTDH.ControlImage
'            Set drChart.Sections(1).Controls("Image4").Picture = CWGraphAmps.ControlImage

        Case 6  'escape out
            Exit Sub
        Case 7  'Unapproved pumps
            drApproved.Sections(1).Controls("lblNow").Caption = Now
            drApproved.Sections(1).Controls("lblVersion").Caption = strVersion
            drApproved.Show
            Exit Sub
        Case 4  'balance holes
            Set dr = drBalanceHoles
            If DataEnvironment3.Recordsets.Item(1).State = adStateOpen Then
                DataEnvironment3.Recordsets.Item(1).Close
            End If
            If (LenB(frmPLCData.txtSN.Text) = 0) Or (LenB(cmbTestDate.List(cmbTestDate.ListIndex)) = 0) Then
                Exit Sub
            End If
            DataEnvironment3.GetBalanceHoles frmPLCData.txtSN.Text, cmbTestDate.List(cmbTestDate.ListCount - 1)
'            DataEnvironment3.GetBalanceHoles frmPLCData.txtSN.Text, cmbTestDate.List(cmbTestDate.ListIndex)
            drBalanceHoles.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
            drBalanceHoles.Sections(1).Controls("lblCustomer").Caption = txtShpNo
            drBalanceHoles.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
            drBalanceHoles.Sections(1).Controls("lblmodel").Caption = txtModelNo
            drBalanceHoles.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
            drBalanceHoles.Sections(1).Controls("lblSN").Caption = txtSN
            drBalanceHoles.Sections(1).Controls("lblToday").Caption = Now
            drBalanceHoles.Sections(1).Controls("lblVersion").Caption = strVersion
            drBalanceHoles.Show
            Exit Sub

        Case 8  'export to excel
            ExportToExcel
            Exit Sub

        Case 9  'Customer no circ flow no axial
            Set dr = drCustNoCircNoAxial

        Case 10  'Customer vibration no axial
            Set dr = drCustVibNoAxial

        Case 11 ' TEMC Inspection Report
            Set dr = drTEMCInspectionSheet
    End Select

    If Index = 11 Then
        dr.Sections(1).Controls("lblOutlineDimensions").Caption = "Are the following dimensions in accordance with those shown on the outline drawing, with tolerances per procedure A-29368?" & Chr(13) & "*   Base anchor bolt hole bolt circle pitch" & Chr(13) & "*   Base anchor bolt hole diameter" & Chr(13) & "*   Suction and discharge flange bolt hole bolt circle pitch" & Chr(13) & "*   Suction and discharge flange bolt hole diameter" & Chr(13) & "*   Face-to-face dimension between suction and discharge flange faces" & Chr(13) & "*   Location dimensions for all other fluid connections"
        dr.Sections(1).Controls("lblMotorNoLoadTest").Caption = "Record motor input current and power: **_________A **_______kW" & Chr(13) & "Confirm test was completed"
        dr.Sections(1).Controls("lblMotorLockedRotorTest").Caption = "Record motor input power at rated current: **_________A **_______kW" & Chr(13) & "Confirm test was completed"
        dr.Sections(1).Controls("lblHydraulicTest").Caption = "Are hydraulic test results per A-15852 acceptable?  Specifically, are the following within acceptable limits?" & Chr(13) & "*   rated head at rated flow" & Chr(13) & "*   input power" & Chr(13) & "*   input current" & Chr(13) & "*   TRG reading" & Chr(13) & "*   Axial thrust force/P-V" & Chr(13) & "*   Vibration" & Chr(13) & "(Note: test data recorded through lab data acquisition system.)"
        dr.Sections(1).Controls("lblNPSHTest").Caption = "Is NPSH required at design flow no greater than the NPSH required value provided to the customer?" & Chr(13) & "(Note: test data recorded through lab data acquisition system.)"
        dr.Sections(1).Controls("lblCustomer").Caption = txtShpNo
        dr.Sections(1).Controls("lblJobNo").Caption = txtJobNum.Text
        If InStr(txtSalesOrderNumber.Text, "/") <> 0 Then
            dr.Sections(1).Controls("lblOrderNo").Caption = Left$(txtSalesOrderNumber.Text, InStr(txtSalesOrderNumber.Text, "/") - 1)
            dr.Sections(1).Controls("lblItemNo").Caption = Right$(txtSalesOrderNumber.Text, Len(txtSalesOrderNumber.Text) - InStr(txtSalesOrderNumber.Text, "/"))
        Else
            dr.Sections(1).Controls("lblOrderNo").Caption = txtSalesOrderNumber.Text
            dr.Sections(1).Controls("lblItemNo").Caption = "****"
        End If
        dr.Sections(1).Controls("lblProductNo").Caption = txtSN
        dr.Sections(1).Controls("lblType").Caption = txtModelNo
        dr.Sections(1).Controls("lblDateInspected").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
        dr.Sections(1).Controls("lblFreq").Caption = cmbFrequency.List(cmbFrequency.ListIndex)
        dr.Sections(1).Controls("lblVolts").Caption = cmbVoltage.List(cmbVoltage.ListIndex)
        dr.Sections(1).Controls("lblFt").Caption = Format(Val(txtDesignTDH) * 0.3048, "###,##0.00")
        dr.Sections(1).Controls("lblCapacityM").Caption = Format(Val(txtDesignFlow) * 0.2271247, "##,##0.00")
        dr.Sections(1).Controls("lblCapacityG").Caption = txtDesignFlow

        dr.Sections(1).Controls("lblInsulationResistance").Caption = Val(txtTestAndInspection(0)) & " V Megger " & Val(txtTestAndInspection(1)) & " MOhms Above "
        dr.Sections(1).Controls("lblDielectricStrength").Caption = "AC " & Val(txtTestAndInspection(2)) & " X  " & Val(txtTestAndInspection(3)) & " min "
        dr.Sections(1).Controls("lblHydrostaticTest").Caption = Val(txtTestAndInspection(4)) & " " & cmbTestAndInspection(0).Text & " X " & Val(txtTestAndInspection(5)) & " min "
        dr.Sections(1).Controls("lblPneumaticTest").Caption = Val(txtTestAndInspection(6)) & " " & cmbTestAndInspection(1).Text & " X " & Val(txtTestAndInspection(7)) & " min "

        For I = 1 To 4
            If TestAndInspectionGood(I - 1).value = 1 Then
                dr.Sections(1).Controls("lblGood" & Trim(str(I))).Caption = "Good _X__"
            Else
                dr.Sections(1).Controls("lblGood" & Trim(str(I))).Caption = "Good ____"
            End If
        Next I

        'print initials and date
        For I = 1 To 11
            If I < 11 Then  '11 is supervisor yes/no
                dr.Sections(1).Controls("lblInitials" & Trim(str(I))).Caption = txtWho.Text
                dr.Sections(1).Controls("lbldate" & Trim(str(I))).Caption = Date
            End If
            If TestAndInspectionGood(I + 3).value = 1 Then
                dr.Sections(1).Controls("lblYesNo" & Trim(str(I))).Caption = "Yes_X No__"
            Else
                dr.Sections(1).Controls("lblYesNo" & Trim(str(I))).Caption = "Yes__ No_X"
            End If
        Next I

        'fix
        dr.Sections(1).Controls("lblPhase").Caption = txtNoPhases.Text
        dr.Sections(1).Controls("lblNPSHReq").Caption = txtNPSHr.Text
        dr.Sections(1).Controls("lblRatedOutput").Caption = txtRatedInputPower.Text
        dr.Sections(1).Controls("lblLiquid").Caption = txtLiquid.Text
        dr.Sections(1).Controls("lblAmps").Caption = txtAmps.Text
        dr.Sections(1).Controls("lblThermalClass").Caption = txtThermalClass.Text
        dr.Sections(1).Controls("lblViscosity").Caption = txtViscosity.Text
        dr.Sections(1).Controls("lblEXPClass").Caption = txtExpClass.Text
        dr.Sections(1).Controls("lblLiquidTemp").Caption = txtLiquidTemperature.Text

        'fix
        If txtSN <> "" Then
            Select Case optMfr(0).value
                Case True
'                    dr.Sections(1).Controls("lblType").Caption = ""
                    dr.Sections(1).Controls("lblPole").Caption = IIf(cmbRPM.List(cmbRPM.ListIndex) = 3450, "2", "4")
                Case False
'                    dr.Sections(1).Controls("lblType").Caption = cmbTEMCModel.List(cmbTEMCModel.ListIndex)
                    dr.Sections(1).Controls("lblPole").Caption = IIf(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1) = "1", "2", "4")
            End Select
        End If
        dr.Sections(1).Controls("lblSG").Caption = txtSpGr

'        Set dr.Sections(2).Controls("Image2").Picture = CWGraphKw.ControlImage
'        Set dr.Sections(2).Controls("Image3").Picture = CWGraphEff.ControlImage
'        Set dr.Sections(2).Controls("Image4").Picture = CWGraphTDH.ControlImage
'        Set dr.Sections(2).Controls("Image5").Picture = CWGraphAmps.ControlImage

    End If

    If Not isChart And Index <> 11 Then
'        dr.Sections(1).Controls("lblTestDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)
        dr.Sections(1).Controls("lblCustomer").Caption = txtShpNo
        dr.Sections(1).Controls("lblmodel").Caption = txtModelNo
        dr.Sections(1).Controls("lblsono").Caption = txtSalesOrderNumber
        dr.Sections(1).Controls("lblSN").Caption = txtSN
        dr.Sections(1).Controls("lblGPM").Caption = txtDesignFlow
        dr.Sections(1).Controls("lblFt").Caption = txtDesignTDH
'        dr.Sections(1).Controls("lblBaroPress").Caption = txtInHgDisplay
        dr.Sections(1).Controls("lblSuctGageHt").Caption = txtSuctHeight
        dr.Sections(1).Controls("lblDischGageHt").Caption = txtDischHeight
        dr.Sections(1).Controls("lblVersion").Caption = strVersion

        If chkTrimmed.value = 1 Then
            If Val(txtImpTrim.Text) <> 0 Then
                dr.Sections(1).Controls("lblImpDia").Caption = txtImpTrim
            Else
                dr.Sections(1).Controls("lblImpDia").Caption = txtImpellerDia
            End If
        Else
            dr.Sections(1).Controls("lblImpDia").Caption = txtImpellerDia
        End If
        Dim stemp As String
        dr.Sections(1).Controls("lblRPM").Caption = IIf(optMfr(0).value = True, cmbRPM.List(cmbRPM.ListIndex), IIf(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1) = "1", "3450", "1750"))
        dr.Sections(1).Controls("lblSpGr").Caption = txtSpGr
        dr.Sections(1).Controls("lblMotor").Caption = IIf(optMfr(0).value = True, cmbMotor.List(cmbMotor.ListIndex), txtTEMCFrameNumber.Text)
'        dr.Sections(1).Controls("lblMotor").Caption = cmbMotor.List(cmbMotor.ListIndex)
        If Len(cmbTEMCVoltage.List(cmbTEMCVoltage.ListIndex)) > 6 Then
            stemp = Right$(cmbTEMCVoltage.List(cmbTEMCVoltage.ListIndex), Len(cmbTEMCVoltage.List(cmbTEMCVoltage.ListIndex)) - 6)
        Else
            stemp = ""
        End If
        dr.Sections(1).Controls("lblVoltage").Caption = IIf(optMfr(0).value = True, cmbVoltage.List(cmbVoltage.ListIndex), stemp)
        dr.Sections(1).Controls("lblEndPlay").Caption = txtEndPlay
        If Len(cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)) > 6 Then
            stemp = Right$(cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex), Len(cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)) - 6)
        Else
            stemp = ""
        End If
        dr.Sections(1).Controls("lblDesPressure").Caption = IIf(optMfr(0).value = True, cmbDesignPressure.List(cmbDesignPressure.ListIndex), stemp)
        dr.Sections(1).Controls("lblStatorFill").Caption = IIf(optMfr(0).value = True, cmbStatorFill.List(cmbStatorFill.ListIndex), "Dry")
        If Len(cmbTEMCModel.List(cmbTEMCModel.ListIndex)) > 11 Then
            stemp = Right$(cmbTEMCModel.List(cmbTEMCModel.ListIndex), Len(cmbTEMCModel.List(cmbTEMCModel.ListIndex)) - 11)
        Else
            stemp = ""
        End If

        dr.Sections(1).Controls("lblCircPath").Caption = IIf(optMfr(0).value = True, cmbCirculationPath.List(cmbCirculationPath.ListIndex), stemp)

        dr.Sections(1).Controls("lblKWMult").Caption = txtKWMult
        If Val(txtHDCor) = 0 Then
            dr.Sections(1).Controls("lblHDCor").Caption = 0
        Else
            dr.Sections(1).Controls("lblHDCor").Caption = txtHDCor
        End If
        dr.Sections(1).Controls("lblSuctPipeDia").Caption = cmbSuctDia.List(cmbSuctDia.ListIndex)
        dr.Sections(1).Controls("lblDischPipeDia").Caption = cmbDischDia.List(cmbDischDia.ListIndex)
        dr.Sections(1).Controls("lblTestSpec").Caption = cmbTestSpec.List(cmbTestSpec.ListIndex)

        dr.Sections(1).Controls("lblToday").Caption = Now

        dr.Sections(1).Controls("lblSuctionID").Caption = txtSuctionID
        dr.Sections(1).Controls("lblDischargeID").Caption = txtDischargeID
        dr.Sections(1).Controls("lblTempID").Caption = txtTemperatureID
        dr.Sections(1).Controls("lblCircFlowID").Caption = txtMagflowID
        dr.Sections(1).Controls("lblFlowID").Caption = txtFlowmeterID

        dr.Sections(1).Controls("lblAnalyzerID").Caption = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)
        dr.Sections(1).Controls("lblLoopID").Caption = cmbLoopNumber.List(cmbLoopNumber.ListIndex)
        dr.Sections(1).Controls("lblTachID").Caption = cmbTachID.List(cmbTachID.ListIndex)
        dr.Sections(1).Controls("lblOrificeID").Caption = cmbOrificeNumber.List(cmbOrificeNumber.ListIndex)
        dr.Sections(1).Controls("lblRunDate").Caption = cmbTestDate.List(cmbTestDate.ListIndex)

        If chkFeathered.value = 1 Then
            dr.Sections(1).Controls("lblImpFeathered").Visible = True
        Else
            dr.Sections(1).Controls("lblImpFeathered").Visible = False
        End If

        If chkOrifice.value = 1 Then
            dr.Sections(1).Controls("lblDischOrifice").Visible = True
            dr.Sections(1).Controls("lblDischOrificeValue").Visible = True
            dr.Sections(1).Controls("lblDischOrificeValue").Caption = txtOrifice
        Else
            dr.Sections(1).Controls("lblDischOrifice").Visible = False
            dr.Sections(1).Controls("lblDischOrificeValue").Visible = False
        End If

        If chkCircOrifice.value = 1 Then
            dr.Sections(1).Controls("lblCircFlowOrifice").Visible = True
            dr.Sections(1).Controls("lblCircFlowOrificeValue").Visible = True
            dr.Sections(1).Controls("lblCircFlowOrificeValue").Caption = frmPLCData.txtCircOrifice
        Else
            dr.Sections(1).Controls("lblCircFlowOrifice").Visible = False
            dr.Sections(1).Controls("lblCircFlowOrificeValue").Visible = False
        End If

        dr.Sections(4).Controls("lblOther").Caption = txtOtherMods
        dr.Sections(4).Controls("lblRemarks").Caption = txtRemarks

        dr.Sections(4).Controls("lblTestRunRemarks").Caption = txtTestSetupRemarks

        dr.Orientation = rptOrientLandscape
    End If

    Printer.Orientation = vbPRORLandscape

'    If DataEnvironment1.rsEff.State = adStateOpen Then
'        DataEnvironment1.rsEff.Close
'    End If
'    DataEnvironment1.rsEff.Open
'    DataEnvironment1.rsEff.MoveFirst

'    DataEnvironment1.Commands.Item(1).Parameters(0).Value = 7

'    DataEnvironment1.rsEff.Requery

    Dim dm As String
    Dim RsNum As Integer
    rsEff.Requery
    RsNum = 9 - UpDown2.value

    dm = "Recs" & Trim(str(frmPLCData.UpDown2.value))

    Select Case Index
        Case 0
            Set drCustNoCirc.DataSource = DataEnvironment1
            drCustNoCirc.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            dr.Sections(4).Controls("lblBH").Caption = strBH
            DoEvents
            drCustNoCirc.Show
        Case 1
            Set drCustCirc.DataSource = DataEnvironment1
            drCustCirc.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            dr.Sections(4).Controls("lblBH").Caption = strBH
            DoEvents
            drCustCirc.Show
        Case 2
            Set drCustVib.DataSource = DataEnvironment1
            drCustVib.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            dr.Sections(4).Controls("lblBH").Caption = strBH
            DoEvents
            drCustVib.Show
        Case 3
            Set drInternal.DataSource = DataEnvironment1
            drInternal.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            dr.Sections(4).Controls("lblBH").Caption = strBH
            DoEvents
            drInternal.Show
            Set drInternal2.DataSource = DataEnvironment1
            drInternal2.DataMember = dm
            DoEvents
            drInternal2.Show
        Case 4
            Set drAnalysis.DataSource = DataEnvironment1
            drAnalysis.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            DoEvents
            drAnalysis.Show
        Case 5
            Set drChart.DataSource = DataEnvironment1
            drChart.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            DoEvents
            drChart.Show
        Case 6

        Case 9
            Set drCustNoCircNoAxial.DataSource = DataEnvironment1
            drCustNoCircNoAxial.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            dr.Sections(4).Controls("lblBH").Caption = strBH
            DoEvents
            drCustNoCircNoAxial.Show
        Case 10
            Set drCustVibNoAxial.DataSource = DataEnvironment1
            drCustVibNoAxial.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
            dr.Sections(4).Controls("lblBH").Caption = strBH
            DoEvents
            drCustVibNoAxial.Show
        Case 11
            Printer.Orientation = vbPRORLandscape
            Printer.PaperSize = vbPRPSLetter
            drTEMCInspectionSheet.Orientation = rptOrientPortrait
            drTEMCInspectionSheet.Height = 16000
            drTEMCInspectionSheet.RightMargin = 0
            drTEMCInspectionSheet.LeftMargin = 0
            drTEMCInspectionSheet.TopMargin = 0
            drTEMCInspectionSheet.BottomMargin = 0
            Set drTEMCInspectionSheet.DataSource = DataEnvironment1
            drTEMCInspectionSheet.DataMember = dm
            DoEvents
            If DataEnvironment1.Recordsets(RsNum).State = adStateClosed Then
                DataEnvironment1.Recordsets(RsNum).Open
            End If
            DataEnvironment1.Recordsets(RsNum).Requery
'            dr.Sections(4).Controls("lblBH").Caption = strBH
            DoEvents

            drTEMCInspectionSheet.Show
    End Select

'    rsEff.Close
'    rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

'    rsEff.Requery
  
End Sub

Private Sub tmrGetDDE_Timer()

'get here every second... get plc and magtrol data

    Dim sSendStr As String
    Dim I As Integer
    Dim VoltMul As Double

    If Calibrating Then
        Exit Sub
    End If

    If debugging Then
        'Exit Sub
    End If

    If boPLCOperating = True Then
        frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on
        DoEvents

        'convert the plc data into real numbers
        'the following data are type real
        txtFlow.Text = ConvertToReal("4050")
        txtSuction.Text = ConvertToReal("4052")
        txtDischarge.Text = ConvertToReal("4054")
        txtTemperature.Text = ConvertToReal("4056")

        txtValvePosition.Text = ConvertToLong("2004")

        frmPLCData.txtTC1.Text = ConvertToLong("2200")
        frmPLCData.txtTC2.Text = ConvertToLong("2202")
        frmPLCData.txtTC3.Text = ConvertToLong("2204")
        frmPLCData.txtTC4.Text = ConvertToLong("2206")

        frmPLCData.txtAI1.Text = ConvertToReal("4060")
        frmPLCData.txtAI2.Text = ConvertToReal("4062")
        frmPLCData.txtAI3.Text = ConvertToReal("4064")
        frmPLCData.txtAI4.Text = ConvertToReal("4066")

        frmPLCData.txtPCoef.Text = ConvertToLong("4036")
        frmPLCData.txtICoef.Text = ConvertToLong("4037")
        frmPLCData.txtDCoef.Text = ConvertToLong("4040")

        frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
        frmPLCData.txtInHg.Text = ConvertToLong("1460")


        'modify the data from PLC format to format that we can use
        'and update the screen
        If txtFlowDisplay.Enabled = False Then
            frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
        End If
        If txtSuctionDisplay.Enabled = False Then
            frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
        End If
        If txtDischargeDisplay.Enabled = False Then
            frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
        End If
        If txtTemperatureDisplay.Enabled = False Then
            frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
        End If
        frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

        frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
        frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
        frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
        frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

        If txtAI1Display.Enabled = False Then
            frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
        End If
        If txtAI2Display.Enabled = False Then
            frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
        End If
        If txtAI3Display.Enabled = False Then
            frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
        End If
        If txtAI4Display.Enabled = False Then
            frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
        End If

        frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

        frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

        frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off
        DoEvents

        frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
        DoEvents
    End If

    If boMagtrolOperating = True Then


        'get the data from the Magtrol
        If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
            sSendStr = vbCrLf
            sData = Space$(68)
            VoltMul = Sqr(3)
        Else
            sSendStr = "OT" & vbCrLf
            sData = Space$(183)
            VoltMul = 1#
        End If
        ibwrt iUD, sSendStr
        ibrd iUD, sData

        'parse the Magrol response
        'vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)

            Dim vSplit() As String
            vSplit = Split(Right(sData, Len(sData) - 1), ",")
            If UBound(vSplit) > 0 Then
                ReDim vResponse(UBound(vSplit))
            End If
            For I = 0 To UBound(vSplit) - 1
                If Len(vSplit(I)) <> 0 Then
                    vResponse(I) = CDbl(vSplit(I))
                End If
            Next I

        'format the parsed response
        Dim dd As String
        dd = "- -"

        On Error GoTo noresponse
        If Not IsEmpty(vResponse) Then
        '8 entries for 5300 and 12 for the 6530
            If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                'put the responses into the correct text box
                txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

                Select Case vResponse(0)
                    Case Is < 1
                        txtI1.Text = Format$(vResponse(0), "0.0000")
                    Case Is < 10
                        txtI1.Text = Format$(vResponse(0), "0.000")
                    Case Is < 100
                        txtI1.Text = Format$(vResponse(0), "00.00")
                    Case Else
                        txtI1.Text = Format$(vResponse(0), "000.0")
                End Select

                Select Case vResponse(3)
                    Case Is < 1
                        txtI2.Text = Format$(vResponse(3), "0.0000")
                    Case Is < 10
                        txtI2.Text = Format$(vResponse(3), "0.000")
                    Case Is < 100
                        txtI2.Text = Format$(vResponse(3), "00.00")
                    Case Else
                        txtI2.Text = Format$(vResponse(3), "000.0")
                End Select

                Select Case vResponse(6)
                    Case Is < 1
                        txtI3.Text = Format$(vResponse(6), "0.0000")
                    Case Is < 10
                        txtI3.Text = Format$(vResponse(6), "0.000")
                    Case Is < 100
                        txtI3.Text = Format$(vResponse(6), "00.00")
                    Case Else
                        txtI3.Text = Format$(vResponse(6), "000.0")
                End Select

                txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
                txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                'txtI2.Text = Format$(vResponse(3), "###0.0")
                txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
                txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                'txtI3.Text = Format$(vResponse(6), "###0.0")
                txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
                If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                    'if we have some measured current
                    'pf = sum of power/sum of VA
                    If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                        'add kw responses and / by 1000 to get to kW
                        txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
                        txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
                    Else
                        txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
                        txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
                    End If
                    Select Case Val(txtKW.Text)
                        Case Is < 1
                            txtKW.Text = Format$(txtKW.Text, "0.00000")
                        Case Is < 10
                            txtKW.Text = Format$(txtKW.Text, "0.0000")
                        Case Is < 100
                            txtKW.Text = Format$(txtKW.Text, "00.000")
                        Case Else
                            txtKW.Text = Format$(txtKW.Text, "000.00")
                    End Select
                Else
                    txtPF = dd
                End If
            Else
                'no response, show all -- in text boxes
                txtV1.Text = dd
                txtI1.Text = dd
                txtP1.Text = dd
                txtV2.Text = dd
                txtI2.Text = dd
                txtP2.Text = dd
                txtV3.Text = dd
                txtI3.Text = dd
                txtP3.Text = dd
                txtPF = dd
                txtKW = dd
            End If
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
        End If
    Else    'magtrol not operating
        Dim dbl As Double

        If optKW(0).value = True Then   'add 3 powers
            txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
        End If
        If optKW(1).value = True Then   'enter kw
            txtP1.Text = Val(txtKW.Text) / 3
            txtP2.Text = Val(txtKW.Text) / 3
            txtP3.Text = Val(txtKW.Text) / 3
        End If
        If optKW(2).value = True Then   'use ai4
            txtKW.Text = txtAI4Display.Text
            txtP1.Text = Val(txtKW.Text) / 3
            txtP2.Text = Val(txtKW.Text) / 3
            txtP3.Text = Val(txtKW.Text) / 3
        End If

        dbl = Val(txtV1.Text) * Val(txtI1.Text)
        dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
        dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
        If dbl <> 0 Then
            txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
        End If
    End If

noresponse:
On Error GoTo 0
    frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off
    DoEvents

    'update the little PLC chart
    For I = 0 To 19
        vPlot(I, 0) = vPlot(I + 1, 0)
        vPlot(I, 1) = vPlot(I + 1, 1)
    Next I
    vPlot(20, 0) = Val(txtSetPointDisplay.Text)
    vPlot(20, 1) = Val(txtFlowDisplay.Text)

'    If Not (txtSetPointDisplay.Text = "" Or txtFlowDisplay.Text = "") Then
'       If vPlot(0, 0) <> Empty Then
            MSChart2 = vPlot

            MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
            MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(vPlot) / 10) + 0.5) + 1)
            MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
            MSChart2.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5

'            CWGraph1.PlotY vPlot
'        End If
'    End If

    'do NPSH stuff
    Dim SuctVelHead As Single
    Dim DischVelHead As Single
    Dim Conversion As Single
    Dim SuctionPSIA As Single
    Dim DischargePSIA As Single
    Dim VaporPress As Single
    Dim SpecVolume As Single
    Dim NPSHa As Single
    Dim TDH As Single
    Dim pd As Single


    'velocity head
    If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
        SuctVelHead = 0
    Else
'        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
        pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
        SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
    End If

    If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
        DischVelHead = 0
    Else
'        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
        pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
        DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
    End If

    'convert gauges to absolute
    If txtInHgDisplay.Text = "" Then
        Conversion = 0
    Else
        Conversion = txtInHgDisplay * 0.491
    End If

    SuctionPSIA = Val(txtSuctionDisplay) + Conversion
    DischargePSIA = Val(txtDischargeDisplay) + Conversion


    'lookup vapor pressure and specific volume in the arrays that we made
    'if temp is out of range, say so and exit
    If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
        txtNPSHa = 0
        Exit Sub
    Else
        I = Val(txtTemperatureDisplay) - 40
'        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
'        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
        VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
        SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
    End If

    If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
        'NPSHa
        NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
'        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
        txtNPSHa = Format$(NPSHa, "##0.00")

        'tdh
        TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
        txtTDH = Format$(TDH, "##0.00")

    Else
        txtNPSHa = 0
    End If
End Sub
Private Sub tmrStartUp_Timer()
    'we waited for a while, disable the timer
    tmrStartUp.Enabled = False
End Sub
Public Function SetCombo(cmbComboName As ComboBox, sName As String, rs As ADODB.Recordset)
'set the pump parameter combo box to the right data based upon
'the number in the database

    Dim I As Integer
    Dim sParam As String
    Dim qy As New ADODB.Command
    Dim rs1 As New ADODB.Recordset

    If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
        sParam = rs.Fields(sName)                'get the index number
        qy.ActiveConnection = cnPumpData
        qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
        Set rs1 = qy.Execute()                                  'get the record for the index number

        If rs1.BOF = True And rs1.EOF = True Then
            cmbComboName.ListIndex = -1                             'else, remove any pointer
            Exit Function
        End If

        For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
            If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
                cmbComboName.ListIndex = I                                              'if we do, set the combo box
                Exit For                                            'and we're done
            End If
            cmbComboName.ListIndex = -1                             'else, remove any pointer
        Next I
    Else
        cmbComboName.ListIndex = -1
    End If

    Exit Function
End Function
Private Function SetComboTestSetup(cmbComboName As ComboBox, sFieldName As String, sTableName As String, rs As ADODB.Recordset)
'set the pump parameter combo box to the right data based upon
'the number in the database

'same as setcombo, except here we also pass in the field name

    FromStoredData = True

    Dim I As Integer
    Dim sParam As String
    Dim qy As New ADODB.Command
    Dim rs1 As New ADODB.Recordset

    If rs.Fields(sFieldName).ActualSize <> 0 Then
        sParam = rs.Fields(sFieldName)
        qy.ActiveConnection = cnPumpData
        qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
        Set rs1 = qy.Execute()

        For I = 0 To cmbComboName.ListCount - 1
            If cmbComboName.ItemData(I) = rs1.Fields(0) Then
                cmbComboName.ListIndex = I
                Exit For
            End If
            cmbComboName.ListIndex = -1
        Next I
    Else
        cmbComboName.ListIndex = -1
    End If

    FromStoredData = False

    Exit Function
End Function
Private Sub DisablePumpDataControls()
    'disable the pump data controls cause we're just showing what we found

    txtSalesOrderNumber.Enabled = False
    frmMfr.Enabled = False
    txtShpNo.Enabled = False
    txtBilNo.Enabled = False
    txtDesignFlow.Enabled = False
    txtDesignTDH.Enabled = False

    frmMiscPumpData.Enabled = False

    txtModelNo.Enabled = False
    txtImpellerDia.Enabled = False

    frmTEMC.Enabled = False
    frmChempump.Enabled = False

    txtRemarks.Enabled = False
    Me.cmdAddNewTestDate.Visible = False

    cmdEnterPumpData.Enabled = False
  
End Sub
Private Sub DisableTestSetupDataControls()

    cmbTestSpec.Enabled = False
    txtWho.Enabled = False
    txtRMA.Enabled = False

    frmLoopAndXducer.Enabled = False
    frmElecData.Enabled = False
    frmPerfMods.Enabled = False
    frmOtherFiles.Enabled = False
    frmInstrumentTags.Enabled = False
    frmTAndI.Enabled = False
    frmThrustBalMods.Enabled = False
    txtTestSetupRemarks.Enabled = False

    cmdEnterTestSetupData.Enabled = False
    cmbPLCNo.Enabled = False
End Sub
Private Sub DisableTestDataControls()

    cmbPLCLoop.Enabled = False
    frmPumpData.Enabled = False
    frmThermocouples.Enabled = False
    frmAI.Enabled = False
    frmMagtrol.Enabled = False
    fmrMiscTestData.Enabled = False
    frmPLCMisc.Enabled = False
    DataGrid1.Enabled = False
    DataGrid2.Enabled = False
    cmdEnterTestData.Enabled = False
  
End Sub
Private Sub EnableTestSetupDataControls()

    cmbTestSpec.Enabled = True
    txtWho.Enabled = True
    txtRMA.Enabled = True

    frmLoopAndXducer.Enabled = True
    frmElecData.Enabled = True
    frmPerfMods.Enabled = True
    frmOtherFiles.Enabled = True
    frmInstrumentTags.Enabled = True
    frmTAndI.Enabled = True
    frmThrustBalMods.Enabled = True
    txtTestSetupRemarks.Enabled = True

    cmdEnterTestSetupData.Enabled = True
    cmbPLCNo.Enabled = True
End Sub
Private Sub EnableTestDataControls()

    cmbPLCLoop.Enabled = True
    frmPumpData.Enabled = True
    frmThermocouples.Enabled = True
    frmAI.Enabled = True
    frmMagtrol.Enabled = True
    fmrMiscTestData.Enabled = True
    frmPLCMisc.Enabled = True
    DataGrid1.Enabled = True
    DataGrid2.Enabled = True
    cmdEnterTestData.Enabled = True
  
End Sub
Private Sub EnablePumpDataControls()
    'disable the pump data controls cause we're just showing what we found

    txtSalesOrderNumber.Enabled = True
    frmMfr.Enabled = True
    txtShpNo.Enabled = True
    txtBilNo.Enabled = True
    txtDesignFlow.Enabled = True
    txtDesignTDH.Enabled = True

    frmMiscPumpData.Enabled = True

    txtModelNo.Enabled = True
    txtImpellerDia.Enabled = True

    frmTEMC.Enabled = True
    frmChempump.Enabled = True

    txtRemarks.Enabled = True
    Me.cmdAddNewTestDate.Visible = True

    cmdEnterPumpData.Enabled = True
  
End Sub
Private Sub EnableMagtrolFields()
    txtV1.Enabled = True
    txtV2.Enabled = True
    txtV3.Enabled = True
    txtI1.Enabled = True
    txtI2.Enabled = True
    txtI3.Enabled = True
    txtP1.Enabled = True
    txtP2.Enabled = True
    txtP3.Enabled = True
    optKW(0).Visible = True
    optKW(1).Visible = True
    optKW(2).Visible = True
    optKW(1).value = True
    optKW_Click (1)
End Sub
Private Sub DisableMagtrolFields()
    txtV1.Enabled = False
    txtV2.Enabled = False
    txtV3.Enabled = False
    txtI1.Enabled = False
    txtI2.Enabled = False
    txtI3.Enabled = False
    txtP1.Enabled = False
    txtP2.Enabled = False
    txtP3.Enabled = False
    txtKW.Enabled = False
    optKW(0).Visible = False
    optKW(1).Visible = False
    optKW(2).Visible = False
End Sub
Private Sub EnablePLCFields()
    frmPLCData.txtAI1Display.Enabled = True
    frmPLCData.txtAI2Display.Enabled = True
    frmPLCData.txtAI3Display.Enabled = True
    frmPLCData.txtAI4Display.Enabled = True
    frmPLCData.txtTC1Display.Enabled = True
    frmPLCData.txtTC2Display.Enabled = True
    frmPLCData.txtTC3Display.Enabled = True
    frmPLCData.txtTC4Display.Enabled = True
    frmPLCData.txtFlowDisplay.Enabled = True
    frmPLCData.txtSuctionDisplay.Enabled = True
    frmPLCData.txtDischargeDisplay.Enabled = True
    frmPLCData.txtTemperatureDisplay.Enabled = True
    frmPLCData.txtInHgDisplay.Enabled = True
End Sub
Private Sub DisablePLCFields()
    frmPLCData.txtAI1Display.Enabled = False
    frmPLCData.txtAI2Display.Enabled = False
    frmPLCData.txtAI3Display.Enabled = False
    frmPLCData.txtAI4Display.Enabled = False
    frmPLCData.txtTC1Display.Enabled = False
    frmPLCData.txtTC2Display.Enabled = False
    frmPLCData.txtTC3Display.Enabled = False
    frmPLCData.txtTC4Display.Enabled = False
    frmPLCData.txtFlowDisplay.Enabled = False
    frmPLCData.txtSuctionDisplay.Enabled = False
    frmPLCData.txtDischargeDisplay.Enabled = False
    frmPLCData.txtTemperatureDisplay.Enabled = False
    frmPLCData.txtInHgDisplay.Enabled = False
End Sub
Private Sub BlankData()
    txtShpNo.Text = vbNullString
    txtBilNo.Text = vbNullString
    txtModelNo.Text = vbNullString
    cmbMotor.ListIndex = -1
    cmbStatorFill.ListIndex = -1
    cmbVoltage.ListIndex = -1
    cmbDesignPressure.ListIndex = -1
    cmbFrequency.ListIndex = -1
    cmbCirculationPath.ListIndex = -1
    cmbRPM.ListIndex = -1
    cmbModel.ListIndex = -1
    cmbModelGroup.ListIndex = -1
    txtSpGr.Text = vbNullString
    txtImpellerDia.Text = vbNullString
    txtEndPlay.Text = vbNullString
    txtGGap.Text = vbNullString
    txtDesignFlow.Text = vbNullString
    txtDesignTDH.Text = vbNullString
    txtOtherMods.Text = vbNullString
    txtRemarks.Text = vbNullString
    txtSalesOrderNumber.Text = vbNullString
    txtTestSetupRemarks.Text = vbNullString
    txtNPSHFile.Text = vbNullString
    txtPicturesFile.Text = vbNullString
    txtVibrationFile.Text = vbNullString
'    cmbOrificeNumber.ListIndex = 18
'    cmbTestSpec.ListIndex = 6       'default = Rev7
    cmbLoopNumber.ListIndex = -1
    cmbSuctDia.ListIndex = -1
    cmbDischDia.ListIndex = -1
    cmbTachID.ListIndex = -1
    cmbAnalyzerNo.ListIndex = -1
    txtTestRemarks.Text = vbNullString
    txtHDCor.Text = 0
    txtDischHeight.Text = 0
    txtSuctHeight.Text = 0
    txtKWMult.Text = 1
    txtWho.Text = LogInInitials
    txtRMA.Text = vbNullString
    frmPLCData.chkNPSH.value = 0
    frmPLCData.chkPictures.value = 0
    frmPLCData.chkVibration.value = 0
    frmPLCData.txtFlowmeterID = vbNullString
    frmPLCData.txtSuctionID = vbNullString
    frmPLCData.txtDischargeID = vbNullString
    frmPLCData.txtTemperatureID = vbNullString
    frmPLCData.txtMagflowID = vbNullString
    frmPLCData.chkBalanceHoles.value = 0
    frmPLCData.chkCircOrifice = 0
    frmPLCData.txtCircOrifice = vbNullString
    frmPLCData.txtImpTrim = vbNullString
    frmPLCData.txtOrifice = vbNullString
    frmPLCData.chkFeathered.value = 0
    frmPLCData.chkTrimmed.value = 0
    frmPLCData.chkCircOrifice.value = 0
    frmPLCData.txtThrustBal = vbNullString
    frmPLCData.txtRPM = vbNullString
    frmPLCData.txtVibAx = vbNullString
    frmPLCData.txtVibRad = vbNullString
    frmPLCData.txtTEMCTRGReading = vbNullString
    dgBalanceHoles.Visible = False
End Sub
Private Sub AddTestData()
    Dim I As Integer
    Dim sFilter As String

    ClearEff
    rsEff.MoveFirst

    For I = 1 To 8
        rsTestData.AddNew
        rsTestData!SerialNumber = txtSN
        rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
        rsTestData!testnumber = I
        rsTestData!DataWritten = False
        rsTestData.Update
        DoEfficiencyCalcs
        rsEff.MoveNext
        rsTestData.MoveNext
    Next I
    boFoundTestData = True
    'rsTestData.Update
    rsTestData.Requery
    rsTestData.Resync

   'select the entries from testdata
    sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

    rsTestData.Filter = sFilter

    Set DataGrid1.DataSource = rsTestData

    ' fix the datagrid

    Dim c As Column
    For Each c In DataGrid1.Columns
       Select Case c.DataField
       Case "TestDataID"
          c.Visible = False
       Case "SerialNumber"
          c.Visible = False
       Case "Date"
          c.Visible = False
       Case Else ' Hide all other columns.
          c.Visible = True
          c.Alignment = dbgRight
       End Select
    Next c

    rsEff.Requery

    DataGrid1.Refresh
    DataGrid2.Refresh

   ' fix the datagrid
'   Set DataGrid1.DataSource = rsTestData
'   Set DataGrid2.DataSource = rsEff



'    ClearEff
End Sub
Private Sub DoEfficiencyCalcs()
    Dim KW As Single, VI As Single, VITemp As Single
    Dim Vave As Single, Iave As Single
    Dim I As Integer
    Dim j As Integer
    Dim HeightDiff As Single

    If Not IsNull(rsTestData.Fields("TotalPower")) Then
        KW = rsTestData.Fields("TotalPower")
    Else
        'if we wrote data with an old version, we will not have written total power
        'if total power = 0 and the three individual powers are not 0, add them

        If rsTestData.Fields("PowerA") > 0 Then
            If rsTestData.Fields("PowerB") > 0 Then
                If rsTestData.Fields("PowerC") > 0 Then
                    KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
                End If
            End If
        End If
   End If

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

    I = 0
    Vave = 0
    Iave = 0
    If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
        VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
        Vave = rsTestData.Fields("VoltageA")
        Iave = rsTestData.Fields("CurrentA")
        If VI <> 0 Then
            I = I + 1
        End If
    End If
    If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
        VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
        If VITemp <> 0 Then
            I = I + 1
            VI = VI + VITemp
            Vave = Vave + rsTestData.Fields("VoltageB")
            Iave = Iave + rsTestData.Fields("CurrentB")
        End If
    End If
    If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
        VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
        If VITemp <> 0 Then
            I = I + 1
            VI = VI + VITemp
            Vave = Vave + rsTestData.Fields("VoltageC")
            Iave = Iave + rsTestData.Fields("CurrentC")
        End If
    End If
    If KW = 0 Then
        For j = 1 To rsEff.Fields.Count - 1
            rsEff.Fields(j) = 0
        Next j
'        Exit Sub
    End If
    If VI <> 0 Then
        rsEff.Fields("Volts") = Vave / I
        rsEff.Fields("Amps") = Iave / I
        rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
        rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
    Else
        rsEff.Fields("PowerFactor") = 0
    End If

    If optMfr(0).value = True Then
        If cmbStatorFill.ListIndex = -1 Then
            rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

        Else
            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
'            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
        End If
    Else
        rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
    End If

    Dim sHDCor As Single
    Dim sDisc As Single
    Dim sSuct As Single
    If IsNull(rsTestSetup.Fields("HDCor")) Then
        sHDCor = 0
    Else
        sHDCor = rsTestSetup.Fields("HDCor")
    End If
    If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
        sDisc = 0
    Else
        sDisc = rsTestSetup.Fields("DischargeGageHeight")
    End If
    If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
        sSuct = 0
    Else
        sSuct = rsTestSetup.Fields("SuctionGageHeight")
    End If
    HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
    If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
        rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
    End If
'    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
    rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
    rsEff.Fields("ElecHP") = 1000 * KW / 746
'    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
        If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
            If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
    '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
            rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
    '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
            rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
            If rsEff.Fields("MotorEfficiency") <> 0 Then
                rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
            Else
                rsEff.Fields("HydraulicEfficiency") = 0
            End If
        Else
            rsEff.Fields("LiquidHP") = 0
            rsEff.Fields("OverallEfficiency") = 0
        End If
    Else
        rsEff.Fields("LiquidHP") = 0
        rsEff.Fields("OverallEfficiency") = 0
    End If

    I = rsEff.AbsolutePosition
    If Not IsNull(rsTestData.Fields("Flow")) Then
        rsEff.Fields("Flow") = rsTestData.Fields("Flow")
        HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
        HeadFlow(1, I - 1) = rsEff.Fields("TDH")
        FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
        FlowHead(I - 1, 1) = rsEff.Fields("TDH")

        EffFlow(0, I - 1) = rsTestData.Fields("Flow")
        EffFlow(1, I - 1) = rsEff.Fields("OverallEfficiency")
        KWFlow(0, I - 1) = rsTestData.Fields("Flow")
        KWFlow(1, I - 1) = KW
        AmpsFlow(0, I - 1) = rsTestData.Fields("Flow")
        AmpsFlow(1, I - 1) = rsEff.Fields("Amps")
    Else
        HeadFlow(0, I - 1) = 0
        HeadFlow(1, I - 1) = 0

        EffFlow(0, I - 1) = 0
        EffFlow(1, I - 1) = 0
        KWFlow(0, I - 1) = 0
        KWFlow(1, I - 1) = 0
        AmpsFlow(0, I - 1) = 0
        AmpsFlow(1, I - 1) = 0
    End If

    Dim Plothead(7, 1) As Single
    Dim HeadPlot(7, 1) As Single
'    Dim PlotEff() As Single
'    Dim PlotKW() As Single
'    Dim PlotAmps() As Single
'    ReDim PlotHead(0, 0)
'    ReDim PlotEff(0, 0)
'    ReDim PlotKW(0, 0)
'
    For j = 0 To UpDown2.value - 1
'        If HeadFlow(1, j) <> 0 Then
            'ReDim Preserve Plothead(1, j)
            'ReDim Preserve HeadPlot(j, 1)
'            Plothead(0, j) = HeadFlow(0, j)
'            Plothead(1, j) = HeadFlow(1, j)
            HeadPlot(j, 0) = FlowHead(j, 0)
            HeadPlot(j, 1) = FlowHead(j, 1)
            Plothead(j, 0) = FlowHead(j, 0)
            Plothead(j, 1) = FlowHead(j, 1)

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
    Next j

    MSChart1 = Plothead
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((SetGraphMax(Plothead) / 10) + 0.5) + 1)
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5



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
    rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
    rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
'    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
'    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
    rsEff.Fields("KW") = KW
    rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
    rsEff.Fields("RPM") = rsTestData.Fields("RPM")
    rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
    rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
    rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
    rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
    rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
    rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
    rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
    rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
    rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
    rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
    rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
    rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
    rsEff.Fields("TC1") = rsTestData.Fields("TC1")
    rsEff.Fields("TC2") = rsTestData.Fields("TC2")
    rsEff.Fields("TC3") = rsTestData.Fields("TC3")
    rsEff.Fields("TC4") = rsTestData.Fields("TC4")
    rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
    rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
    rsEff.Fields("AI4") = rsTestData.Fields("AI4")
    rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
    rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
    rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
    rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
    rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
    rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
    rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
    If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
        txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
    End If
    If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
        txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
    End If
    If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
        txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
    End If
    If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
        txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
    End If
    If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
        txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
    End If

    CalculateTEMCForce

    If Not IsNull(txtTEMCCalcForce.Text) Then
        rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
    Else
        rsEff.Fields("TEMCCalculatedForce") = 0
    End If

    If Not IsNull(txtTEMCPVValue.Text) Then
        rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
    Else
        rsEff.Fields("TEMCPV") = 0
    End If

    If Val(txtTEMCFrontThrust.Text) <> 0 Then
        rsEff.Fields("TEMCFR") = "F"
'        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
    Else
        If Val(txtTEMCRearThrust.Text) = 0 Then
            'no thrust
            rsEff.Fields("TEMCFR") = " "
            rsEff.Fields("TEMCFrontThrust") = 0
        Else
            rsEff.Fields("TEMCFR") = "R"
'            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
        End If
    End If

    rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

    rsEff.Update

    If rsEffDisp.State = adStateOpen Then
        rsEffDisp.Close
    End If

    Dim qyEffDisp As New ADODB.Command
    qyEffDisp.ActiveConnection = cnEffData
    qyEffDisp.CommandText = "SELECT Flow, TDH, KW, Volts, Amps, OverallEfficiency FROM Efficiency;"

    With rsEffDisp     'open the recordset for the query
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open qyEffDisp
    End With

    rsEffDisp.Requery
    DataGrid2.Refresh

  
End Sub
Private Sub ClearEff()
'    Dim I As Integer, j As Integer
    Dim qy As New ADODB.Command

    If rsEff.State = adStateOpen Then
        If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
            rsEff.CancelUpdate
        End If
        rsEff.Close
    End If
    qy.ActiveConnection = cnEffData
    qy.CommandText = "DROP TABLE Efficiency"
    rsEff.Open qy
    qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
    rsEff.Open qy
    rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

    rsEff.Requery
    DataGrid2.Refresh

    Dim c As Column
    For Each c In DataGrid2.Columns
        c.Alignment = dbgCenter
        c.Width = 750
        Select Case c.ColIndex
            Case 1
                c.Caption = "Flow"
                c.NumberFormat = "###0.00"
            Case 2
                c.Caption = "TDH"
                c.NumberFormat = "00.0"
            Case 3
                c.Caption = "Overall Eff"
                c.NumberFormat = "00.00"
                c.Width = 850
            Case 4
                c.Caption = "PF"
                c.NumberFormat = "00.0"
            Case 5
                c.Caption = "Vel Head"
                c.NumberFormat = "00.00"
            Case 6
                c.Caption = "Elec HP"
                c.NumberFormat = "#00.0"
            Case 7
                c.Caption = "Liq HP"
                c.NumberFormat = "#00.0"
            Case Else
                c.Visible = False
        End Select
    Next c
  
End Sub
Function JustAlphaNumeric(char As String) As String
    Select Case Asc(char)
        Case 42             ' *
            JustAlphaNumeric = char
        Case 48 To 57       ' 0 - 9
            JustAlphaNumeric = char
        Case 65 To 90       ' A - Z
            JustAlphaNumeric = char
        Case 97 To 122      ' a - z
            JustAlphaNumeric = UCase(char)
        Case Else
            JustAlphaNumeric = ""
    End Select
End Function



Private Sub txtI1_Change()
    txtI2.Text = txtI1.Text
    txtI3.Text = txtI1.Text
End Sub

Private Sub txtModelNo_Change()
    Dim I As Integer
    Dim S As String
    Dim sFull As String
    Dim boDone As Boolean
    Dim boRepeat As Boolean

    Static bo3Digits As Boolean         '3 digits in frame number
    Static bo2Digits As Boolean         '2 digits in stages

    If optMfr(0).value = True Then
        Exit Sub
    End If

    cmbTEMCAdapter.ListIndex = -1
    cmbTEMCAdditions.ListIndex = -1
    cmbTEMCCirculation.ListIndex = -1
    cmbTEMCDesignPressure.ListIndex = -1
    cmbTEMCNominalDischargeSize.ListIndex = -1
    cmbTEMCDivisionType.ListIndex = -1
    cmbTEMCImpellerType.ListIndex = -1
    cmbTEMCInsulation.ListIndex = -1
    cmbTEMCJacketGasket.ListIndex = -1
    cmbTEMCMaterials.ListIndex = -1
    cmbTEMCModel.ListIndex = -1
    cmbTEMCNominalImpSize.ListIndex = -1
    cmbTEMCOtherMotor.ListIndex = -1
    cmbTEMCPumpStages.ListIndex = -1
    cmbTEMCNominalSuctionSize.ListIndex = -1
    cmbTEMCTRG.ListIndex = -1
    cmbTEMCVoltage.ListIndex = -1


    'first, get rid of spaces, dashes, etc

    S = ""
    For I = 1 To Len(txtModelNo.Text)
        S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
    Next I

    'next, fill out the model number to it's max length of 24 characters

    boDone = False
    boRepeat = False

    Do While Not boDone
        sFull = ""
        For I = 1 To Len(S)
            Select Case I
                Case 1
                    'type
                    sFull = sFull & Mid$(S, I, 1)
                Case 2
                    'adapter
                    If IsNumeric(Mid$(S, I, 1)) Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 3
                    'materials
                    sFull = sFull & Mid$(S, I, 1)
                Case 4
                'design pressure
                    sFull = sFull & Mid$(S, I, 1)
                Case 5
                'motor frame number - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 6
                'motor frame number - digit 2
                    sFull = sFull & Mid$(S, I, 1)
                Case 7
                'motor frame number - digit 3
                    sFull = sFull & Mid$(S, I, 1)
                Case 8
                'motor frame number - digit 4
                    If IsNumeric(Mid$(S, I, 1)) Then
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    Else    '3 digits
'                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
                        S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
                        boRepeat = True
                        Exit For
                    End If
                Case 9
                'insulation
                    sFull = sFull & Mid$(S, I, 1)
                Case 10
                'voltage
                    sFull = sFull & Mid$(S, I, 1)
                Case 11
                'other motor specs
                    If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 12
                ' TRG
                    sFull = sFull & Mid$(S, I, 1)
                Case 13
                'Nominal discharge - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 14
                'nominal discharge - digit 2
                    sFull = sFull & Mid$(S, I, 1)
                Case 15
                'nominal suction - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 16
                'nominal suction - digit 2
                    sFull = sFull & Mid$(S, I, 1)
                Case 17
                'nominal impeller size
                    sFull = sFull & Mid$(S, I, 1)
                Case 18
                'impeller type
                    If Val(Mid$(sFull, 5, 1)) < 3 Then
                        If IsNumeric(Mid$(S, I, 1)) Or Mid$(S, I, 1) = "*" Then
                            sFull = sFull & Mid$(S, I, 1)
                            boRepeat = False
                        Else
                            S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                            boRepeat = True
                            Exit For
                        End If
                    Else
                        If Mid$(S, I, 1) = "*" Then
                            boRepeat = False
                        Else
                            S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                            boRepeat = True
                            Exit For
                        End If
                    End If
                Case 19
                'Division type
                    If IsNumeric(Mid$(S, I, 1)) Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 20
                'pump stages - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 21
                'pump stages - digit 2
                    If IsNumeric(Mid$(S, I, 1)) Or Mid$(S, I, 1) = "*" Then
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    Else
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    End If
                Case 22
                'pump jacket
                    If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                               Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 23
                'additions
                      sFull = sFull & Mid$(S, I, 1)
                Case 24
                'circulation
                      sFull = sFull & Mid$(S, I, 1)
            End Select
        Next I
        If Not boRepeat Then
            boDone = True
        End If
    Loop

    sFull = S
    For I = 1 To Len(sFull)
        Select Case I
            Case 1
                ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
            Case 2
                ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
            Case 3
                ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
            Case 4
                ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
            Case 5
'                If IsNumeric(Mid$(sFull, i, 1)) Then  '4 digit frame number
                    If Val(Mid$(sFull, I, 1)) = 0 Then
                        txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
                    Else
                        txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
                    End If
'                    bo3Digits = False
'                Else
'                    txtTEMCFrameNumber.Text = Mid$(sFull, 5, 3)
'                    ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, i, 1)
'                    bo3Digits = True
'                End If
            Case 9
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, i, 1)
'                Else
                    ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
'                End If
            Case 10
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, i, 1)
'                Else
                    ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
'                End If
            Case 11
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, i, 1)
'                Else
                    ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
'                End If
            Case 12
'                If bo3Digits Then
'                Else
                    ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
'                End If
            Case 13
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCNominalDischargeSize, Right$(sFull, 2)
'                Else
'                End If
                    ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
            Case 14
'                If bo3Digits Then
'                Else
'                End If
            Case 15
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCNominalSuctionSize, Right$(sFull, 2)
'                Else
'                End If
                    ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
            Case 16
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, i, 1)
'                Else
'                End If
            Case 17
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, i, 1)
'                Else
                    ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
'                End If
            Case 18
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, i, 1)
'                Else
                    ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
'                End If
            Case 19
'                If bo3Digits Then
'                    ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, i, 1)
'                Else
                    ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
'                End If
            Case 20
'                If bo3Digits Then
'                    If IsNumeric(Mid$(sFull, i, 1)) Then  '2 digit stages
'                        ParseTEMCModelNo cmbTEMCPumpStages, Right$(sFull, 2)
'                        bo2Digits = True
'                    Else
'                        ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, i, 1)
'                        bo2Digits = False
'                    End If
'                Else
                    If IsNumeric(Mid$(sFull, I + 1, 1)) Then
                        ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 2)
                    Else
                        ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
                    End If
'                End If
            Case 21
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
            Case 22
'                If bo3Digits Then
'                    If bo2Digits Then
'                        ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, i, 1)
'                    Else
'                        ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, i, 1)
'                    End If
'                Else
'                    If bo2Digits Then
                        ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
'                    Else
'                        ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, i, 1)
'                    End If
'                End If
            Case 23
'                If bo3Digits Then
'                    If bo2Digits Then
'                        ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, i, 1)
'                    Else
'                    End If
'                Else
'                    If bo2Digits Then
                        ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
'                    Else
'                        ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, i, 1)
'                    End If
'                End If
            Case 24
'                If bo3Digits Then
'                    If bo2Digits Then
'                    Else
'                    End If
'                Else
'                    If bo2Digits Then
                        ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)
'                    Else
'                    End If
'                End If

        End Select
    Next I
End Sub

Private Sub txtModelNo_Validate(Cancel As Boolean)
    Dim I As Integer
    Dim S As String

'    s = txtModelNo.Text
'    S = Replace(S, "-", "")
'    S = Replace(S, " ", "")
'    S = Replace(S, "/", "")

'    txtModelNo.Text = ""

'    For i = 1 To Len(s)
'        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
'    Next i
    txtModelNo_Change
  
End Sub

Private Sub txtNPSHFile_GotFocus()
    On Error GoTo FileCancel
    If LenB(txtNPSHFile.Text) <> 0 Then
        CommonDialog1.filename = txtNPSHFile.Text
    End If
    CommonDialog1.ShowOpen
    txtNPSHFile.Text = CommonDialog1.filename
    Exit Sub
FileCancel:
On Error GoTo 0
    CommonDialog1.CancelError = False
End Sub

Private Sub txtP1_Change()
    txtP2.Text = txtP1.Text
    txtP3.Text = txtP1.Text
End Sub

Private Sub txtPicturesFile_gotfocus()
    CommonDialog1.CancelError = True
    On Error GoTo FileCancel
    If LenB(txtPicturesFile.Text) <> 0 Then
        CommonDialog1.filename = txtPicturesFile.Text
    End If
    CommonDialog1.ShowOpen
    txtPicturesFile.Text = CommonDialog1.filename
    Exit Sub
FileCancel:
On Error GoTo 0
    CommonDialog1.CancelError = False
End Sub

Private Sub txtSN_Change()
    cmdFindPump.Default = True
End Sub

Private Sub txtTEMCFrontThrust_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCMomentArm_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCRearThrust_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCThrustRigPressure_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCViscosity_Change()
    CalculateTEMCForce
End Sub

Private Sub txtV1_Change()
    txtV2.Text = txtV1.Text
    txtV3.Text = txtV1.Text
End Sub

Private Sub txtVibrationFile_gotfocus()
    On Error GoTo FileCancel
    If LenB(txtVibrationFile.Text) <> 0 Then
        CommonDialog1.filename = txtVibrationFile.Text
    End If
    CommonDialog1.ShowOpen
    txtVibrationFile.Text = CommonDialog1.filename
    Exit Sub
FileCancel:
On Error GoTo 0
    CommonDialog1.CancelError = False
End Sub

Private Sub ExportToExcel()

    Me.UpDown2.value = Val(Me.txtUpDn2.Text)
    Dim SaveFileName As String
    Dim WorkSheetName As String

    Dim I As Integer
    Dim iRowNo As Integer
    Dim sImp As String
    Dim ans As Integer

    Dim bCanShowSpeed As Boolean
    Dim CantShowReason As String

'close any running excel processes
    Dim objWMIService, colProcesses
    Set objWMIService = GetObject("winmgmts:")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
    If colProcesses.Count > 0 Then
        Set xlApp = Excel.Application
    Else
        'use existing copy
'        Set xlApp = New Excel.Application
        Set xlApp = CreateObject("Excel.Application")
    End If


    CommonDialog1.CancelError = True        'in case the user
    On Error GoTo ErrHandler                '  chooses the cancel button

    'set up dialog box
    CommonDialog1.DialogTitle = "Open Excel Files"
    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
    CommonDialog1.InitDir = App.Path
'    CommonDialog1.InitDir = "C:\"    'in this directory
    CommonDialog1.ShowOpen                              'open the file selection dialog box

    If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
        SaveFileName = CommonDialog1.filename           'get the name of the file
        If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
             xlApp.Workbooks.Close
        End If
        ' Create the Excel Workbook Object.
On Error GoTo 0
        Set xlBook = xlApp.Workbooks.Add                'add a workbook
        WorkSheetName = NewWorkBook                     'do some stuff for the new workbook
        ActiveWorkbook.CheckCompatibility = False
        xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                                 FileFormat:=xlNormal                        'save the file
    Else                                                'the file name already exists
        SaveFileName = CommonDialog1.filename
        ' Create the Excel Workbook Object.
        If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
             xlApp.Workbooks.Close
        End If
        Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
        If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
            MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
            Exit Sub
        Else
        End If
    End If

On Error GoTo 0

    'see if we can export Speed and SG and if we can, ask user if s/he wants it
    'assume that we can show speed calcs

    bCanShowSpeed = False
'open the template and copy the data from the sheet
'  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
    'write the data to the spreadsheet
    With xlApp

    Dim xlTemplateName As String
    xlTemplateName = ParentDirectoryName & "\PumpData Excel Template.xls"
    Dim xlTemplate As Excel.Workbook
    Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
    Dim TemplateWS As Excel.Worksheet
    Dim sheetName As String
    sheetName = xlTemplate.Sheets(1).Name
    xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

    xlTemplate.Close savechanges:=False

    Set xlTemplate = Nothing

    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(WorkSheetName).Delete
    Application.DisplayAlerts = True
    ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

    'WorkSheetName = sheetName

    'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
    ' if there is, get the coefficients, else make the coefficients 0

        Dim ACoef As Double
        Dim BCoef As Double
        Dim CCoef As Double

        Dim qy As New ADODB.Command
        Dim rs As New ADODB.Recordset
        qy.ActiveConnection = cnPumpData
'        Dim VoltageForLookup As Integer
'        If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
'            VoltageForLookup = 460
'        ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
'            VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
'        End If

        If rsPumpData!ChempumpPump = True Then
            qy.CommandText = "SELECT * FROM CalculatedRPM WHERE Frame = '" & cmbMotor.List(cmbMotor.ListIndex) & "'"
        Else
            qy.CommandText = "SELECT * FROM CalculatedRPM WHERE Frame = '" & txtTEMCFrameNumber.Text & "'"
        End If

        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Open qy

        'if temc pump and not found, try right 3 digits of frame
        If rs.RecordCount = 0 And rsPumpData!ChempumpPump = False Then
            rs.Close
            qy.CommandText = "SELECT * FROM CalculatedRPM WHERE Frame = '" & Right(txtTEMCFrameNumber.Text, 3) & "'"
            rs.Open qy
        End If

        If rs.RecordCount = 0 Then
            ACoef = 0
            BCoef = 0
            CCoef = 0
            MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                          " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                          " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
        Else
            ACoef = rs.Fields("x2")
            BCoef = rs.Fields("x")
            CCoef = rs.Fields("b")
            .Range("H8").Select
            .ActiveCell.FormulaR1C1 = rs.Fields("Poles")
            .Range("H54").Select
            .ActiveCell.FormulaR1C1 = rs.Fields("Rotor OD")
            .Range("H55").Select
            .ActiveCell.FormulaR1C1 = rs.Fields("Rotor Length")
        End If


    'write header data
        'first write the revision
        Dim RundownRev As String
        RundownRev = App.Major & "." & App.Minor & "." & App.Revision

        .Range("AM3").Select
        .ActiveCell.FormulaR1C1 = RundownRev

        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Serial Number"
        .Range("C2").Select
        .ActiveCell.FormulaR1C1 = txtSN

        .Range("F1").Select
        .ActiveCell.FormulaR1C1 = "Customer"
        .Range("H1").Select
        .ActiveCell.FormulaR1C1 = txtShpNo

        .Range("A3").Select
        .ActiveCell.FormulaR1C1 = "Model"
        .Range("C3").Select
        .ActiveCell.FormulaR1C1 = txtModelNo

        .Range("F2").Select
        .ActiveCell.FormulaR1C1 = "Sales Order"
        .Range("H2").Select
        .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

        .Range("A9").Select
        .ActiveCell.FormulaR1C1 = "Design Flow"
        .Range("C9").Select
        .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

        .Range("A10").Select
        .ActiveCell.FormulaR1C1 = "Design Head"
        .Range("C10").Select
        .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

        .Range("A12").Select
        .ActiveCell.FormulaR1C1 = "Specific Heat"
        .Range("C12").Select
        .ActiveCell.FormulaR1C1 = Val(txtSpHeat.Text)

        .Range("P13").Select
        .ActiveCell.FormulaR1C1 = "Barometric Pressure"
        .Range("R13").Select
        .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

        .Range("P11").Select
        .ActiveCell.FormulaR1C1 = "Suction Gage Height"
        .Range("R11").Select
        .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

        .Range("P12").Select
        .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
        .Range("R12").Select
        .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = "Run Date"
        .Range("C1").Select
        .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

        .Range("D10:E10").Select
        With xlApp.Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        xlApp.Selection.Merge

        'determine rpm

        Dim RPMvalue As String
        If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
        '1 says 2 pole
            If Me.cmbFrequency.ListIndex = 0 Then
                '0 says 50Hz
                RPMvalue = "2900"
            ElseIf Me.cmbFrequency.ListIndex = 1 Then
                ' says 60Hz
                RPMvalue = "3450"
            Else
                'vfd or other, no rpm
                RPMvalue = ""
            End If
        Else
        '2 says 4 pole
            If Me.cmbFrequency.ListIndex = 0 Then
                '0 says 50Hz
                RPMvalue = "1450"
            ElseIf Me.cmbFrequency.ListIndex = 1 Then
                ' says 60Hz
                RPMvalue = "1750"
            Else
                'vfd or other, no rpm
                RPMvalue = ""
            End If
        End If

'        .Range("G1").Select
'        .ActiveCell.FormulaR1C1 = "RPM"
'        .Range("I1").Select
'        .ActiveCell.FormulaR1C1 = RPMvalue

        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Sp Gravity"
        .Range("C5").Select
        .ActiveCell.FormulaR1C1 = txtSpGr

        .Range("A6").Select
        .ActiveCell.FormulaR1C1 = "Viscosity"
        .Range("C6").Select
        .ActiveCell.FormulaR1C1 = txtViscosity

        .Range("F4").Select
        .ActiveCell.FormulaR1C1 = "Motor"
        .Range("H4").Select
        If rsPumpData!ChempumpPump = True Then
            .ActiveCell.FormulaR1C1 = Me.cmbMotor.List(Me.cmbMotor.ListIndex)
        Else
            .ActiveCell.FormulaR1C1 = Me.txtTEMCFrameNumber.Text
        End If

        .Range("H12").Select
'        .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text

        .Range("F5").Select
        .ActiveCell.FormulaR1C1 = "Voltage"
        .Range("H5").Select
        .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

        .Range("K6").Select
        .ActiveCell.FormulaR1C1 = "End Play"
        .Range("M6").Select
        .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

        .Range("K7").Select
        .ActiveCell.FormulaR1C1 = "G-Gap"
        .Range("M7").Select
        .ActiveCell.FormulaR1C1 = txtGGap.Text

        .Range("A8").Select
        .ActiveCell.FormulaR1C1 = "Design Pressure"
        .Range("C8").Select

        If rsPumpData!ChempumpPump = False Then
            Dim DesPress As String
            DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
            Dim j As Integer
            j = InStrRev(DesPress, "-")
            .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)
        Else
            .ActiveCell.FormulaR1C1 = Me.cmbDesignPressure.List(Me.cmbDesignPressure.ListIndex)
        End If

'        .Range("G8").Select
'        .ActiveCell.FormulaR1C1 = "Stator Fill"
'        .Range("I8").Select
'        .ActiveCell.FormulaR1C1 = "Dry"

        .Range("K4").Select
        .ActiveCell.FormulaR1C1 = "Circulation Path"

        .Range("M4").Select
        If rsPumpData!ChempumpPump = False Then
            .ActiveCell.FormulaR1C1 = Me.cmbTEMCModel.List(Me.cmbTEMCModel.ListIndex)
        Else
            .ActiveCell.FormulaR1C1 = Me.cmbCirculationPath.List(Me.cmbCirculationPath.ListIndex)
        End If

        .Range("M8").Select
        .ActiveCell.FormulaR1C1 = txtNPSHr.Text

        .Range("K1").Select
        .ActiveCell.FormulaR1C1 = "Impeller Dia"
        .Range("M1").Select


'        If LenB(txtImpTrim) <> 0 Then
'            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
'        Else
'            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
'        End If
'
        If chkTrimmed.value = 1 Then
            If Val(txtImpTrim.Text) <> 0 Then
                .ActiveCell.FormulaR1C1 = txtImpTrim
            Else
                .ActiveCell.FormulaR1C1 = txtImpellerDia
            End If
        Else
            .ActiveCell.FormulaR1C1 = txtImpellerDia
        End If



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

        .Range("P9").Select
        .ActiveCell.FormulaR1C1 = "Suction Dia"
        .Range("R9").Select
        .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

        .Range("P10").Select
        .ActiveCell.FormulaR1C1 = "Discharge Dia"
        .Range("R10").Select
        .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

        .Range("A11").Select
        .ActiveCell.FormulaR1C1 = "Test Spec"
        .Range("C11").Select
        .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

        .Range("K3").Select
        .ActiveCell.FormulaR1C1 = "Impeller Feathered"
        .Range("M3").Select
        If chkFeathered.value = 1 Then
            .ActiveCell.FormulaR1C1 = "Yes"
        Else
            .ActiveCell.FormulaR1C1 = "No"
        End If

        .Range("K2").Select
        .ActiveCell.FormulaR1C1 = "Disch Orifice"
        .Range("M2").Select
        If chkOrifice.value = 1 Then
            .ActiveCell.FormulaR1C1 = Val(txtOrifice)
        Else
            .ActiveCell.FormulaR1C1 = "None"
        End If


        .Range("K5").Select
        .ActiveCell.FormulaR1C1 = "Circulation Orifice"
        .Range("M5").Select
        If chkCircOrifice.value = 1 Then
            .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
        Else
            .ActiveCell.FormulaR1C1 = "None"
        End If

        .Range("A13").Select
        .ActiveCell.FormulaR1C1 = "Other Mods"
        .Range("C13").Select
        .ActiveCell.FormulaR1C1 = txtOtherMods

        .Range("A14").Select
        .ActiveCell.FormulaR1C1 = "Remarks"
        .Range("C14").Select
        .ActiveCell.FormulaR1C1 = txtRemarks

        .Range("A15").Select
        .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
        .Range("C15").Select
        .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

        .Range("P1").Select
        .ActiveCell.FormulaR1C1 = "Suct ID"
        .Range("R1").Select
        .ActiveCell.FormulaR1C1 = Me.txtSuctionID.Text

        .Range("P2").Select
        .ActiveCell.FormulaR1C1 = "Disch ID"
        .Range("R2").Select
        .ActiveCell.FormulaR1C1 = Me.txtDischargeID.Text

        .Range("P3").Select
        .ActiveCell.FormulaR1C1 = "Temp ID"
        .Range("R3").Select
        .ActiveCell.FormulaR1C1 = Me.txtTemperatureID.Text
        .Range("P4").Select
        .ActiveCell.FormulaR1C1 = "Circ Flow ID"
        .Range("R4").Select
        .ActiveCell.FormulaR1C1 = Me.txtMagflowID.Text

        .Range("P5").Select
        .ActiveCell.FormulaR1C1 = "Flow ID"
        .Range("R5").Select
        .ActiveCell.FormulaR1C1 = Me.txtFlowmeterID.Text 'cmbFlowMeter.List(cmbFlowMeter.ListIndex)

        .Range("P6").Select
        .ActiveCell.FormulaR1C1 = "Analyzer ID"
        .Range("R6").Select
        .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

        .Range("P7").Select
        .ActiveCell.FormulaR1C1 = "Loop ID"
        .Range("R7").Select
        .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Fluid"
        .Range("C4").Select
        .ActiveCell.FormulaR1C1 = txtLiquid.Text

        .Range("F3").Select
        .ActiveCell.FormulaR1C1 = "RMA"
        .Range("H3").Select
        .ActiveCell.FormulaR1C1 = Me.txtRMA.Text

        .Range("F12").Select
        .ActiveCell.FormulaR1C1 = "No Of Diodes"
        .Range("H12").Select
        .ActiveCell.FormulaR1C1 = Me.txtNoOfDiodes.Text

'        .ActiveCell.FormulaR1C1 = txtRMA.Text
'        If rsPumpData.Fields("RVSPartNo") <> "" Then
'            .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
'        End If
'        If rsPumpData.Fields("CustPN") <> "" Then
'            .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
'        End If

        .Range("A7").Select
        .ActiveCell.FormulaR1C1 = "Temperature"
        .Range("C7").Select
        .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

        .Range("F6").Select
        .ActiveCell.FormulaR1C1 = "Frequency"
        .Range("H6").Select
        If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
            .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
        Else
            .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
        End If
'        .Range("K2").Select
'        .ActiveCell.FormulaR1C1 = "Disch Orifice"
'        .Range("M2").Select
'        .ActiveCell.FormulaR1C1 = txtOrifice.Text

'        .Range("K12").Select
'        .ActiveCell.FormulaR1C1 = "Flow Orifice"
'        .Range("L12").Select
'        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

        .Range("P8").Select
        .ActiveCell.FormulaR1C1 = "PLC No"
        .Range("R8").Select
        .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

        .Range("F7").Select
        .ActiveCell.FormulaR1C1 = "Phases"
        .Range("H7").Select
        .ActiveCell.FormulaR1C1 = txtNoPhases.Text

        .Range("F8").Select
        .ActiveCell.FormulaR1C1 = "Poles"

        .Range("F9").Select
        .ActiveCell.FormulaR1C1 = "Rated Current"
        .Range("H9").Select
        .ActiveCell.FormulaR1C1 = txtAmps.Text

        .Range("F10").Select
        .ActiveCell.FormulaR1C1 = "Rated Input Power"
        .Range("H10").Select
        .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

        .Range("F11").Select
        .ActiveCell.FormulaR1C1 = "Insulation Class"
        .Range("H11").Select
        .ActiveCell.FormulaR1C1 = txtThermalClass.Text

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

        .Range("A17").Select
        .ActiveCell.FormulaR1C1 = "Flow"
        .Range("A18").Select
        .ActiveCell.FormulaR1C1 = "(GPM)"

        .Range("B17").Select
        .ActiveCell.FormulaR1C1 = "TDH"
        .Range("B18").Select
        .ActiveCell.FormulaR1C1 = "(Ft)"

        .Range("C17").Select
        .ActiveCell.FormulaR1C1 = "KW"

        .Range("D17").Select
        .ActiveCell.FormulaR1C1 = "Ave"
        .Range("D18").Select
        .ActiveCell.FormulaR1C1 = "Volts"

        .Range("E17").Select
        .ActiveCell.FormulaR1C1 = "Ave"
        .Range("E18").Select
        .ActiveCell.FormulaR1C1 = "Amps"

        .Range("F17").Select
        .ActiveCell.FormulaR1C1 = "Power"
        .Range("F18").Select
        .ActiveCell.FormulaR1C1 = "Factor"

        .Range("G17").Select
        .ActiveCell.FormulaR1C1 = "Overall"
        .Range("G18").Select
        .ActiveCell.FormulaR1C1 = "Eff"

        .Range("H17").Select
        .ActiveCell.FormulaR1C1 = "Measured"
        .Range("H18").Select
        .ActiveCell.FormulaR1C1 = "RPM"

        .Range("I17").Select
        .ActiveCell.FormulaR1C1 = "Calculated"
        .Range("I18").Select
        .ActiveCell.FormulaR1C1 = "RPM"

        .Range("J17").Select
        .ActiveCell.FormulaR1C1 = "Suction"
        .Range("J18").Select
        .ActiveCell.FormulaR1C1 = "Temp(F)"

        .Range("K17").Select
        .ActiveCell.FormulaR1C1 = "Disch"
        .Range("K18").Select
        .ActiveCell.FormulaR1C1 = "Pressure"

        .Range("L17").Select
        .ActiveCell.FormulaR1C1 = "Suction"
        .Range("L18").Select
        .ActiveCell.FormulaR1C1 = "Pressure"

        .Range("M17").Select
        .ActiveCell.FormulaR1C1 = "Vel"
        .Range("M18").Select
        .ActiveCell.FormulaR1C1 = "Head"

        .Range("N17").Select
        .ActiveCell.FormulaR1C1 = "Axial"
        .Range("N18").Select
        .ActiveCell.FormulaR1C1 = "Position"

        .Range("O17").Select
        .ActiveCell.FormulaR1C1 = "Pct of"
        .Range("O18").Select
        .ActiveCell.FormulaR1C1 = "End Play"

        .Range("P17").Select
        .ActiveCell.FormulaR1C1 = "Hydraulic"
        .Range("P18").Select
        .ActiveCell.FormulaR1C1 = "Efficiency"

'        .Range("P17").Select
'        .ActiveCell.FormulaR1C1 = "Circ"
'        .Range("P18").Select
'        .ActiveCell.FormulaR1C1 = "Flow"

        .Range("Q17").Select
        .ActiveCell.FormulaR1C1 = "Motor"
        .Range("Q18").Select
        .ActiveCell.FormulaR1C1 = "Efficiency"

        .Range("S17").Select
        .ActiveCell.FormulaR1C1 = "NPSHa"

        .Range("T17").Select
        .ActiveCell.FormulaR1C1 = "Phase 1"
        .Range("T18").Select
        .ActiveCell.FormulaR1C1 = "Current"

        .Range("U17").Select
        .ActiveCell.FormulaR1C1 = "Phase 2"
        .Range("U18").Select
        .ActiveCell.FormulaR1C1 = "Current"

        .Range("V17").Select
        .ActiveCell.FormulaR1C1 = "Phase 3"
        .Range("V18").Select
        .ActiveCell.FormulaR1C1 = "Current"

        .Range("W17").Select
        .ActiveCell.FormulaR1C1 = "Phase 1"
        .Range("W18").Select
        .ActiveCell.FormulaR1C1 = "Voltage"

        .Range("X17").Select
        .ActiveCell.FormulaR1C1 = "Phase 2"
        .Range("X18").Select
        .ActiveCell.FormulaR1C1 = "Voltage"

        .Range("Y17").Select
        .ActiveCell.FormulaR1C1 = "Phase 3"
        .Range("Y18").Select
        .ActiveCell.FormulaR1C1 = "Voltage"

        .Range("Z17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

        .Range("Z18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

        .Range("AA17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

        .Range("AA18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

        .Range("AB17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

        .Range("AB18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

'        .Range("AC17").Select
'        .ActiveCell.FormulaR1C1 = "HR"

'        .Range("AC18").Select
'        .ActiveCell.FormulaR1C1 = "(ft)"

        .Range("AC17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

        .Range("AC18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

        .Range("AD17").Select
        .ActiveCell.FormulaR1C1 = "TRG"
        .Range("AD18").Select
        .ActiveCell.FormulaR1C1 = "Position"

        .Range("AE17").Select
        .ActiveCell.FormulaR1C1 = "Thrust"

        .Range("AF17").Select
        .ActiveCell.FormulaR1C1 = "F/R"

        .Range("AG17").Select
        .ActiveCell.FormulaR1C1 = "Moment"
        .Range("AG18").Select
        .ActiveCell.FormulaR1C1 = "Arm"

        .Range("AH17").Select
        .ActiveCell.FormulaR1C1 = "Rig"
        .Range("AH18").Select
        .ActiveCell.FormulaR1C1 = "Pressure"

'        .Range("AI17").Select
'        .ActiveCell.FormulaR1C1 = "Viscosity"

        .Range("AI19").Select
        .ActiveCell.FormulaR1C1 = "Rear"
        .Range("AI18").Select
        .ActiveCell.FormulaR1C1 = "Force"

        .Range("AJ17").Select
        .ActiveCell.FormulaR1C1 = "PV"

        .Range("R17").Select
        .ActiveCell.FormulaR1C1 = "Shaft"
        .Range("R18").Select
        .ActiveCell.FormulaR1C1 = "Power"

'        .Range("AM17").Select
'        .ActiveCell.FormulaR1C1 = "Pct Full"
'        .Range("AM18").Select
'        .ActiveCell.FormulaR1C1 = "Scale"

        .Range("AK17").Select
        .ActiveCell.FormulaR1C1 = "NPSHr"

        .Range("AL17").Select
        .ActiveCell.FormulaR1C1 = "Remarks"




        'now output the data

        iRowNo = 20

        rsEff.MoveFirst
        For I = 1 To frmPLCData.UpDown2.value
            .Range("A" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

            .Range("B" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

            .Range("C" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

            .Range("D" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

            .Range("E" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

            .Range("F" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

            .Range("G" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

            .Range("H" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

            .Range("I" & iRowNo).Select
            'use the coefficients from above to calculate rpm
            Dim f As Double
            f = .Range("H6").value
            If f <> 60 Then
                .ActiveCell.FormulaR1C1 = 0
            Else
                .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)
            End If

            .Range("J" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

            .Range("K" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

            .Range("L" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

            .Range("M" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

            .Range("N" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

            .Range("O" & iRowNo).Select
            If Val(txtEndPlay) > 0 Then
                .ActiveCell.FormulaR1C1 = 100 * rsEff.Fields("Pos") / Val(txtEndPlay)
            End If

            .Range("P" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

'            .Range("P" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

            .Range("Q" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

            .Range("S" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

            .Range("T" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

            .Range("U" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

            .Range("V" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

            .Range("W" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

            .Range("X" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

            .Range("Y" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

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

            .Range("Z" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

            .Range("AA" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

            .Range("AB" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

'            .Range("AC" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

            .Range("AC" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

            .Range("AD" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

            .Range("AE" & iRowNo).Select

            If rsEff.Fields("TEMCFrontThrust") = 0 Then
                If rsEff.Fields("TEMCRearThrust") = 0 Then
                    .ActiveCell.FormulaR1C1 = " "
                    .Range("AF" & iRowNo).Select
                    .ActiveCell.FormulaR1C1 = " "
                Else
                    .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
                    .Range("AF" & iRowNo).Select
                    .ActiveCell.FormulaR1C1 = "R"
                End If
            Else
                .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
                .Range("AF" & iRowNo).Select
                .ActiveCell.FormulaR1C1 = "F"
            End If

            .Range("AG" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

            .Range("AH" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

'            .Range("AJ" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

            .Range("AI" & iRowNo).Select
            If rsEff.Fields("TEMCForceDirection") = "F" Then
                .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
            Else
                .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
            End If

            .Range("AJ" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

            .Range("R" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

            .Range("AK" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")

'            If RatedKW = 999 Then
'                .ActiveCell.FormulaR1C1 = ""
'            Else
'                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
'            End If

            .Range("AL" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


            rsEff.MoveNext
            iRowNo = iRowNo + 1
        Next I

        .Range("AL20:AL57").Select
        .Selection.WrapText = False

        .Range("C13:C15").Select
        .Selection.WrapText = False

        .Range("A20:AS30").Select
        .Selection.NumberFormat = "0.00"

        .Range("N20:N27").Select
        .Selection.NumberFormat = "0.000"

    'set up formulas to calculate BEP
    '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
    '  the formulas for doing that are in E68, F68 and G68
    '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
    '
    Dim AColumnRow As String
    Dim PColumnRow As String

    AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
    PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

        .Range("E68").Select
'        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

        .Range("F68").Select
'        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

        .Range("G68").Select
'        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

    'export balance holes
    If boGotBalanceHoles Then
        If rsBalanceHoles.State = adStateClosed Then
            rsBalanceHoles.ActiveConnection = cnPumpData
            rsBalanceHoles.Open
        End If 'rsBalanceHoles.State = adStateClosed

        If rsBalanceHoles.RecordCount <> 0 Then

            .Range("K9:N9").Merge
            .Range("K9:N9").Formula = "Balance Hole Data"
            .Range("K9:N9").HorizontalAlignment = xlCenter

            .Range("K10").Select
            .ActiveCell.Formula = "Date"

            .Range("L10").Select
            .ActiveCell.Formula = "Number"

            .Range("M10").Select
            .ActiveCell.Formula = "Diameter"

            .Range("N10").Select
            .ActiveCell.Formula = "Bolt Circle"

            iRowNo = 11

            If rsBalanceHoles.RecordCount > 3 Then
                For I = 1 To rsBalanceHoles.RecordCount - 3
                    Rows("13:13").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next I
            End If

            rsBalanceHoles.MoveFirst
            For I = 1 To rsBalanceHoles.RecordCount

                .Range("K" & iRowNo).Select
                .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
                .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
                .Range("L" & iRowNo).Select
                .ActiveCell = rsBalanceHoles.Fields("Number")
                .ActiveCell.NumberFormat = "0"
                .Range("M" & iRowNo).Select
                If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
                    .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
                    .ActiveCell.NumberFormat = "0.0000"
                Else
                    .ActiveCell = rsBalanceHoles.Fields("Diameter1")
                End If

                .Range("N" & iRowNo).Select
                If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
                    .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
                    .ActiveCell.NumberFormat = "0.0000"
                Else
                    .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
                End If

                rsBalanceHoles.MoveNext
                iRowNo = iRowNo + 1
            Next I
            .Range("K10:N" & iRowNo - 1).Select
            With .Selection.Interior
                .ColorIndex = 34
                .Pattern = xlSolid
            End With
        End If 'rsBalanceHoles.RecordCount <> 0
    End If ' boGotBalanceHoles

    'plot graphs

    Dim SeriesName As String
    Dim XVals As String
    Dim YVals As String
    Dim RowNo As Long
    Dim RowStr As String
    Dim LastPoint As Integer
    Dim LineType As String
    Dim AxisGroup As Integer
    Dim LabelPos As Integer
    Dim LineColor As Long

        .ActiveSheet.ChartObjects("HydRepChart").Activate
        Dim S As Series
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
        'see if we have #N/A
'        Range("AX56").Select
'        If Not IsError(ActiveCell.value) Then
            Dim aq As Double
            Range("AQ56", "AQ71").Select
            aq = .Max(Selection)
            Dim ax As Double
            Range("AX56", "AX71").Select
            ax = .Max(Selection)

            'then current (as and az)
            Dim at As Double
            Range("AS56", "AS71").Select
            at = .Max(Selection)
            Dim ba As Double
            Range("AZ56", "AZ71").Select
            ba = .Max(Selection)


            Dim CurrentScaleMax As Integer
            Dim TDHScaleMax As Integer

            Dim MaxTDH As Integer
            With Application.WorksheetFunction
                If aq > ax Then
                    MaxTDH = .Ceiling(aq, 25)
                Else
                    MaxTDH = .Ceiling(ax, 25)
                End If
            End With

            Dim MaxCurrent As Integer
            With Application.WorksheetFunction
                If at > ba Then
                    Select Case at
                        Case Is <= 5
                            CurrentScaleMax = 5

                        Case Is <= 10
                            CurrentScaleMax = 10

                        Case Else
                            CurrentScaleMax = 25
                    End Select

                    MaxCurrent = .Ceiling(at, CurrentScaleMax)
                Else
                   Select Case ba
                        Case Is <= 5
                            CurrentScaleMax = 5

                        Case Is <= 10
                            CurrentScaleMax = 10

                        Case Else
                            CurrentScaleMax = 25
                    End Select

                    MaxCurrent = .Ceiling(ba, CurrentScaleMax)
                End If
        End With
        'End If
        ActiveSheet.ChartObjects("HydRepChart").Activate
         Dim ShtName As String
         ShtName = "'" & ActiveSheet.Name & "'"

        RowStr = 56 + 15
         For I = 1 To 8

             Select Case I
                 Case 1
                     SeriesName = "=""TDH"""
                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
                     YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
                     LineType = msoLineSolid
                     AxisGroup = 1
                     LabelPos = xlLabelPositionRight
                     LineColor = vbBlue

                 Case 2
                     SeriesName = "=""Input Power"""
                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
                     YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
                     LineType = msoLineSolid
                     AxisGroup = 2
                     LabelPos = xlLabelPositionRight
                     LineColor = vbRed

                 Case 3
                     SeriesName = "=""Current"""
                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
                     YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
                     LineType = msoLineSolid
                     AxisGroup = 2
                     LabelPos = xlLabelPositionRight
                     LineColor = vbGreen

                 Case 4
'                     SeriesName = "=""Overall Eff"""
'                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
'                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
'                     LineType = msoLineSolid
'                     AxisGroup = 2
'                     LabelPos = xlLabelPositionRight
'                     LineColor = vbCyan

                 Case 5
                     SeriesName = "=""TDH (Adj)"""
                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
                     YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
                     LineType = msoLineDash
                     AxisGroup = 1
                     LabelPos = xlLabelPositionBelow
                     LineColor = vbBlue

                 Case 6
                     SeriesName = "=""Input Power (Adj)"""
                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
                     YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
                     LineType = msoLineDash
                     AxisGroup = 2
                     LabelPos = xlLabelPositionBelow
                     LineColor = vbRed

                 Case 7
                     SeriesName = "=""Current (Adj)"""
                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
                     YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
                     LineType = msoLineDash
                     AxisGroup = 2
                     LabelPos = xlLabelPositionBelow
                     LineColor = vbGreen

                 Case 8
'                     SeriesName = "=""Overall Eff (Adj)"""
'                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
'                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
'                     LineType = msoLineDash
'                     AxisGroup = 2
'                     LabelPos = xlLabelPositionBelow
'                     LineColor = vbCyan

            End Select
            LastPoint = 16
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(I).Name = SeriesName
            ActiveChart.SeriesCollection(I).XValues = XVals
            ActiveChart.SeriesCollection(I).Values = YVals
            ActiveChart.SeriesCollection(I).Select
            ActiveChart.SeriesCollection(I).Points(LastPoint).Select
            ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
            ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
            If I < 5 Then
                Selection.ShowSeriesName = True
                Selection.Position = LabelPos
            Else
                Selection.ShowSeriesName = False
            End If
            Selection.ShowValue = False
            ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
            ActiveChart.SeriesCollection(I).Select
            With Selection.Format.line
                .Visible = msoTrue
                .DashStyle = LineType
                .ForeColor.RGB = LineColor
            End With


            ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
            ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
            ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
        Next I

        'show design point
        SeriesName = "=""Design Point"""
        XVals = "=" & ShtName & "!$L$63"
        YVals = "=" & ShtName & "!$L$64"
        LineType = msoLineSolid
        AxisGroup = 1
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(I).Name = SeriesName
        ActiveChart.SeriesCollection(I).XValues = XVals
        ActiveChart.SeriesCollection(I).Values = YVals
        ActiveChart.SeriesCollection(I).Select

        Selection.MarkerStyle = 4
        Selection.MarkerSize = 7
        With Selection.Format.line
            .Visible = msoTrue
            .Weight = 2.25
            .ForeColor.RGB = vbBlack
        End With


        ActiveChart.Axes(xlValue).Select
        ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
        ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

        ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
        ActiveChart.Axes(xlValue).MinimumScale = 0
        ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
        Selection.TickLabels.NumberFormat = "0"

        ActiveChart.Axes(xlValue, xlSecondary).Select
        ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
        ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True

        ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
        ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
        ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
        Selection.TickLabels.NumberFormat = "0"

        ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
'        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
        ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
        'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"

        Range("A1").Select

        'delete all macros in the excel file

        ' Declare variables to access the macros in the workbook.
        Dim objProject As VBIDE.VBProject
        Dim objComponent As VBIDE.VBComponent
        Dim objCode As VBIDE.CodeModule

        ' Get the project details in the workbook.
        Set objProject = xlBook.VBProject

        ' Iterate through each component in the project.
        For Each objComponent In objProject.VBComponents

            ' Delete code modules
            Set objCode = objComponent.CodeModule
            objCode.DeleteLines 1, objCode.CountOfLines

            Set objCode = Nothing
            Set objComponent = Nothing
        Next

        Set objProject = Nothing


        xlApp.Visible = True                    'show the sheet

'        xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
'        xlApp.Run "AssignButton"
    End With

'    Exit Sub

ErrHandler:
    'User pressed the Cancel button

    On Error GoTo notopen
    If Not xlApp.ActiveWorkbook Is Nothing Then
        ActiveWorkbook.CheckCompatibility = False
        xlApp.ActiveWorkbook.Save               'save the workbook
        'xlApp.ActiveWorkbook.Close

    End If

notopen:

'    xlApp.Application.Quit

'    xlApp.Quit
'    Set xlApp = Nothing

'    If CommonDialog1.filename <> "" Then
'        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
'    End If

On Error GoTo 0

    Exit Sub
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

    'see what worksheet tabs alread exist in the excel worksheet

    Dim intSheets As Integer    'number of sheets in the workbook
    Dim I As Integer
    Dim S As String
    Dim ans
    Dim NameOK As Boolean

    intSheets = xlApp.Worksheets.Count      'how many sheets are there?

    'define a crlf string
    S = vbCrLf

    For I = 1 To intSheets
        S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
    Next I

    'tell the user the names so far and ask if he/she wants to add another
    ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

    'get the answer
    If ans = vbNo Then
        GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
        Exit Function
    End If

    'get worksheet name from user and check to see that it's not already used

    NameOK = False  'start assuming that the name is bad

    While Not NameOK    'as long as it's bad, stay in this loop
        WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

        If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
            GetWorksheetTabs = vbNo
            Exit Function
        End If

        For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
            If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
                MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
                NameOK = False
                Exit For
            End If
            NameOK = True       'if we make it thru say the name is ok
        Next I
    Wend

    xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
    xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
    GetWorksheetTabs = vbYes                                            'say that the results were ok
  
End Function
Function NewWorkBook() As String

    Dim WorkSheetName As String

    'we've just added a new workbook, delete sheet1, sheet2, etc
    xlApp.DisplayAlerts = False
    While xlApp.Worksheets.Count > 1
        xlApp.Worksheets(1).Delete          'delete the sheet
    Wend
    xlApp.DisplayAlerts = True

    WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
    xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

    NewWorkBook = WorkSheetName
  
End Function

Sub FindMagtrols()
    Dim I As Integer
    Dim j As Integer
    Dim rs As New ADODB.Recordset

    Do While cmbMagtrol.ListCount > 0
        cmbMagtrol.RemoveItem cmbMagtrol.ListCount - 1
    Loop

'==============
    Dim sGPIBAddress As String
    Dim sGPIBName As String
    rs.Open "GPIBAddresses", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

    rs.MoveFirst                                'goto the top
    For I = 0 To rs.RecordCount - 1             'go through the whole recordset
        sGPIBAddress = rs.Fields("IPAddress")        'get the description
        sGPIBName = rs.Fields("GPIBName")                      'get the index number - promary key
        j = PingSilent(sGPIBAddress)
        If j <> 0 Then
            'also get the type of magtrol (5300 or 6530) from CheckMagtrolModel
            sGPIBName = sGPIBName & CheckMagtrolModel(Val(Right(sGPIBName, 1)))
            If iberr = 0 Then
                cmbMagtrol.AddItem sGPIBName
                cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = Val(Mid(sGPIBName, 5, 1))
            End If
        End If
        rs.MoveNext                             'get the next record
    Next I
    rs.Close
    Set rs = Nothing

    cmbMagtrol.AddItem "Add Manually"
    cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
    cmbMagtrol.ListIndex = 0
  
End Sub
Private Function CheckMagtrolModel(GPIBNo As Integer) As String
    Dim I As Integer
    Dim strRead As String
    Dim sSendStr As String
    strRead = Space$(182)

    'if we're talking to a magtrol, close the connection
    If iUD <> 0 Then
        ibonl iUD, 0
'        UnregisterGPIBGlobals
        iUD = 0
    End If

    'open a new connection to the magtrol:
        'primary address = 14
        'secondary address = 0
        'timeout = 3 second
        'eoi mode = 1
        'stop reading when line feed character is received - 0x10
        'and return iUD

    ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

    If iberr Then
        I = 0
'        Debug.Print GPIBNo & " - i=" & iberr
        CheckMagtrolModel = ""
    Else    'if no error
        'ask who it is
        sSendStr = "*IDN?" & vbCrLf
        ibwrt iUD, sSendStr

        Sleep (1000)

        'see what the Magtrol says
        ibrd iUD, strRead
        '6530 will return a string like 6530 R 1.16"
        '5300 will return measurement data

        If Left(strRead, 4) = "6530" Then
            CheckMagtrolModel = " - 6530"
        ElseIf Left(strRead, 2) = "A=" Then
            CheckMagtrolModel = " - 5300"
        Else
            CheckMagtrolModel = " - Unknown"
        End If
'        Debug.Print GPIBNo & " - " & strRead
        If iberr Then
'            Debug.Print iberr
        End If
    End If
End Function

Private Sub CalibrateSoftware()
        frmCalibrate.Show
        'Calibrating = True
  
End Sub

Function ParseTEMCModelNo(cmbComboName As ComboBox, ltr As String)
    Dim I As Integer
    Dim iStart As Integer
    Dim iStop As Integer
    Dim strCompare As String

    For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
        iStart = InStr(1, cmbComboName.List(I), "[")
        iStop = InStr(1, cmbComboName.List(I), "]")
        strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
        If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
            cmbComboName.ListIndex = I                                              'if we do, set the combo box
            Exit For                                            'and we're done
        End If
'        cmbComboName.ListIndex = -1                             'else, remove any pointer
        cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
    Next I

    txtModelNo.Text = UCase(txtModelNo.Text)
    txtModelNo.SelStart = Len(txtModelNo.Text)
End Function
Public Function LoadCombo(cmbComboName As ComboBox, sTableName As String)
'load all of the pump parameter combo boxes from the tables on the database

    Dim I As Integer
    Dim sItem As String
    Dim iID As Integer
    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

    qy.ActiveConnection = cnPumpData
    If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
        qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
    Else
        qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
    End If
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic

    rs.Open qy

    On Error GoTo NoField

    rs.MoveFirst                                'goto the top

    For I = 0 To rs.RecordCount - 1             'go through the whole recordset
        sItem = rs.Fields("Description")        'get the description
        iID = rs.Fields(0)                      'get the index number - promary key
        cmbComboName.AddItem sItem, I                                   'add the description to the combo box
        cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
        rs.MoveNext                             'get the next record
    Next I
    rs.Close
    cmbComboName.ListIndex = -1
On Error GoTo 0
    Set rs = Nothing
    Set qy = Nothing
    Exit Function

NoField:
On Error GoTo 0
    Resume Next
  
End Function
Public Function LoadInstrumentationCombo(cmbComboName As ComboBox, sTableName As String)
'load all of the pump parameter combo boxes from the tables on the database

    Dim I As Integer
    Dim sItem As String
    Dim iID As Integer
    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset


    qy.ActiveConnection = cnPumpData
    If sTableName = "AnalyzerNo" Then
        qy.CommandText = "SELECT * FROM " & sTableName & " WHERE UseInDropdown = true ORDER BY val(Description)"
    Else
        qy.CommandText = "SELECT * FROM " & sTableName & " WHERE UseInDropdown = true ORDER BY Description"
    End If
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic

    rs.Open qy
    Dim j As Integer

    On Error GoTo NoField
    rs.MoveFirst                                'goto the top
    For I = 0 To rs.RecordCount - 1             'go through the whole recordset
        sItem = rs.Fields("Description")        'get the description
        iID = rs.Fields(0)                      'get the index number - promary key
        cmbComboName.AddItem sItem, I                                   'add the description to the combo box
        cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
        rs.MoveNext                             'get the next record
        j = I + 1
    Next I
    rs.Close

    cmbComboName.AddItem "---- Legacy Items Below ---", j
    j = j + 1

    qy.CommandText = "SELECT * FROM " & sTableName & " WHERE UseInDropdown = false ORDER BY val(Description)"
    rs.Open qy

    rs.MoveFirst                                'goto the top
    For I = 0 To rs.RecordCount - 1             'go through the whole recordset
        sItem = rs.Fields("Description")        'get the description
        iID = rs.Fields(0)                      'get the index number - promary key
        cmbComboName.AddItem sItem, I + j                                   'add the description to the combo box
        cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
        rs.MoveNext                             'get the next record
    Next I
    rs.Close

    cmbComboName.ListIndex = -1
On Error GoTo 0
    Set rs = Nothing
    Set qy = Nothing
    Exit Function

NoField:
'    bUseDropdown = False
On Error GoTo 0
    Resume Next
  
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

    Dim I As Integer
    Dim m As Single

    m = 0
    For I = 0 To UBound(GraphArray, 1)
        If GraphArray(I, 1) > m Then
            m = GraphArray(I, 1)
        End If
    Next I
    SetGraphMax = m
  
End Function
Public Function CalculateSpeed(CoefSq As Double, CoefLin As Double, CoefConstant As Double, InputHP As Double, SG As Double) As Integer
    Dim I As Integer
    Dim OldResult As Double
    Dim NewResult As Double

    CalculateSpeed = 0

    If SG > 5 Or SG < 0.01 Then
        MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
        Exit Function
    End If

    OldResult = 1000
    NewResult = 0

    I = 1

    Do While Abs(NewResult - OldResult) > 0.1
        ReDim Preserve results(I)
        Select Case I
            Case 1
                results(I - 1).HP = InputHP
            Case 2
                results(I - 1).HP = results(I - 2).HP * SG
            Case Else
                results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
        End Select
        OldResult = NewResult
        results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
        NewResult = results(I - 1).Speed
        If I > 15 Then
            If I = 0 Or I > 15 Then
                MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
                Exit Function
            End If
            Exit Function
        End If
        I = I + 1
    Loop
    CalculateSpeed = I - 1
End Function
Public Function CalcPoly(CoefSq As Double, CoefLin As Double, CoefConstant As Double, DataIn As Double) As Double
    CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
End Function
Sub FixCoef()
'format the coefficient box
'
    With xlApp
        .Range("C30:D30").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("E30:F30").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("G30:H30").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("C31:D31").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("C32:D32").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("C33:D33").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("E31:F31").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("E32:F32").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("E33:F33").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("G31:H31").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("G32:H32").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("G33:H33").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("B29:H29").Select
        With .Selection
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        .Selection.UnMerge
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Selection.Merge
        .Range("B29:H33").Select
        With .Selection.Interior
            .ColorIndex = 34
            .Pattern = xlSolid
        End With
        .Range("B29:H33").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("B29:H29").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Range("B30:H30").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Range("B30:B33").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Range("C30:D33").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Range("E30:F33").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Range("J29").Select
    End With
End Sub
Sub FixFormat()
'
'   format the final data
'
    With xlApp
        .Range("B49:E58").Select
        With .Selection.Interior
            .ColorIndex = 6
            .Pattern = xlSolid
        End With
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("B49:E49").Select
        .Selection.Merge
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideVertical).LineStyle = xlNone
        .Range("B50:B58").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("C50:C58").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("D50:D58").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        .Range("B50:E50").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        .Range("B49:E58").Select
        .Selection.Font.Bold = True

        .Range("B51:E58").Select
        .Selection.NumberFormat = "0.00"
        .Range("B49:E58").Select
        With .Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        .Range("B49:E49").Select
        .Selection.Merge
    End With
End Sub

Sub GetBalanceHoleData(SerialNumber As String, TestDate As String)
    If rsBalanceHoles.State = adStateOpen Then
        rsBalanceHoles.Close
    End If
    qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
               "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
               "FROM BalanceHoles " & _
               "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
               "ORDER BY [Date], Val([BoltCircle]);"

    rsBalanceHoles.Open qyBalanceHoles
    rsBalanceHoles.Filter = ""

    Set dgBalanceHoles.DataSource = rsBalanceHoles

    Dim c As Column
    For Each c In dgBalanceHoles.Columns
        Select Case c.DataField
        Case "BalanceHoleID"
            c.Visible = False
        Case "SerialNo"
            c.Visible = False
        Case "Date"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 2000
        Case "Number"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "Diameter"
            c.Visible = False
        Case "Diameter1"
            c.Caption = "Diameter"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "BoltCircle1"
            c.Caption = "Bolt Circle"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 800
        Case "BoltCircle"
            c.Visible = False
        Case Else ' hide all other columns.
            c.Visible = False
        End Select
    Next c
  
End Sub
Public Sub FixPointsToPlot()
    'count valid data test entry and set points to plot
    If DataGrid2.Row = -1 Then
        Exit Sub
    End If
    Dim PresentGridRow As Integer
    PresentGridRow = DataGrid2.Row
    Dim GridIndex As Integer
    UpDown2.value = 8
    If DataGrid2.Row <> -1 Then
        For GridIndex = 0 To 7
            DataGrid2.Row = GridIndex
            If DataGrid2.Columns("Flow") = 0 And DataGrid2.Columns("TDH") = 0 Then
                txtUpDn2.Text = GridIndex
                Exit Sub
            End If
        Next GridIndex
    End If
    DataGrid2.Row = PresentGridRow
End Sub


Private Sub ReportToExcel()

    frmReportOptions.Show 1

    Dim PosRPM As Integer
    Dim PosAxPos As Integer
    Dim PosCircFlow As Integer
    Dim PosVib As Integer
    Dim PosRem As Integer
    Dim PosTRG As Integer

    PosTRG = frmReportOptions.chkTRG.value * 12
    PosRPM = frmReportOptions.chkSelectRPM.value * 12 + frmReportOptions.chkTRG.value
    PosAxPos = frmReportOptions.chkSelectAxPos.value * 12 + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value
    PosCircFlow = frmReportOptions.chkSelectCircFlow.value * 12 + frmReportOptions.chkSelectAxPos.value + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value
    PosVib = frmReportOptions.chkVibration.value * 12 + frmReportOptions.chkSelectCircFlow.value + frmReportOptions.chkSelectAxPos.value + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value
    PosRem = 12 + frmReportOptions.chkVibration.value * 2 + frmReportOptions.chkSelectCircFlow.value + frmReportOptions.chkSelectAxPos.value + frmReportOptions.chkSelectRPM.value + frmReportOptions.chkTRG.value

    Dim SaveFileName As String
    Dim WorkSheetName As String

    Dim I As Integer
    Dim iRowNo As Integer
    Dim sImp As String
    Dim ans As Integer

    'excel
    Dim ReportWorkbookName As String
'    ReportWorkbookName = "C:\Users\MRosenbaum.CHEMPUMP\Desktop\HydraulicTestReportTemplate.xls"
    ReportWorkbookName = "\\tei-main-01\F\EN\GROUPS\SHARED\Software\Rundown Test Sheet Templates\HydraulicTestReportTemplate.xls"

    Dim SaveReportFileName As String
    Dim TemplateWorkSheetName As String
    TemplateWorkSheetName = "TestReport"

    Dim oXLApp As Excel.Application
    Dim oXLBook As Excel.Workbook
    Dim oXLNewBook As Excel.Workbook
    Dim oXLSheet As Excel.Worksheet
    Dim oXLSheetToCopy As Excel.Worksheet

    'open excel
    Set oXLApp = New Excel.Application

    'open the template as readonly
    Set oXLBook = oXLApp.Workbooks.Open(ReportWorkbookName, ReadOnly:=True)

    'open the report sheet
    Set oXLSheet = oXLBook.Worksheets(TemplateWorkSheetName)


    oXLApp.Visible = False

    'get the name for the saved report file
    CommonDialog1.CancelError = True        'in case the user
    On Error GoTo CancelErrHandler                '  chooses the cancel button

    CommonDialog1.DialogTitle = "Save Excel Hydraulic Report Files"
    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen                     'open the file selection dialog box

    If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
    Else                                                'the file name already exists
        ans = MsgBox(CommonDialog1.filename & " already exists.  Overwrite?", vbYesNo, "File Exists")
        If ans = vbYes Then
        Else
            MsgBox "Exiting routine.  Please reenter and select valid Report File Name", vbOKOnly, "Exiting . . ."
            Exit Sub
        End If

    End If

    SaveReportFileName = CommonDialog1.filename

    Set oXLNewBook = oXLApp.Workbooks.Add

    Set oXLSheetToCopy = oXLSheet
    oXLSheetToCopy.Copy oXLNewBook.Sheets(1)

    oXLApp.DisplayAlerts = False
    While oXLApp.Worksheets.Count > 1
        If oXLApp.Worksheets(oXLApp.Worksheets.Count).Name <> oXLSheet.Name Then
            oXLApp.Worksheets(oXLApp.Worksheets.Count).Delete           'delete the sheet
        End If
    Wend
    oXLApp.DisplayAlerts = True


    Set oXLSheet = Nothing
    oXLBook.Close savechanges:=False
    Set oXLBook = Nothing
    Set oXLSheet = oXLNewBook.Worksheets(TemplateWorkSheetName)

    Dim oXLVariant As Variant
    oXLVariant = oXLSheet.Range("A1:T44").value



    Dim XA(44, 20) As String
    Dim ir As Integer
    Dim ic As Integer

    For ir = 0 To 41
        For ic = 0 To 19
            XA(ir, ic) = oXLVariant(ir + 1, ic + 1)
        Next ic
    Next ir

    'write the data to the spreadsheet
    With oXLApp

    'write header data
        XA(1, 1) = "Run Date:"
        XA(1, 3) = CStr(cmbTestDate.List(cmbTestDate.ListIndex))
        XA(1, 15) = "Instrumentation / Setup"
        XA(2, 1) = "Serial Number:"
        XA(2, 3) = txtSN.Text
        XA(2, 7) = "Customer:"
        XA(2, 9) = Me.txtShpNo.Text
        XA(3, 1) = "Model:"
        XA(3, 3) = txtModelNo.Text
        XA(3, 13) = "Suction:"
        XA(3, 15) = txtSuctionID.Text
        XA(3, 17) = "Loop:"
        XA(3, 19) = cmbLoopNumber.List(cmbLoopNumber.ListIndex)
        XA(4, 13) = "Discharge:"
        XA(4, 15) = txtDischargeID.Text
        XA(4, 17) = "Orifice:"
        XA(4, 19) = cmbOrificeNumber.List(cmbOrificeNumber.ListIndex)
        XA(5, 1) = "Sales Order:"
        XA(5, 3) = txtSalesOrderNumber.Text
        XA(5, 5) = "Fluid:"
        XA(5, 6) = txtLiquid.Text
        XA(5, 9) = "Motor:"
        XA(5, 11) = cmbMotor.List(cmbMotor.ListIndex)
        XA(5, 13) = "Temperature:"
        XA(5, 15) = txtTemperatureID.Text
        XA(5, 17) = "Circ Flow:"
        XA(5, 19) = txtMagflowID.Text
        XA(6, 1) = "RMA:"
        XA(6, 3) = txtRMA.Text
        XA(6, 5) = "S. G.:"
        XA(6, 7) = txtSpGr.Text
        XA(6, 9) = "Voltage:"
        XA(6, 11) = cmbVoltage.List(cmbVoltage.ListIndex)
        XA(6, 13) = "Flow:"
        XA(6, 15) = txtFlowmeterID.Text
        XA(6, 17) = "PLC:"
        XA(6, 19) = cmbPLCNo.List(cmbPLCNo.ListIndex)
        XA(7, 5) = "Viscosity (cP):"
        XA(7, 7) = txtViscosity.Text
        XA(7, 9) = "Frequency (Hz):"
        XA(7, 11) = cmbFrequency.List(cmbFrequency.ListIndex)
        XA(7, 13) = "Power Analyzer:"
        XA(7, 15) = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)
        XA(7, 17) = "Tach:"
        XA(7, 19) = cmbTachID.List(cmbTachID.ListIndex)
        XA(8, 1) = "Test Spec:"
        XA(8, 3) = cmbTestSpec.List(cmbTestSpec.ListIndex)
        XA(8, 5) = "Temperature:"
        XA(8, 7) = Me.txtLiquidTemperature.Text
        XA(8, 9) = "Nominal RPM:"
        XA(8, 11) = cmbRPM.List(cmbRPM.ListIndex)
        XA(9, 13) = "Suction Pipe Dia (in):"
        XA(9, 16) = cmbSuctDia.List(cmbSuctDia.ListIndex)
        XA(10, 1) = "Design Point"
        XA(10, 5) = "Impeller Dia:"
        If chkTrimmed.value = 1 Then
            If Val(txtImpTrim.Text) <> 0 Then
                XA(10, 7) = txtImpTrim.Text
            Else
                XA(10, 7) = txtImpellerDia.Text
            End If
        Else
            XA(10, 7) = txtImpellerDia.Text
        End If
        XA(10, 9) = "Stator Fill:"
        XA(10, 11) = cmbStatorFill.List(cmbStatorFill.ListIndex)
        XA(10, 13) = "Suction Gage Height (in):"
        XA(10, 16) = txtSuctHeight.Text
        XA(11, 1) = "Flow Rate (GPM):"
        XA(11, 3) = txtDesignFlow.Text
        XA(11, 5) = "Design Pressure (psi):"
        XA(11, 7) = cmbDesignPressure.List(cmbDesignPressure.ListIndex)
        XA(11, 9) = "Full Load Current (A):"
        XA(11, 13) = "Discharge Pipe Dia (in):"
        XA(11, 16) = cmbDischDia.List(cmbDischDia.ListIndex)
        XA(12, 1) = "TDH (ft):"
        XA(12, 3) = txtDesignTDH.Text
        XA(12, 5) = "Circulation Path:"
        XA(12, 7) = cmbCirculationPath.List(cmbCirculationPath.ListIndex)
        XA(12, 9) = "Insulation Class:"
        XA(12, 13) = "Discharge Gage Height (in):"
        XA(12, 16) = txtDischHeight.Text
        XA(16, 1) = "Flow"
        XA(16, 2) = "TDH"
        XA(16, 3) = "KW"
        XA(16, 4) = "Ave"
        XA(16, 5) = "Ave"
        XA(16, 6) = "Power"
        XA(16, 7) = "Overall"
        XA(16, 8) = "Suction"
        XA(16, 9) = "Disch"
        XA(16, 10) = "Suction"
        XA(16, 11) = "Vel"

        XA(17, 1) = "(GPM)"
        XA(17, 2) = "(Ft)"
        XA(17, 4) = "Volts"
        XA(17, 5) = "Amps"
        XA(17, 6) = "Factor"
        XA(17, 7) = "Eff"
        XA(17, 8) = "Temp(F)"
        XA(17, 9) = "Pressure"
        XA(17, 10) = "Pressure"
        XA(17, 11) = "Head"

        XA(18, 11) = "(ft)"

        'variable data from user selection
        If PosTRG >= 12 Then
            XA(16, PosTRG) = "TRG"
            XA(17, PosTRG) = "Position"
        End If

        If PosRPM >= 12 Then
            XA(16, PosRPM) = "RPM"
        End If

        If PosVib >= 12 Then
            XA(16, PosVib) = "Vibration"
            XA(17, PosVib) = "Data X"
            XA(18, PosVib) = "(in/sec)"
            XA(16, PosVib + 1) = "Vibration"
            XA(17, PosVib + 1) = "Data Y"
            XA(18, PosVib + 1) = "(in/sec)"
        End If

        If PosAxPos >= 12 Then
            XA(16, PosAxPos) = "Axial"
            XA(17, PosAxPos) = "Position"
            XA(18, PosAxPos) = "(in)"
        End If

        If PosCircFlow >= 12 Then
            XA(16, PosCircFlow) = "Circ Flow"
            XA(17, PosCircFlow) = "(GPM)"
        End If

        XA(16, PosRem) = "Remarks"

        Dim j As Integer
        rsEff.MoveFirst
        For j = 1 To frmPLCData.UpDown2.value
            XA(18 + j, 1) = rsEff.Fields("Flow")
            XA(18 + j, 2) = Format(rsEff.Fields("TDH"), "##.00")
            XA(18 + j, 3) = Format(rsEff.Fields("KW"), "##.00")
            XA(18 + j, 4) = Format(rsEff.Fields("Volts"), "##.00")
            XA(18 + j, 5) = Format(rsEff.Fields("Amps"), "##.00")
            XA(18 + j, 6) = Format(rsEff.Fields("PowerFactor"), "##.00")
            XA(18 + j, 7) = Format(rsEff.Fields("OverallEfficiency"), "##.00")
            XA(18 + j, 8) = Format(rsEff.Fields("Temperature"), "##.00")
            XA(18 + j, 9) = Format(rsEff.Fields("DischPress"), "##.00")
            XA(18 + j, 10) = Format(rsEff.Fields("SuctPress"), "##.00")
            XA(18 + j, 11) = Format(rsEff.Fields("VelocityHead"), "##.00")

            If PosTRG >= 12 And Not IsNull(rsEff.Fields("TEMCTRG")) Then
                XA(18 + j, PosTRG) = rsEff.Fields("TEMCTRG")
            End If

            If PosRPM >= 12 And Not IsNull(rsEff.Fields("RPM")) Then
                XA(18 + j, PosRPM) = rsEff.Fields("RPM")
            End If

            If PosVib >= 12 And Not IsNull(rsEff.Fields("VibrationX")) And Not IsNull(rsEff.Fields("VibrationY")) Then
                XA(18 + j, PosVib) = rsEff.Fields("VibrationX")
                XA(18 + j, PosVib + 1) = rsEff.Fields("VibrationY")
            End If

            If PosAxPos >= 12 And Not IsNull(rsEff.Fields("Pos")) Then
                XA(18 + j, PosAxPos) = rsEff.Fields("Pos")
            End If

            If PosCircFlow >= 12 And Not IsNull(rsEff.Fields("CircFlow")) Then
                XA(18 + j, PosCircFlow) = rsEff.Fields("CircFlow")
            End If

            If Not IsNull(rsEff.Fields("Remarks")) Then
                XA(18 + j, PosRem) = rsEff.Fields("Remarks")
            End If

            rsEff.MoveNext
        Next j

        XA(28, 3) = "Thrust Balance Settings"
        If Me.chkFeathered.value = True Then
            XA(29, 10) = "Impeller has been feathered."
        Else
            XA(29, 10) = ""
        End If
        XA(30, 10) = "Discharge Orifice Size (in):"
        XA(30, 14) = Me.txtOrifice.Text
        XA(31, 10) = "Circulation Flow Orifice Size (in):"
        XA(31, 14) = Me.txtCircOrifice.Text
        XA(38, 1) = "Pump Remarks:"
        XA(38, 3) = Me.txtRemarks.Text
        XA(40, 1) = "Test Setup Remarks:"
        XA(40, 3) = Me.txtTestSetupRemarks.Text
        XA(42, 1) = "Other Modifications:"
        XA(42, 3) = Me.txtOtherMods.Text

        If boGotBalanceHoles Then
            If rsBalanceHoles.State = adStateClosed Then
                rsBalanceHoles.ActiveConnection = cnPumpData
                rsBalanceHoles.Open
            End If

            If rsBalanceHoles.RecordCount <> 0 Then
                rsBalanceHoles.MoveFirst
                For I = 1 To rsBalanceHoles.RecordCount
                    XA(29 + I, 1) = rsBalanceHoles.Fields("Date")
                    XA(29 + I, 4) = rsBalanceHoles.Fields("Number")
                    XA(29 + I, 5) = rsBalanceHoles.Fields("Diameter1")
                    XA(29 + I, 6) = rsBalanceHoles.Fields("BoltCircle1")
                    rsBalanceHoles.MoveNext
                Next I
            Else
            End If
        End If

        XA(29, 7) = "End Play(in):"
        XA(29, 9) = Me.txtEndPlay.Text
        XA(31, 7) = "G-Gap:"
        XA(31, 9) = Me.txtGGap.Text

        .Range("A1:T44").value = XA

    End With



    oXLNewBook.CheckCompatibility = False
    oXLNewBook.DoNotPromptForConvert = True
    oXLApp.DisplayAlerts = False
    oXLNewBook.SaveAs CommonDialog1.filename, FileFormat:=xlWorkbookNormal
    oXLNewBook.Close savechanges:=False
    oXLApp.DisplayAlerts = True

CancelErrHandler:

 '   oXLApp.Visible = True

    Set oXLSheet = Nothing
    Set oXLNewBook = Nothing
    Set oXLApp = Nothing

 '   oXLApp.Quit
On Error GoTo 0

    If CommonDialog1.filename <> "" Then
        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
    End If
  
End Sub



