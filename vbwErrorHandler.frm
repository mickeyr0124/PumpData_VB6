VERSION 5.00
Begin VB.Form vbwFrmErrHandler 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "vbwErrorHandler.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   270
      ScaleHeight     =   3405
      ScaleWidth      =   6195
      TabIndex        =   6
      Top             =   2040
      Width           =   6195
      Begin VB.CommandButton cmdAction 
         Caption         =   "Always i&gnore this error"
         Height          =   435
         Index           =   6
         Left            =   4560
         TabIndex        =   9
         Top             =   2460
         Width           =   1200
      End
      Begin VB.CommandButton cmdLook 
         Caption         =   "Look at report"
         Height          =   345
         Left            =   4440
         TabIndex        =   8
         Top             =   195
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtDescribe 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   4065
      End
      Begin VB.CommandButton cmdSendMail 
         Caption         =   "Report &Error"
         Default         =   -1  'True
         Height          =   345
         Left            =   4440
         TabIndex        =   0
         Top             =   1470
         Width           =   1395
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Abort"
         Height          =   435
         Index           =   3
         Left            =   270
         TabIndex        =   2
         Top             =   2460
         Width           =   1200
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Retry"
         Height          =   435
         Index           =   4
         Left            =   1700
         TabIndex        =   3
         Top             =   2460
         Width           =   1200
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "&Ignore"
         Height          =   435
         Index           =   5
         Left            =   3130
         TabIndex        =   4
         Top             =   2460
         Width           =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   2
         X1              =   240
         X2              =   5804
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   240
         X2              =   5804
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label lblReport 
         BackStyle       =   0  'Transparent
         Caption         =   "lblReport"
         Height          =   885
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   5565
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   255
         X2              =   5804
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   255
         X2              =   5804
         Y1              =   2325
         Y2              =   2325
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "vbwErrorHandler.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblErrorString 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblErrorString"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   210
      Width           =   885
   End
End
Attribute VB_Name = "vbwFrmErrHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Const MF_BYPOSITION = &H400&

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Const vbMsgBoxSetTopMost = &H40000

Dim sReport As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

' vbwNoTraceProc vbwNoTraceLine ' don't remove this !
' vbwNoErrorHandler ' don't remove this !

Private Sub cmdAction_Click(Index As Integer)

    ' Index can be 3 (vbAbort), 4 (vbRetry), 5 (vbIgnore) or 6 for Ignore Procedure

    vbwErrHandler.vbwRetCode = Index              ' Returns an action
    Unload Me                                     ' and quits
  
End Sub

Private Sub cmdSendMail_Click()
    SetWindowPos Me.hwnd, -&H2, 0&, 0&, 0&, 0&, (&H1 Or &H2) ' lose the topmost status


    vbwRetCode = vbwDoDumpVariable
    vbwCircumstancesString = txtDescribe.Text

    MessageBox 0&, "This application will now retrieve some data needed to fix this error, open your email messenger and prepare a message to send. Then please press the ""Send Message"" button." _
                   & vbCrLf & "This may take from a few seconds up to a few minutes depending on the amount of data to retrieve." _
                   & vbCrLf & "During this time, you should normally notice hard drive and floppy drive activity." _
                   , App.title, vbInformation + vbMsgBoxSetTopMost
    Unload Me
    ' This is not finished !
    ' We just return to the failing procedure to dump its variables
    ' then end the email process in Public Function vbwErrorHandler()
End Sub

Private Sub Form_Load()

    Caption = App.title
    lblErrorString = vbwMessageString                ' vbwMessageString is already initialized  with the error message

    If vbwfHasReported Then
        lblReport = "Thanks for having reported this error !" & vbCrLf
        lblReport = lblReport & "Please select a continuation below..."
        txtDescribe = ""
        cmdSendMail.Enabled = False
    Else
        lblReport = "We are very sorry for this inconvenience and we would like to provide a fix for this error. "
        lblReport = lblReport & "You may help us a great deal by reporting this error via email. "
        lblReport = lblReport & "If you allow this, please fill the text box below and press the Report Error button. "
        lblReport = lblReport & "Thanks for your time."
        txtDescribe = "< Please describe here what you were doing exactly when this error occurred >" & vbCrLf
    End If

    ' Do some formating:
    ' Adjust form width to lblErrorString or picButtons
    Width = lblErrorString.Left + lblErrorString.Width + 260
    If Width < picButtons.Width Then
        Width = picButtons.Width
    End If
    ' picButtons is centered horizontally and moved below lblErrorString
    picButtons.Move (Width - picButtons.Width) / 2, lblErrorString.Top + lblErrorString.Height
    ' Adjust form height to the bottom of picButtons
    Height = picButtons.Top + picButtons.Height
    ' Center form
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    DisableCloseButton ' disable close button to force the user to make a choice
    Beep

    ' Set form topmost or the user might not see it if an other form is already topmost (such as a splash screen)
    If Not vbwfHasReported Then
         SetWindowPos Me.hwnd, -&H1, 0&, 0&, 0&, 0&, (&H1 Or &H2)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If vbwRetCode = 0 Then ' Default to Retry
         vbwRetCode = vbwRetry
    End If
End Sub

Private Sub txtDescribe_GotFocus()

    ' temporarily disable the cmdSendMail Default status to allow the "Enter" key
    cmdSendMail.Default = False

    ' automatically select the existing text
    txtDescribe.SelStart = 0
    txtDescribe.SelLength = 65000
  
End Sub

Private Sub txtDescribe_LostFocus()

    cmdSendMail.Default = True
  
End Sub

Private Sub DisableCloseButton()
     Dim hMenu As Long
     Dim nCount As Long
     hMenu = GetSystemMenu(Me.hwnd, 0)
     nCount = GetMenuItemCount(hMenu)

     Call RemoveMenu(hMenu, nCount - 1, MF_BYPOSITION)
     Call RemoveMenu(hMenu, nCount - 2, MF_BYPOSITION)

     DrawMenuBar Me.hwnd
End Sub

Public Property Let Report(ByVal sNewValue As String)
    sReport = sNewValue
    cmdLook.Visible = True
End Property

Private Sub cmdLook_Click()
    ShellExecute hwnd, "open", sReport, "", "", SW_SHOWNORMAL
End Sub



