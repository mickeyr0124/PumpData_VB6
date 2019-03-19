Attribute VB_Name = "vbwProtector"
 ' vbwProtector.bas file - Location: \VB Watch 2\Templates\VB6\Protector\ '
'                                                                        '
' This module contains all procedures common to the VB Watch tools.      '
' It will be added to every project instrumented with VB Watch.          '
'                                                                        '
' ************************* WARNING *******************************      '
' You should not modify it unles you know what you are doing.            '
' To modify it, remove the read-only attribute of vbwProtector.bas.      '
 ' WARNING: modifications of this file will apply to all error handling   '
'          plans !!!                                                     '

Option Explicit

' Options '
Public vbwCatchException As Boolean
Public vbwTraceProc As Boolean
Public vbwTraceParameters As Boolean
Public vbwTraceLine As Boolean
Public vbwCallStack As Boolean
Public vbwEmailRecipientAdress As String
Public vbwDumpStringMaxLength As Long
Public vbwSystemInfo As Boolean
Public vbwScreenshot As Boolean

' Variables for use with vbwFunctions.dll '
Public vbwAdvancedFunctions As Object          ' this will be used only if vbwFunctions.dll is installed on the enduser machine '
Public fIsVbwFunctionsInitialized As Boolean   ' true if vbwFunctions.dll is installed and instanciated                         '

' Call Stack '
Public vbwStackCalls() As String     ' array containing each call of the stack '
Public vbwStackCallsNumber As Long   ' number of calls = Ubound(vbwStackCalls) '

' Trace '
Public vbwTraceCallsNumber As Long   ' number of calls '

' Log File
Dim fIsLogInitialize As Boolean
Public vbwLogFile As String
Public vbwLogTraceToFile As Boolean
Dim fLogFileOpen As Boolean
Dim lLogFileNumber As Long
Dim lLogFileOffset As Long
' file I/O
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const OPEN_ALWAYS = 4
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        OffSet As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

' Var Dump
Const VBW_STRING = "**************************"
Global Const VBW_LOCAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* LOCAL LEVEL VARIABLES  *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_MODULE_STRING = vbCrLf & VBW_STRING & vbCrLf & "* MODULE LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_GLOBAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* GLOBAL LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_TYPE_STRING = " (User Defined Type Array)"
Global Const VBW_UNKNOWN_STRING = " = {Unknown Type}"
Global Const VBW_LOCAL_NOT_REPORTED = "Local Variables: not reported"
Global Const VBW_MODULE_NOT_REPORTED = "Module Variables: not reported"
Global Const VBW_GLOBAL_NOT_REPORTED = "Global Variables: not reported"
Global Const VBW_NO_LOCAL_VARIABLES = "No Local Variables"
Global vbwDumpFile As String
Global vbwDumpFileNum As Long

' Thread & processes
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

' Exception handling declarations
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long    ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type
Private Type EXCEPTION_DEBUG_INFO
        pExceptionRecord As EXCEPTION_RECORD
        dwFirstChance As Long
End Type
Private Type CONTEXT
    dblVar(66) As Double ' The real structure is more complex
    lngVar(6) As Long    ' but we don't need those details
End Type
Private Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type
Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_PRIV_INSTRUCTION = &HC0000096
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_GUARD_PAGE = &H80000001
Private Const EXCEPTION_INVALID_HANDLE = &HC0000008
Private Const CONTROL_C_EXIT = &HC000013A

' Variable to Save the Err object
Dim ErrObjectDescription As String
Dim ErrObjectHelpContext As Long
Dim ErrObjectHelpFile As String
Dim ErrObjectLastDllError As Long
Dim ErrObjectNumber As Long
Dim ErrObjectSource As String
Dim ErrLine As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public VBWPROTECTOR_EMPTY As Variant ' for use with vbwExecuteLine() in IIf structures

'Const VBW_EXE_EXTENSION = ".exe" ' this line will be rewritten by VB Watch with the right extension

' vbwNoTraceProc vbwNoTraceLine ' don't remove this !

#Const PROJECT = "PumpData.vbp"
Sub vbwInitializeProtector()

    Static vbwIsInitialized As Boolean

    If vbwIsInitialized Then
        Exit Sub
    End If

' Don't remove the following comments !                                                         '
' VB Watch will replace next line with the initialization code as set in the plan being applied '
vbwEmailRecipientAdress = "mrosenbaum@teikokupumps.com" ' this will be replaced with the value found in the 'General options' tab below
vbwCatchException = True ' set to False if you have already an exception catcher
vbwTraceLine = InStr(Command$, "/trace") > 0 Or GetSetting(App.title, "Init", "vbwTrace", "") = "1"
vbwTraceProc = InStr(Command$, "/trace") > 0 Or GetSetting(App.title, "Init", "vbwTrace", "") = "1"
vbwTraceParameters = vbwTraceProc
vbwCallStack = True
vbwSystemInfo = True
vbwScreenshot = True

    vbwLogTraceToFile = vbwTraceProc Or vbwTraceLine
    If vbwCallStack Then ' needed to track call stack
         vbwTraceProc = True
    End If

    vbwLogFile = Replace(App.Path & "\", ":\\", ":\") & "vbw" & App.EXEName & VBW_EXE_EXTENSION & ".log"
    vbwDumpFile = Replace(App.Path & "\", ":\\", ":\") & "vbw" & App.EXEName & VBW_EXE_EXTENSION & ".dmp"

    If vbwCatchException Then
        vbwHandleException
    End If

    vbwDumpStringMaxLength = 128 ' change this value to suit your need - make it 0 to remove the size check (to use with caution)

    vbwIsInitialized = True
  
End Sub

Sub vbwReportVariable(ByVal lName As String, ByVal lValue As Variant, Optional ByVal lTab As Long)
' vbwNoErrorHandler ' don't remove this !
    Dim I As Long, j As Long, k As Long, L As Long
    Dim tDim As Long

    On Error GoTo ErrDump

    If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
        ' array '
        tDim = GetArrayDimension(lValue)
        Select Case tDim
            Case 1
                vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & ") As " & TypeName(lValue))
                For I = LBound(lValue) To UBound(lValue)
                    vbwReportVariable lName & "(" & I & ")", lValue(I), lTab + 1
                Next I
            Case 2
                vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & ") As " & TypeName(lValue))
                For j = LBound(lValue, 2) To UBound(lValue, 2)
                    For I = LBound(lValue, 1) To UBound(lValue, 1)
                        vbwReportVariable lName & "(" & I & "," & j & ")", lValue(I, j), lTab + 1
                    Next I
                Next j
            Case 3
                vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & ") As " & TypeName(lValue))
                For k = LBound(lValue, 3) To UBound(lValue, 3)
                    For j = LBound(lValue, 2) To UBound(lValue, 2)
                        For I = LBound(lValue, 1) To UBound(lValue, 1)
                            vbwReportVariable lName & "(" & I & "," & j & "," & k & ")", lValue(I, j, k), lTab + 1
                        Next I
                    Next j
                Next k
            Case 4
                vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & "," & LBound(lValue, 4) & " To " & UBound(lValue, 4) & ") As " & TypeName(lValue))
                For L = LBound(lValue, 4) To UBound(lValue, 4)
                    For k = LBound(lValue, 3) To UBound(lValue, 3)
                        For j = LBound(lValue, 2) To UBound(lValue, 2)
                            For I = LBound(lValue, 1) To UBound(lValue, 1)
                                vbwReportVariable lName & "(" & I & "," & j & "," & k & "," & L & ")", lValue(I, j, k, L), lTab + 1
                            Next I
                        Next j
                    Next k
                Next L
            Case Else
                vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "() not processed: " & tDim & " dimensions")
        End Select
    Else
        ' non-array '
        If IsObject(lValue) Then
            vbwReportObject lName, lValue, lTab
        Else
            If VarType(lValue) = vbString Then
                lValue = FormatString(lValue)
            End If
            vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = " & lValue & " (" & TypeName(lValue) & ")")
        End If
    End If
    Exit Sub

ErrDump:
    Err.Clear
    vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = {Variable Dumping Error}")
End Sub

Public Sub vbwReportObject(lName As String, ByVal lObject As Object, Optional ByVal lTab As Long)
' vbwNoErrorHandler ' don't remove this !

    On Error GoTo ErrDump

    If TypeName(lObject) <> "ErrObject" Then
        If fIsVbwFunctionsInitialized Then
            ' this should be executed only if you are using a global error handler '
            ' that prepares properly the vbwAdvancedFunctions for object dumping   '
            vbwCloseDumpFile       ' close it because vbwAdvancedFunctions uses its own file writing routines '
            vbwAdvancedFunctions.ReportObject lName, lObject, lTab, TypeOf lObject Is Form, TypeOf lObject Is MDIForm
            vbwOpenDumpFile
        Else
            ' no vbwFunctions.dll available                         '
            ' only report the default value of objects and controls '
            If TypeOf lObject Is Form Or TypeOf lObject Is MDIForm Then
               On Error Resume Next
               vbwReportToFile vbwEncryptString("Form " & lName)
               Dim c As Control
               For Each c In lObject.Controls
                   vbwReportObject c.Name & vbwGetIndex(c), c, 1
               Next c
            Else
                If IsNumeric(lObject) Then
                    vbwReportVariable lName, CDbl(lObject), lTab
                Else
                    vbwReportVariable lName, CStr(lObject), lTab
                End If
            End If
        End If
    Else
        vbwReportToFile vbCrLf & vbwEncryptString("**** ErrObject Err ****")
        vbwReportVariable "Err.Number", ErrObjectNumber
        vbwReportVariable "Err.Source", ErrObjectSource
        vbwReportVariable "Err.Description", ErrObjectDescription
        vbwReportVariable "Err.HelpContext", ErrObjectHelpContext
        vbwReportVariable "Err.HelpFile", ErrObjectHelpFile
        If ErrObjectLastDllError = 0 Then
            vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
        Else
            ' get the API error description from the system
            Dim sBuffer As String * 512
            Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
            FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, Null, ErrObjectLastDllError, 0, sBuffer, 512, 0
            If InStr(sBuffer, Chr(0)) Then
                vbwReportVariable "Err.LastDllError", ErrObjectLastDllError & " (" & Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1) & ")"
            Else
                vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
            End If
        End If
    End If

    Exit Sub

ErrDump:
    Err.Clear
    vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & ".Value = {No Value Property}")
End Sub

Public Function vbwEncryptString(ByRef sString As String, Optional sKey) As String
' vbwNoErrorHandler ' don't remove this ! '

    On Error Resume Next
    If fIsVbwFunctionsInitialized = False Then
        ' no encryption without vbwFunctions.dll            '
        ' you may want to write your own encryption routine '
        vbwEncryptString = sString
    Else
        If IsMissing(sKey) Then
            ' If you filled the vardump encryption key in the VB Watch Options, your key will   '
            ' be already embeded in the vbwAdvancedFunctions.ObjectInfo property, so you do not '
            ' have to care about provideing a key                                               '
            vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString)
        Else
            ' Yet if you wish to overide  your default encryption key, simply pass it '
            ' in the sKey parameter                                                   '
            vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString, sKey)
        End If
    End If
End Function

Function vbwReportParameter(ByVal lName As String, ByRef lValue As Variant) As String
' vbwNoErrorHandler ' don't remove this !
    Dim I As Long, j As Long, k As Long
    Dim tDim As Long
    Dim retString As String

    On Error GoTo ErrDump

    If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
        ' array '
        tDim = GetArrayDimension(lValue)
        If tDim Then
            retString = lName & "("
            For I = 1 To tDim
                retString = retString & LBound(lValue, I) & " To " & UBound(lValue, I) & ","
            Next I
            Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
        Else
            retString = lName & "(Undimensioned Array)"
        End If
    Else
        ' non-array '
        If IsObject(lValue) Then
            ' object
            On Error Resume Next
            retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
            If Err.Number Then
                On Error GoTo ErrDump
                retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
            End If
        Else
            ' non-object
            If VarType(lValue) = vbString Then
               retString = lName & " = " & FormatString(lValue)
            Else
               retString = lName & " = " & lValue
            End If
        End If
    End If

    vbwReportParameter = retString
    Exit Function

ErrDump:
    Err.Clear
    vbwReportParameter = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Function vbwReportParameterByVal(ByVal lName As String, ByVal lValue As Variant) As String
' vbwNoErrorHandler ' don't remove this !
    Dim I As Long, j As Long, k As Long
    Dim tDim As Long
    Dim retString As String

    On Error GoTo ErrDump

    If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
        ' array '
        tDim = GetArrayDimension(lValue)
        If tDim Then
            retString = lName & "("
            For I = 1 To tDim
                retString = retString & LBound(lValue, I) & " To " & UBound(lValue, I) & ","
            Next I
            Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
        Else
            retString = lName & "(Undimensioned Array)"
        End If
    Else
        ' non-array '
        If IsObject(lValue) Then
            ' object
            On Error Resume Next
            retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
            If Err.Number Then
                On Error GoTo ErrDump
                retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
            End If
        Else
            ' non-object
            If VarType(lValue) = vbString Then
               retString = lName & " = " & FormatString(lValue)
            Else
               retString = lName & " = " & lValue
            End If
        End If
    End If

    vbwReportParameterByVal = retString
    Exit Function

ErrDump:
    Err.Clear
    vbwReportParameterByVal = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Sub vbwReportToFile(ByRef lString As String)
' vbwNoVariableDump ' don't remove this !
     If vbwDumpFileNum = 0 Then
          vbwOpenDumpFile
     End If
     On Error Resume Next
     Print #vbwDumpFileNum, lString
     If Err = 52 Then
        vbwCloseDumpFile
        vbwOpenDumpFile
        Print #vbwDumpFileNum, lString
     End If
End Sub

Sub vbwOpenDumpFile()
' vbwNoVariableDump ' don't remove this !
   vbwDumpFileNum = FreeFile
   Open vbwDumpFile For Append As #vbwDumpFileNum
End Sub

Sub vbwCloseDumpFile()
' vbwNoVariableDump ' don't remove this !
   Close #vbwDumpFileNum
End Sub

Private Function GetArrayDimension(ByRef arg As Variant) As Long
' vbwNoErrorHandler ' don't remove this !
    Dim I As Long, j As Long
    On Error Resume Next
    I = 0
    Do
        I = I + 1
        j = LBound(arg, I)
    Loop Until Err.Number
    GetArrayDimension = I - 1
End Function

Function vbwGetIndex(tObject As Variant) As String
' vbwNoErrorHandler ' don't remove this !
    On Error Resume Next
    vbwGetIndex = "(" & tObject.Index & ")"
End Function

Private Function FormatString(ByVal arg As String) As String
' vbwNoVariableDump ' don't remove this !

    If Right$(arg, 1) = "}" Then ' probably a VB Watch built-in message
         FormatString = arg
         Exit Function
    End If

    ' 1. truncate according to the vbwDumpStringMaxLength value
    If vbwDumpStringMaxLength Then
        If Len(arg) > vbwDumpStringMaxLength Then
            arg = Left$(arg, vbwDumpStringMaxLength + 1)   ' +1: avoids to cut inside a vbCrLf '
            If Right$(arg, 2) = vbCrLf Then
                ' don't cut inside a vbCrLf
            Else
                arg = Left$(arg, vbwDumpStringMaxLength)
            End If
            arg = arg & "{...}" ' truncated
        End If
    End If

    ' 2. make sure string isn't multiline
    arg = Replace(arg, vbCrLf, "<CrLf>", , , vbBinaryCompare)
    arg = Replace(arg, Chr(13), "<Cr>", , , vbBinaryCompare)
    arg = Replace(arg, Chr(10), "<Lf>", , , vbBinaryCompare)

    ' 3. add quotes
    FormatString = Chr(34) & arg & Chr(34)
End Function

Sub vbwProcIn(ByRef lProc As String, Optional ByRef lParameters As String)

    vbwTraceCallsNumber = vbwTraceCallsNumber + 1

    vbwStackCallsNumber = vbwStackCallsNumber + 1
    ReDim Preserve vbwStackCalls(1 To vbwStackCallsNumber)
    vbwStackCalls(vbwStackCallsNumber) = lProc

    Dim lString As String
    lString = String$(vbwTraceCallsNumber - 1, vbTab) & lProc

    If vbwLogTraceToFile Then
         vbwSendLog lString & lParameters
    End If
  
End Sub

Sub vbwProcOut(ByRef lProc As String)

    If vbwTraceCallsNumber > 0 Then ' should always be true
       vbwTraceCallsNumber = vbwTraceCallsNumber - 1
    End If

    If vbwStackCallsNumber > 0 Then ' should always be true
       vbwStackCallsNumber = vbwStackCallsNumber - 1
    End If
  
End Sub


Function vbwExecuteLine(ByRef fEncrypted As String, ByRef lline As String) As Boolean

    If vbwTraceLine Then

        If fEncrypted Then
            lline = "<CRY>" & lline & "</CRY>"
        End If

        If vbwLogTraceToFile Then
            If vbwTraceCallsNumber > 0 Then
                vbwSendLog String$(vbwTraceCallsNumber - 1, vbTab) & " -> " & lline
            Else
                vbwSendLog " -> " & lline
            End If
        End If

    End If

    ' This function always returns false
End Function

Function vbwGetStack() As String

    If vbwTraceProc = False Then
        vbwGetStack = "{Unavailable}"
        Exit Function
    End If

    Dim vbwStackString As String
    Dim I As Long

    For I = vbwStackCallsNumber To 1 Step -1
        vbwStackString = vbwStackString & String$(I - 1, vbTab) & vbwStackCalls(I) & vbCrLf
    Next I
    vbwGetStack = IIf(vbwStackString <> "", vbwStackString, "{Empty}")
End Function

Sub vbwSendLog(ByRef tMsg As String)
    If Err.Number Then
        ' Save Err object before being cleared by "On Error Resume Next"
        Dim ErrDescription As String, ErrHelpFile As String, ErrSource As String
        Dim ErrHelpContext As Long, ErrNumber As Long
        ErrDescription = Err.Description
        ErrHelpContext = Err.HelpContext
        ErrHelpFile = Err.HelpFile
        ErrNumber = Err.Number
        ErrSource = Err.Source
    End If

    On Error Resume Next

    If Not fLogFileOpen Then
        fLogFileOpen = True
        Dim suffix As Long
        Do
            Kill vbwLogFile
            lLogFileNumber = CreateFile(vbwLogFile, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_ALWAYS, FILE_FLAG_OVERLAPPED, 0)
            If lLogFileNumber < 0 Then
                ' under some circumstances (retained in memory applications or while in the IDE)
                ' the previous log file might not have been freed yet, so we must use another one
                suffix = suffix + 1
                vbwLogFile = Replace(App.Path & "\", ":\\", ":\") & "vbw" & App.EXEName & VBW_EXE_EXTENSION & suffix & ".log"
            End If
        Loop Until lLogFileNumber >= 0 Or suffix > 1000
    End If

    If Not fIsLogInitialize Then
        ' init file '
        fIsLogInitialize = True
        WriteToLogFile "Tracing " & App.title
        WriteToLogFile "Session started " & Now
        WriteToLogFile ""
    End If

    ' log to file
    WriteToLogFile tMsg

   If ErrNumber Then
        ' Restore Err object if cleared by "On Error Resume Next"
        Err.Description = ErrDescription
        Err.HelpContext = ErrHelpContext
        Err.HelpFile = ErrHelpFile
        Err.Number = ErrNumber
        Err.Source = ErrSource
    End If
  
End Sub

' Writes Str as a new line in the log file (adding a vbCrLf to the end)
Private Function WriteToLogFile(str As String) As Long
    Dim ol As OVERLAPPED
    Dim bBytes() As Byte, StrLength As Long
    StrLength = Len(str) + 2
    ReDim bBytes(0 To StrLength - 1)
    CopyMemory bBytes(0), ByVal str & vbCrLf, StrLength
    ol.OffSet = lLogFileOffset
    WriteToLogFile = WriteFileEx(lLogFileNumber, bBytes(0), StrLength, ol, ByVal 0&)
    lLogFileOffset = lLogFileOffset + StrLength
End Function


' Ends a component's thread. If this was the last active thread, ends the component's process.
Public Sub vbwExitThread()
    If vbwIsInIDE Then
        ' Executing ExitThread within the IDE will terminate VB without ceremony !
        Stop ' Press the End button now
    Else
        Dim lpExitCode As Long
        If GetExitCodeThread(GetCurrentThread(), lpExitCode) Then
            ExitThread lpExitCode
        End If
    End If
End Sub

' Ends a component's process. Equivalent to the End statement.
Public Sub vbwExitProcess()
    If vbwIsInIDE Then
        ' Executing ExitProcess within the IDE will terminate VB without ceremony !
        Stop ' Press the End button now
    Else
        Dim lpExitCode As Long
        If GetExitCodeProcess(GetCurrentProcess(), lpExitCode) Then
            ExitProcess lpExitCode
        End If
    End If
End Sub

' determines if the program is running in the IDE or an EXE File
Private Function vbwIsInIDE() As Boolean

    Dim strFileName As String
    Dim lngCount As Long

    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left(strFileName, lngCount)

    vbwIsInIDE = UCase$(Right$(strFileName, 8)) Like "\VB#.EXE"
  
End Function

' Exception handling stuff
Public Sub vbwHandleException()
    ' Exceptions will be caught and redirected to the failing procedure
    SetUnhandledExceptionFilter AddressOf vbwExceptionFilter
End Sub

' Exception handling stuff
Public Sub vbwUnHandleException()
    ' Exceptions are no longer caught and will cause Exceptions
    ' Whenever possible, call this procedure before returning to the VB's IDE
    SetUnhandledExceptionFilter 0
End Sub

' Exception handling stuff
Public Function vbwExceptionFilter(ByRef pExceptionInfo As EXCEPTION_POINTERS) As Long
'vbwNoErrorHandler ' DO NOT remove this !!!

    Dim ExceptionRecord As EXCEPTION_RECORD
    ExceptionRecord = pExceptionInfo.pExceptionRecord

    Do While ExceptionRecord.pExceptionRecord ' Empties the exceptions stack
        CopyMemory ExceptionRecord, ByVal ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
    Loop

    vbwExceptionFilter = EXCEPTION_CONTINUE_EXECUTION

'vbwExitProc ' because the next instruction causes to exit the function ' ' DO NOT remove this !!!

    ' Convert the exception to a normal VB error and go back to the failing procedure '
    Err.Raise 65535, , ExceptionDescription(ExceptionRecord.ExceptionCode)
  
End Function

' Exception handling stuff
Private Function ExceptionDescription(ByVal ExceptionCode As Long) As String
' vbwNoErrorHandler ' don't remove this !
    Select Case ExceptionCode
        Case EXCEPTION_ACCESS_VIOLATION
            ExceptionDescription = "Exception: Access Violation"
        Case EXCEPTION_DATATYPE_MISALIGNMENT
            ExceptionDescription = "Exception: Datatype Misalignment"
        Case EXCEPTION_BREAKPOINT
            ExceptionDescription = "Exception: Breakpoint"
        Case EXCEPTION_SINGLE_STEP
            ExceptionDescription = "Exception: Single Step"
        Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
            ExceptionDescription = "Exception: Array Bounds Exceeded"
        Case EXCEPTION_FLT_DENORMAL_OPERAND
            ExceptionDescription = "Exception: Float Denormal Operand"
        Case EXCEPTION_FLT_DIVIDE_BY_ZERO
            ExceptionDescription = "Exception: Float Divide By Zero"
        Case EXCEPTION_FLT_INEXACT_RESULT
            ExceptionDescription = "Exception: Float Inexact Result"
        Case EXCEPTION_FLT_INVALID_OPERATION
            ExceptionDescription = "Exception: Float Invalid Operation"
        Case EXCEPTION_FLT_OVERFLOW
            ExceptionDescription = "Exception: Float Overflow"
        Case EXCEPTION_FLT_STACK_CHECK
            ExceptionDescription = "Exception: Float Stack Check"
        Case EXCEPTION_FLT_UNDERFLOW
            ExceptionDescription = "Exception: Float Underflow"
        Case EXCEPTION_INT_DIVIDE_BY_ZERO
            ExceptionDescription = "Exception: Integer Divide By Zero"
        Case EXCEPTION_INT_OVERFLOW
            ExceptionDescription = "Exception: Integer Overflow"
        Case EXCEPTION_PRIV_INSTRUCTION
            ExceptionDescription = "Exception: Priv Instruction"
        Case EXCEPTION_IN_PAGE_ERROR
            ExceptionDescription = "Exception: In Page Error"
        Case EXCEPTION_ILLEGAL_INSTRUCTION
            ExceptionDescription = "Exception: Illegal Instruction"
        Case EXCEPTION_NONCONTINUABLE_EXCEPTION
            ExceptionDescription = "Exception: Non Continuable Exception"
        Case EXCEPTION_STACK_OVERFLOW
            ExceptionDescription = "Exception: Stack Overflow"
        Case EXCEPTION_INVALID_DISPOSITION
            ExceptionDescription = "Exception: Invalid Disposition"
        Case EXCEPTION_GUARD_PAGE
            ExceptionDescription = "Exception: Guard Page"
        Case EXCEPTION_INVALID_HANDLE
            ExceptionDescription = "Exception: Invalid Handle"
        Case CONTROL_C_EXIT
            ExceptionDescription = "Exception: Control C Exit"
        Case Else
            ExceptionDescription = "Unknown Exception"
    End Select
  
End Function

Public Sub vbwSaveErrObject()
' vbwNoErrorHandler ' don't remove this !
    ErrObjectDescription = Err.Description
    ErrObjectHelpContext = Err.HelpContext
    ErrObjectHelpFile = Err.HelpFile
    ErrObjectLastDllError = Err.LastDllError
    ErrObjectNumber = Err.Number
    ErrObjectSource = Err.Source
End Sub

Public Sub vbwRestoreErrObject()
' vbwNoErrorHandler ' don't remove this !
   Err.Description = ErrObjectDescription
   Err.HelpContext = ErrObjectHelpContext
   Err.HelpFile = ErrObjectHelpFile
   Err.Number = ErrObjectNumber
   Err.Source = ErrObjectSource
End Sub


