Attribute VB_Name = "NIGLOBAL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 32-bit Visual Basic Language Interface
' Version 1.7
' Copyright 1998 National Instruments Corporation.
' All Rights Reserved.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This module contains the variable  declarations,
'   constant definitions, and type information that
'   is recognized by the entire application.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Global ibsta As Integer
Global iberr As Integer
Global ibcnt As Integer

Global ibcntl As Long

' Needed to register for GPIB global Thread.
Global Longibsta As Long
Global Longiberr As Long
Global Longibcnt As Long
Global GPIBglobalsRegistered As Integer

' Error messages returned in global variable iberr

Global Const EDVR As Integer = 0      ' System error


' <VB WATCH>
Const VBWMODULE = "NIGLOBAL"
' </VB WATCH>

' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
