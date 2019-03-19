Attribute VB_Name = "HPRoutines"
Option Explicit

'Global Definitions for the Database Connections
Global cnHP As New ADODB.Connection
Global rsHP As New ADODB.Recordset
Global QyHP As New ADODB.Command
Global cnHPOpen As Boolean      'status

Global rsHPDetail As New ADODB.Recordset
Global rsHPLineNo As New ADODB.Recordset

'Global Definitions for the parameters from the sifil
Global strShipTo As String
Global strBillTo As String
Global strModelNo() As String
Global strSerialNo() As String
Global strCapacity() As String
Global strTDH() As String
Global strImpellers() As String
Global strRPM() As String
Global strSpGr() As String
Global strFluid() As String
Global strPumpTemp() As String
Global strViscosity() As String
Global strVaporPress() As String
Global strSuctPress() As String
Global strDesignPress() As String
Global strSuctFlg() As String
Global strDischFlg() As String
Global strStatorFill() As String
Global strTestProcedure() As String
Global strVoltage() As String
Dim boolUseArchive As Boolean       'true - use the archive files, false - use the current files
Dim boUsingHP As Boolean
Dim boUsingArchive As Boolean
Global intMaxEntries As Integer
Global intLineNo As Integer
Global LogInInitials As String

'initials to approve and delete testing
Global Const strApproveInitials As String = "Admin"
Global boCanApprove As Boolean


' <VB WATCH>
Const VBWMODULE = "HPRoutines"
' </VB WATCH>

Function intFindTheValue(strLineText As String) As String
           'This function will find the value of the parameter on a line of text.
           'The line will consist of a parameter name, a colon, some whitespace and the value.
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "HPRoutines.intFindTheValue"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strLineText", strLineText) & ") "
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>

           'The input, strLineText, will be TDH(FT):       70
           'The output intFindTheValue will be 70.  This will be a string.  Any conversion,
           '  say to a number, must be made by the calling program.

           'If there is no value, an empty string "" will be returned

10         Dim strTemp As String
11         Dim lngWhere As Long

           'Find the :
12         lngWhere = InStr(strLineText, ":")

           'if it's not there, return an empty string
13         If IsNull(lngWhere) Then
14             intFindTheValue = vbNullString
' <VB WATCH>
15         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
16             Exit Function
17         End If

           'Else, get the string to the right of the :
18         strTemp = Right$(strLineText, Len(strLineText) - lngWhere)

           'Get rid of the white space around it, and return it
19         intFindTheValue = Trim$(strTemp)

' <VB WATCH>
20         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
21         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "intFindTheValue"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strLineText", strLineText
            vbwReportVariable "strTemp", strTemp
            vbwReportVariable "lngWhere", lngWhere
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function SearchSalesOrder(strSerialNumber As String)
' <VB WATCH>
22         On Error GoTo vbwErrHandler
23         Const VBWPROCNAME = "HPRoutines.SearchSalesOrder"
24         If vbwProtector.vbwTraceProc Then
25             Dim vbwProtectorParameterString As String
26             If vbwProtector.vbwTraceParameters Then
27                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strSerialNumber", strSerialNumber) & ") "
28             End If
29             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
30         End If
' </VB WATCH>

           'The function writes parameters from the sales order data to the file, PumpDataFromManMan
           'This is called with a sales order number

31         Dim I As Integer
32         Dim j As Integer
33         Dim c As String
34         Dim strTemp As String

35         Dim strSalesOrderNumber As String

36         strSalesOrderNumber = Left$(strSerialNumber, 7)

       '    'First, get the bill to and ship to info from soefil

37         boolUseArchive = False    'start by using the current files
38         GetBillToShipTo strSalesOrderNumber, strShipTo, strBillTo

       '
           'if there is no entry, skip the rest
39         If Len(strShipTo) < 1 Or IsNull(strShipTo) Then
40             I = SaveAnError(strSalesOrderNumber, "Bad Number", strShipTo)
' <VB WATCH>
41         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
42             Exit Function
43         End If


           'Get the Detail Data from sifil into a local recordset
44         GetDetail (strSalesOrderNumber)



           'get each piece of data

45         Dim intNoOfModelNo As Integer
46         intNoOfModelNo = intFindData(strModelNo, "MODEL NO:")
47         intMaxEntries = intMax(0, intNoOfModelNo)

48         Dim intNoOfSerialNo As Integer
49         intNoOfSerialNo = intFindData(strSerialNo, "SERIAL NO:")
50         intMaxEntries = intMax(intMaxEntries, intNoOfSerialNo)

51         Dim intNoOfCapacity As Integer
52         intNoOfCapacity = intFindData(strCapacity, "CAPACITY(GPM):")
53         intMaxEntries = intMax(intMaxEntries, intNoOfCapacity)

54         Dim intNoOfTDH As Integer
55         intNoOfTDH = intFindData(strTDH, "TDH(FT):")
56         intMaxEntries = intMax(intMaxEntries, intNoOfTDH)

57         Dim intNoOfImpeller As Integer
58         intNoOfImpeller = intFindData(strImpellers, "*IMPELLER DIA")    'just look for impeller
59         intMaxEntries = intMax(intMaxEntries, intNoOfImpeller)         ' could say dia or dia[ in]

60         Dim intNoOfRPM As Integer
61         intNoOfRPM = intFindData(strRPM, "SPEED:")
62         intMaxEntries = intMax(intMaxEntries, intNoOfRPM)

63         Dim intNoOfSpGr As Integer
64         intNoOfSpGr = intFindData(strSpGr, "SPECIFIC GRAVITY:")
65         intMaxEntries = intMax(intMaxEntries, intNoOfSpGr)

66         If intNoOfSpGr = 0 Then
67             intNoOfSpGr = intFindData(strSpGr, "SP GR:")
68             intMaxEntries = intMax(intMaxEntries, intNoOfSpGr)
69         End If

70         Dim intNoOfFluid As Integer
71         intNoOfFluid = intFindData(strFluid, "FLUID:")
72         intMaxEntries = intMax(intMaxEntries, intNoOfFluid)

73         Dim intNoOfPumpTemp As Integer
74         intNoOfPumpTemp = intFindData(strPumpTemp, "PUMPING TEMP:")
75         intMaxEntries = intMax(intMaxEntries, intNoOfPumpTemp)

76         Dim intNoOfViscosity As Integer
77         intNoOfViscosity = intFindData(strViscosity, "VISCOSITY:")
78         intMaxEntries = intMax(intMaxEntries, intNoOfViscosity)

79         Dim intNoOfVaporPress As Integer
80         intNoOfVaporPress = intFindData(strVaporPress, "VAPOR PRESS:")
81         intMaxEntries = intMax(intMaxEntries, intNoOfVaporPress)

82         Dim intNoOfSuctPress As Integer
83         intNoOfSuctPress = intFindData(strSuctPress, "SUCT PRESS:")
84         intMaxEntries = intMax(intMaxEntries, intNoOfSuctPress)

85         Dim intNoOfDesignPress As Integer
86         intNoOfDesignPress = intFindData(strDesignPress, "DESIGN PRESS")
87         intMaxEntries = intMax(intMaxEntries, intNoOfDesignPress)

88         Dim intNoOfSuctFlg As Integer
89         intNoOfSuctFlg = intFindData(strSuctFlg, "SUCT FLG:")
90         intMaxEntries = intMax(intMaxEntries, intNoOfSuctFlg)

91         Dim intNoOfDischFlg As Integer
92         intNoOfDischFlg = intFindData(strDischFlg, "DISCHARGE FLG:")
93         intMaxEntries = intMax(intMaxEntries, intNoOfDischFlg)

94         Dim intNoOfStatorFill As Integer
95         intNoOfStatorFill = intFindData(strStatorFill, "STATOR FILL:")
96         intMaxEntries = intMax(intMaxEntries, intNoOfStatorFill)

97         Dim intNoOfVoltages As Integer
98         intNoOfVoltages = intFindData(strVoltage, "*VOLT:")
99         intMaxEntries = intMax(intMaxEntries, intNoOfVoltages)

100        Dim intNoOfTestProcedure As Integer
101        intNoOfTestProcedure = intFindData(strTestProcedure, "*15852*")
102        intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
103        I = 15852

104        If intNoOfTestProcedure = 0 Then
105            intNoOfTestProcedure = intFindData(strTestProcedure, "*15605*")
106            intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
107            I = 15605
108        End If

109        If intNoOfTestProcedure = 0 Then
110            intNoOfTestProcedure = intFindData(strTestProcedure, "*19021*")
111            intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
112            I = 19021
113        End If

114        If intNoOfTestProcedure = 0 Then
115            intNoOfTestProcedure = intFindData(strTestProcedure, "*19530*")
116            intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
117            I = 19530
118        End If

119        If intNoOfTestProcedure = 0 Then
120            intNoOfTestProcedure = intFindData(strTestProcedure, "*17550*")
121            intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
122            I = 17550
123        End If

124        If I <> 0 Then
125            For j = 1 To 1000                               'get rid of any non-numbers
126                If Len(strTestProcedure(1, j)) <> 0 Then
127                    strTestProcedure(1, j) = Trim$(Str$((I)))
128                End If
129            Next j
130        End If

131        For I = 1 To 1000                               'get rid of any non-numbers
132            If Len(strRPM(1, I)) <> 0 Then
133                strTemp = vbNullString
134                For j = 1 To Len(strRPM(1, I))
135                    c = Mid$(strRPM(1, I), j, 1)
136                    If c >= "0" And c <= "9" Then
137                        strTemp = strTemp & c
138                    Else
139                        j = Len(strRPM(1, I))
140                    End If
141                Next j
142                strRPM(1, I) = strTemp
143            End If
144        Next I

           'Open the SOtoShipToBillTo table, and lookup the bill to and ship to names


145        If intMaxEntries = 0 Then       'no data found
146            I = SaveAnError(strSalesOrderNumber, "No Data Found", strShipTo)
' <VB WATCH>
147        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
148            Exit Function
149        End If

           'Get the line numbers and quantities from SODFIL and store them in rsHPLineNo
150        GetLineNoQuan (strSalesOrderNumber)


           'can we directly find the serial number that we're looking for?


151        With rsHPDetail
152            .Filter = "SICOM like '*" & strSerialNumber & "*'"
153            If .BOF = True And .EOF = True Then
154                intLineNo = 0
155            Else
156                intLineNo = Int(rsHPDetail.Fields(1))
157            End If
158        End With


' <VB WATCH>
159        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
160        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SearchSalesOrder"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strSerialNumber", strSerialNumber
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "c", c
            vbwReportVariable "strTemp", strTemp
            vbwReportVariable "strSalesOrderNumber", strSalesOrderNumber
            vbwReportVariable "intNoOfModelNo", intNoOfModelNo
            vbwReportVariable "intNoOfSerialNo", intNoOfSerialNo
            vbwReportVariable "intNoOfCapacity", intNoOfCapacity
            vbwReportVariable "intNoOfTDH", intNoOfTDH
            vbwReportVariable "intNoOfImpeller", intNoOfImpeller
            vbwReportVariable "intNoOfRPM", intNoOfRPM
            vbwReportVariable "intNoOfSpGr", intNoOfSpGr
            vbwReportVariable "intNoOfFluid", intNoOfFluid
            vbwReportVariable "intNoOfPumpTemp", intNoOfPumpTemp
            vbwReportVariable "intNoOfViscosity", intNoOfViscosity
            vbwReportVariable "intNoOfVaporPress", intNoOfVaporPress
            vbwReportVariable "intNoOfSuctPress", intNoOfSuctPress
            vbwReportVariable "intNoOfDesignPress", intNoOfDesignPress
            vbwReportVariable "intNoOfSuctFlg", intNoOfSuctFlg
            vbwReportVariable "intNoOfDischFlg", intNoOfDischFlg
            vbwReportVariable "intNoOfStatorFill", intNoOfStatorFill
            vbwReportVariable "intNoOfVoltages", intNoOfVoltages
            vbwReportVariable "intNoOfTestProcedure", intNoOfTestProcedure
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function GetBillToShipTo(strSalesOrderNumber As String, strShipTo As String, strBillTo As String)
           'enter with the Sales Order Number and return the Ship To and Bill To Names
           ' sets the boUsingHP if we found some data
           ' sets boUsingArchive if we found it in the Archive
' <VB WATCH>
161        On Error GoTo vbwErrHandler
162        Const VBWPROCNAME = "HPRoutines.GetBillToShipTo"
163        If vbwProtector.vbwTraceProc Then
164            Dim vbwProtectorParameterString As String
165            If vbwProtector.vbwTraceParameters Then
166                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strSalesOrderNumber", strSalesOrderNumber) & ", "
167                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("strShipTo", strShipTo) & ", "
168                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("strBillTo", strBillTo) & ") "
169            End If
170            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
171        End If
' </VB WATCH>

172        Dim sBillTo As String
173        Dim sShipTo As String

174        boUsingArchive = False  'assume that we're not usint the archive data

           'define the connection and set up the query
175        QyHP.ActiveConnection = cnHP
176        QyHP.CommandText = "SELECT SONUM, SHPNO, BILNO FROM SOEFIL WHERE SONUM = '" & _
                       strSalesOrderNumber & "'"

           'execute the query and get back a record set
177        Set rsHP = QyHP.Execute()

           'if the record set is empty, try the archive data
178        If rsHP.BOF = True And rsHP.EOF = True Then
179            boUsingArchive = True
180            QyHP.CommandText = "SELECT SONUM, SHPNO, BILNO FROM XSOEFIL WHERE SONUM = '" & _
                          strSalesOrderNumber & "'"
181            Set rsHP = QyHP.Execute()
182        End If

           'if this recordset is empty, we can't find the pump
           ' return null data for the bill to and ship to names
           'set boolean to say we can't find the pump
183        If rsHP.BOF = True And rsHP.EOF = True Then
184            boUsingHP = False
       '        MsgBox ("Record Not Found")
185            strShipTo = vbNullString
186            strBillTo = vbNullString
' <VB WATCH>
187        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
188            Exit Function
189        Else
               'else say we found it in the archive and we found the pump
190            boUsingHP = True
191        End If

           'get the numbers
192        sShipTo = rsHP.Fields(1)
193        sBillTo = rsHP.Fields(2)

           'get the data
           'set up the query for bill to

194        QyHP.CommandText = "SELECT BILNO, BILNAM FROM FINDB.BILMAS WHERE BILNO = '" & _
                      sBillTo & "'"

           'get the recordset -- should be one record
195        Set rsHP = QyHP.Execute()

           'extract the name
196        strBillTo = rsHP.Fields("BILNAM")

           'do the same for the ship to
197        QyHP.CommandText = "SELECT SHPNO, CUSNAM FROM FINDB.CUSFIL WHERE SHPNO = '" & _
                      sShipTo & "'"

198        Set rsHP = QyHP.Execute()

           'get the name
199        strShipTo = rsHP.Fields("CUSNAM")

' <VB WATCH>
200        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
201        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetBillToShipTo"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strSalesOrderNumber", strSalesOrderNumber
            vbwReportVariable "strShipTo", strShipTo
            vbwReportVariable "strBillTo", strBillTo
            vbwReportVariable "sBillTo", sBillTo
            vbwReportVariable "sShipTo", sShipTo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function GetLineNoQuan(strSalesOrderNumber As String)
' <VB WATCH>
202        On Error GoTo vbwErrHandler
203        Const VBWPROCNAME = "HPRoutines.GetLineNoQuan"
204        If vbwProtector.vbwTraceProc Then
205            Dim vbwProtectorParameterString As String
206            If vbwProtector.vbwTraceParameters Then
207                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strSalesOrderNumber", strSalesOrderNumber) & ") "
208            End If
209            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
210        End If
' </VB WATCH>

           'Select the correct file

211        Dim strFileName As String
212        Dim qry As String

213        If boUsingArchive = False Then
214            strFileName = "SODFIL"
215        Else
216            strFileName = "XSODFIL"
217        End If

           'Get the Line Numbers and Quantities

218        qry = "SELECT SONUM, LINE, SODQO " & _
                                 "FROM " & strFileName & " " & _
                                  "WHERE " & _
                                  " SONUM= '" & strSalesOrderNumber & "' " & _
                                  "ORDER BY LINE;"

           'put it in a new recordset, rsHPDetail

219        If rsHPLineNo.State = adStateOpen Then
220            rsHPLineNo.Close
221        End If

222        rsHPLineNo.CursorLocation = adUseClient
223        rsHPLineNo.Open qry, cnHP, adOpenStatic


' <VB WATCH>
224        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
225        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetLineNoQuan"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strSalesOrderNumber", strSalesOrderNumber
            vbwReportVariable "strFileName", strFileName
            vbwReportVariable "qry", qry
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function GetDetail(strSalesOrderNumber As String)
' <VB WATCH>
226        On Error GoTo vbwErrHandler
227        Const VBWPROCNAME = "HPRoutines.GetDetail"
228        If vbwProtector.vbwTraceProc Then
229            Dim vbwProtectorParameterString As String
230            If vbwProtector.vbwTraceParameters Then
231                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strSalesOrderNumber", strSalesOrderNumber) & ") "
232            End If
233            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
234        End If
' </VB WATCH>

           'Select the correct file

235        Dim strFileName As String

236        If boUsingArchive = False Then
237            strFileName = "SIFIL"
238        Else
239            strFileName = "XSIFIL"
240        End If

           'Get the details from SIFIL and write them into the SODetails table

           'get the data from the file

241        Dim qry As String
242        qry = "SELECT SONUM, SILIN, SICOM " & _
                                 "FROM " & strFileName & _
                                  " WHERE " & _
                                  " SONUM = '" & strSalesOrderNumber & _
                                  "' ORDER BY SONUM, SILIN;"

           'put it in a new recordset, rsHPDetail

243        If rsHPDetail.State = adStateOpen Then
244            rsHPDetail.Close
245        End If

246        rsHPDetail.CursorLocation = adUseClient
247        rsHPDetail.Open qry, cnHP, adOpenStatic

' <VB WATCH>
248        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
249        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetDetail"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strSalesOrderNumber", strSalesOrderNumber
            vbwReportVariable "strFileName", strFileName
            vbwReportVariable "qry", qry
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function intFindData(strArray() As String, strParameter As String) As Integer
           'This will load the array strArray with the
           '  data found in the rst recordset using the strCriteria.
           '  The array is 2 dimensional, the first entry is the line item and the second is the data.
' <VB WATCH>
250        On Error GoTo vbwErrHandler
251        Const VBWPROCNAME = "HPRoutines.intFindData"
252        If vbwProtector.vbwTraceProc Then
253            Dim vbwProtectorParameterString As String
254            If vbwProtector.vbwTraceParameters Then
255                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strArray", strArray) & ", "
256                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("strParameter", strParameter) & ") "
257            End If
258            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
259        End If
' </VB WATCH>

           'The function returns the number of entries found for the criteria

260        Dim strLocalData As String      'local variable for data
261        Dim strCriteria As String       'criteria to search for
262        Dim intCounter As Integer       'counts how many we found
263        Dim boolFoundAll As Boolean     'says we found all instances
264        Dim intMaxEntry As Integer      'maximum line item number entered into array
265        Dim intLineNo As Integer        'line number from sifil
266        Dim boolFirstLook As Boolean    'first time we're looking?

267        ReDim strArray(1, 1000)            'start with an array of 1x1000

268        intCounter = 0                  'none found yet
269        intMaxEntry = 0
270        boolFoundAll = False            'haven't found all instances

271        boolFirstLook = True
272        strCriteria = "SICOM like '" & strParameter & "*'"
273        With rsHPDetail
274            .Filter = 0
275            .Filter = strCriteria
276            If .BOF = True And .EOF = True Then
277                boolFoundAll = True
278            Else
279                .MoveFirst
280            End If
281        End With

282        While Not boolFoundAll
283            With rsHPDetail
284                If Not .EOF Then                                'if we found one
285                    intLineNo = Int(.Fields(1))                     'int portion of line number
286                    strLocalData = intFindTheValue(.Fields(2))         'strip the model no data
287                    strLocalData = Replace(strLocalData, "'", " ")
288                    strArray(1, intLineNo) = strLocalData          'store it in the array
289                    strArray(0, intLineNo) = Str$(Int(.Fields(1)))
290                    If intLineNo > intMaxEntry Then
291                        intMaxEntry = intLineNo
292                    End If
293                    intCounter = intCounter + 1                     'inc the counter
294                    .MoveNext
295                Else
296                    boolFoundAll = True
297                End If
298            End With
299        Wend
300        intFindData = intCounter       'return the number of ModelNos found
' <VB WATCH>
301        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
302        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "intFindData"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strArray", strArray
            vbwReportVariable "strParameter", strParameter
            vbwReportVariable "strLocalData", strLocalData
            vbwReportVariable "strCriteria", strCriteria
            vbwReportVariable "intCounter", intCounter
            vbwReportVariable "boolFoundAll", boolFoundAll
            vbwReportVariable "intMaxEntry", intMaxEntry
            vbwReportVariable "intLineNo", intLineNo
            vbwReportVariable "boolFirstLook", boolFirstLook
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function intMax(intOne As Integer, intTwo As Integer) As Integer
       'Returns the greater of the 2 integers sent to it
' <VB WATCH>
303        On Error GoTo vbwErrHandler
304        Const VBWPROCNAME = "HPRoutines.intMax"
305        If vbwProtector.vbwTraceProc Then
306            Dim vbwProtectorParameterString As String
307            If vbwProtector.vbwTraceParameters Then
308                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("intOne", intOne) & ", "
309                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("intTwo", intTwo) & ") "
310            End If
311            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
312        End If
' </VB WATCH>

313        If intOne > intTwo Then
314            intMax = intOne
315        Else
316            intMax = intTwo
317        End If
' <VB WATCH>
318        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
319        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "intMax"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "intOne", intOne
            vbwReportVariable "intTwo", intTwo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function SaveAnError(strSalesOrderNumber As String, ErrorType As String, strShipTo As String)
' <VB WATCH>
320        On Error GoTo vbwErrHandler
321        Const VBWPROCNAME = "HPRoutines.SaveAnError"
322        If vbwProtector.vbwTraceProc Then
323            Dim vbwProtectorParameterString As String
324            If vbwProtector.vbwTraceParameters Then
325                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strSalesOrderNumber", strSalesOrderNumber) & ", "
326                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ErrorType", ErrorType) & ", "
327                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("strShipTo", strShipTo) & ") "
328            End If
329            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
330        End If
' </VB WATCH>
331        Dim strSQLStatement As String

332        strSQLStatement = "INSERT INTO PumpDataFromManMan ([ErrorType], [SalesOrderNumber], [ShipTo], [BillTo], [ModelNo], [SerialNumber], " & _
                             "[Capacity], [TDH], [ImpellerDia], [RPM], [SpGr], [Fluid], [PumpTemp], [Viscosity], [VaporPress], " & _
                             "[SuctPress], [DesignPress], [SuctFlange], [DischFlange], [StatorFill], [TestProcedure]) " & _
                             "VALUES ('" & ErrorType & "', '" & strSalesOrderNumber & "', '" & _
                             strShipTo & "', '', '" & _
                             "', '', '" & _
                             "', '', '" & _
                             "', '', '" & _
                             "', '', '" & _
                             "', '', '" & _
                             "', '', '" & _
                             "', '', '" & _
                             "', '', '" & _
                             "');"


' <VB WATCH>
333        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
334        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SaveAnError"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strSalesOrderNumber", strSalesOrderNumber
            vbwReportVariable "ErrorType", ErrorType
            vbwReportVariable "strShipTo", strShipTo
            vbwReportVariable "strSQLStatement", strSQLStatement
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function


Function strFindTheNumber(strInput As String, intStart As Integer) As String
           'Enter with a string and a starting location
           'Return the substring containing the numeric value
' <VB WATCH>
335        On Error GoTo vbwErrHandler
336        Const VBWPROCNAME = "HPRoutines.strFindTheNumber"
337        If vbwProtector.vbwTraceProc Then
338            Dim vbwProtectorParameterString As String
339            If vbwProtector.vbwTraceParameters Then
340                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strInput", strInput) & ", "
341                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("intStart", intStart) & ") "
342            End If
343            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
344        End If
' </VB WATCH>

           'Ex.  enter with "PHASE: 3 CY: 60  VOLTS: 460" and Start=10
           '  the C of CY.  Return 60.

345        Dim L As Integer
346        Dim I As Integer
347        Dim j As Integer


348        If intStart = 0 Then
349            strFindTheNumber = vbNullString
' <VB WATCH>
350        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
351            Exit Function
352        End If

353        L = Len(strInput)
354        For I = intStart To L
               'move to the first numeric character
355            If IsNumeric(Mid$(strInput, I, 1)) Then
356                j = I
357                Exit For
358            End If
359        Next I

           'we're at a number or at the end of the string
           'if its the end of the string, return a null string
360        If I - 1 = L Then
361            strFindTheNumber = vbNullString
' <VB WATCH>
362        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
363            Exit Function
364        End If

           'else we're at a number
365        For I = j To L
366            If IsNumeric(Mid$(strInput, I, 1)) Then
367                strFindTheNumber = strFindTheNumber & Mid$(strInput, I, 1)
368            Else
369                Exit For
370            End If
371        Next I

' <VB WATCH>
372        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
373        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "strFindTheNumber"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
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
            vbwReportVariable "intStart", intStart
            vbwReportVariable "L", L
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function


' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
    vbwReportVariable "boolUseArchive", boolUseArchive
    vbwReportVariable "boUsingHP", boUsingHP
    vbwReportVariable "boUsingArchive", boUsingArchive
End Sub
' </VB WATCH>
