Attribute VB_Name = "EpicorRoutines"
'modified for E10
    Option Explicit
    Public Type SNRecord
        SONumber As String
        SOLine As String
        ModelNo As String
        MotorSize As String
        PartNum As String
        Customer As String
        ShipTo As String
        CustNum As String
        ShipToNum As String
        TDH As String
        Flow As String
        ImpellerDiameter As String
        SuctionPressure As String
        SpGr As String
        Fluid As String
        PumpTemperature As String
        Viscosity As String
        VaporPressure As String
        SuctFlangeSize As String
        DischFlangeSize As String
        RPM As String
        Voltage As String
        StatorFill As String
        CirculationPath As String
        TestProcedure As String
        DesignPressure As String
        Frequency As String
        XPartNum As String
        '
        Phases As String
        NPSHr As String
'        RatedInputPower As String
        RatedInputPower As String
        FLCurrent As String
        ThermalClass As String
        ExpClass As String
        LiquidTemp As String
        JobNumber As String
        CustomerPO As String
    End Type

' <VB WATCH>
Const VBWMODULE = "EpicorRoutines"
' </VB WATCH>

Public Function GetEpicorODBCData(SerialNumber As String, EpicorConnectionString As String) As SNRecord
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "EpicorRoutines.GetEpicorODBCData"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SerialNumber", SerialNumber) & ", "
7                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("EpicorConnectionString", EpicorConnectionString) & ") "
8              End If
9              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
10         End If
' </VB WATCH>
11         Dim conConn As New ADODB.Connection
12         Dim cmdCommand As New ADODB.Command
13         Dim rstRecordSet As New ADODB.Recordset
14         Dim SQLString As String

15         Dim MyRecord As SNRecord

           'construct connection string
16         conConn.Open EpicorConnectionString

       '   first see if there is an order number in the job file.  if there is, it is a make direct
       '       job and we bring back all of the information from Epicor as normal.
       '       if there is no order number, it is a make to stock job (supermarket), and we want
       '       to return the job number and part number only.  there is a table in the database
       '       that will get referenced for the supermarket data to put into temppumpdata

17         SQLString = "SELECT"
18         SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
19         SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
20         SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
21         SQLString = SQLString & " JobProd.OrderNum     AS SONumber, "
22         SQLString = SQLString & " JobProd.JobNum AS  JobProdJobNum "
23         SQLString = SQLString & " FROM Erp.SerialNo, Erp.JobProd "
24         SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
25         SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
26         SQLString = SQLString & ";"

27         With cmdCommand
28             .ActiveConnection = conConn
29             .CommandText = SQLString
30             .CommandType = adCmdText
31         End With

32         With rstRecordSet
33            .CursorType = adOpenStatic
34            .CursorLocation = adUseClient
35            .LockType = adLockBatchOptimistic
36            .Open cmdCommand
37          End With

           'if we have a record, save the data, else tell user and leave
38         If rstRecordSet.RecordCount > 0 Then    'there is no order no
39             If rstRecordSet.Fields("SONumber") = 0 Then
40                 rstRecordSet.MoveFirst
41                 MyRecord.PartNum = rstRecordSet.Fields("PartNum")
42                 MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
43                 MyRecord.SONumber = 0
                   'close the recordset and connection
44                 rstRecordSet.Close
45                 conConn.Close

46                 Set rstRecordSet = Nothing
47                 Set conConn = Nothing

48                 GetEpicorODBCData = MyRecord
' <VB WATCH>
49         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
50                 Exit Function
51             End If
52         End If

           'get job number, order number, order line and misc data from serial number  and order detail tables
53         SQLString = "SELECT"
54         SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
55         SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
56         SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
57         SQLString = SQLString & " JobProd.OrderNum     AS SONumber,"
58         SQLString = SQLString & " JobProd.OrderLine    AS SOLine,"
59         SQLString = SQLString & " OrderDtl.Character01  AS ModelNo, "
60         SQLString = SQLString & " OrderDtl.XPartNum  AS XPartNum, "
61         SQLString = SQLString & " OrderHed.CustNum  AS CustNum, "
62         SQLString = SQLString & " OrderHed.ShiptoNum AS ShipToNum, "
63         SQLString = SQLString & " OrderHed.PONum AS CustPONum, "
64         SQLString = SQLString & " Customer.Name AS CustomerName,  "
65         SQLString = SQLString & " ShipTo.Name AS ShipToName  "
66         SQLString = SQLString & " FROM dbo.OrderDtl AS OrderDtl, Erp.SerialNo, Erp.JobProd, Erp.OrderHed, Erp.Customer, Erp.ShipTo "
67         SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
68         SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
69         SQLString = SQLString & " AND OrderDtl.OrderNum = JobProd.OrderNum"
70         SQLString = SQLString & " AND OrderDtl.OrderLine = JobProd.OrderLine"
71         SQLString = SQLString & " AND OrderHed.OrderNum = JobProd.OrderNum"
72         SQLString = SQLString & " AND Customer.CustNum = OrderHed.CustNum"
73         SQLString = SQLString & " AND ShipTo.ShipToNum = OrderHed.ShipToNum"
74         SQLString = SQLString & ";"

75         With cmdCommand
76             .ActiveConnection = conConn
77             .CommandText = SQLString
78             .CommandType = adCmdText
79         End With

80         With rstRecordSet
81             If rstRecordSet.State = adStateOpen Then
82                 .Close
83             End If
84            .CursorType = adOpenStatic
85            .CursorLocation = adUseClient
86            .LockType = adLockBatchOptimistic
87            .Open cmdCommand
88          End With

           'if we have a record, save the data, else tell user and leave
89         If rstRecordSet.RecordCount > 0 Then
90             rstRecordSet.MoveFirst
91             MyRecord.SONumber = rstRecordSet.Fields("SONumber")
92             MyRecord.SOLine = rstRecordSet.Fields("SOLine")
93             MyRecord.ModelNo = rstRecordSet.Fields("ModelNo")
94             MyRecord.PartNum = rstRecordSet.Fields("PartNum")
95             MyRecord.CustNum = rstRecordSet.Fields("CustNum")
96             MyRecord.CustomerPO = rstRecordSet.Fields("CustPONum")
97             MyRecord.ShipToNum = rstRecordSet.Fields("ShipToNum")
98             MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
99             MyRecord.Customer = rstRecordSet.Fields("CustomerName")
100            MyRecord.XPartNum = rstRecordSet.Fields("XPartNum")
101            MyRecord.ShipTo = IIf(MyRecord.ShipToNum = "", rstRecordSet.Fields("CustomerName"), rstRecordSet.Fields("ShipToName"))
102        Else
103            MsgBox ("No Records found for Serial Number = " & SerialNumber)
' <VB WATCH>
104        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
105            Exit Function
106        End If

           'get ud02 data
107        SQLString = "SELECT"
108        SQLString = SQLString & " UD02.Number01         AS TDH,"
109        SQLString = SQLString & " UD02.Number02         AS Flow,"
110        SQLString = SQLString & " UD02.Number07         AS ImpellerDiameter,"
111        SQLString = SQLString & " UD02.Number03         AS SuctionPressure,"
112        SQLString = SQLString & " UD02.Number17         AS DesignPressure,"
113        SQLString = SQLString & " UD02.Number14         AS NPSHr"
114        SQLString = SQLString & " FROM Ice.UD02"
115        SQLString = SQLString & " WHERE UD02.Key1 = '" & MyRecord.SONumber & "' "
116        SQLString = SQLString & " AND UD02.Key2 = '" & MyRecord.SOLine & "' "

117        With rstRecordSet
118            .Close
119            cmdCommand.CommandText = SQLString
120            .Open cmdCommand
121        End With

122        If rstRecordSet.RecordCount > 0 Then
123            rstRecordSet.MoveFirst
124            MyRecord.TDH = rstRecordSet.Fields("TDH")
125            MyRecord.Flow = rstRecordSet.Fields("Flow")
126            MyRecord.ImpellerDiameter = rstRecordSet.Fields("ImpellerDiameter")
127            MyRecord.SuctionPressure = rstRecordSet.Fields("SuctionPressure")
128            MyRecord.DesignPressure = rstRecordSet.Fields("DesignPressure")
129            MyRecord.NPSHr = rstRecordSet.Fields("NPSHr")
130        End If

           'get ud03 data
131        SQLString = "SELECT"
132        SQLString = SQLString & " UD03.Number09         AS SpGr,"
133        SQLString = SQLString & " UD03.Character02      AS Fluid,"
134        SQLString = SQLString & " UD03.Number07         AS PumpTemperature,"
135        SQLString = SQLString & " UD03.Number11         AS Viscosity,"
136        SQLString = SQLString & " UD03.Number13         AS VaporPressure,"
137        SQLString = SQLString & " UD03.Number07           As LiquidTemp"
138        SQLString = SQLString & " FROM ice.UD03"
139        SQLString = SQLString & " WHERE UD03.Key1 = '" & MyRecord.SONumber & "' "
140        SQLString = SQLString & " AND UD03.Key2 = '" & MyRecord.SOLine & "' "

141        With rstRecordSet
142            .Close
143            cmdCommand.CommandText = SQLString
144            .Open cmdCommand
145        End With

146        If rstRecordSet.RecordCount > 0 Then
147            rstRecordSet.MoveFirst
148            MyRecord.SpGr = rstRecordSet.Fields("SpGr")
149            MyRecord.Fluid = rstRecordSet.Fields("Fluid")
150            MyRecord.PumpTemperature = rstRecordSet.Fields("PumpTemperature")
151            MyRecord.Viscosity = rstRecordSet.Fields("Viscosity")
152            MyRecord.VaporPressure = rstRecordSet.Fields("VaporPressure")
153            MyRecord.LiquidTemp = rstRecordSet.Fields("LiquidTemp")
154        End If

           'get ud04 data
155        SQLString = "SELECT"
156        SQLString = SQLString & " UD04.Character01      AS SuctFlangeSize,"
157        SQLString = SQLString & " UD04.Character04      AS DischFlangeSize"
158        SQLString = SQLString & " FROM ice.UD04"
159        SQLString = SQLString & " WHERE UD04.Key1 = '" & MyRecord.SONumber & "' "
160        SQLString = SQLString & " AND UD04.Key2 = '" & MyRecord.SOLine & "' "

161        With rstRecordSet
162            .Close
163            cmdCommand.CommandText = SQLString
164            .Open cmdCommand
165        End With

166        If rstRecordSet.RecordCount > 0 Then
167            rstRecordSet.MoveFirst
168            MyRecord.SuctFlangeSize = rstRecordSet.Fields("SuctFlangeSize")
169            MyRecord.DischFlangeSize = rstRecordSet.Fields("DischFlangeSize")
170        End If

           'get ud05 data
171        SQLString = "SELECT"
172        SQLString = SQLString & " UD05.Character05      AS RPM,"
173        SQLString = SQLString & " UD05.Character01      AS Voltage,"
174        SQLString = SQLString & " UD05.Character08      AS StatorFill,"
175        SQLString = SQLString & " UD05.Character02      AS Frequency,"
176        SQLString = SQLString & " UD05.Character03      AS Phases,"
177        SQLString = SQLString & " UD05.Character06        As ThermalClass,"
178        SQLString = SQLString & " UD05.Number01         AS RatedInputPower,"
179        SQLString = SQLString & " UD05.Number02         AS FLCurrent"
180        SQLString = SQLString & " FROM ice.UD05"
181        SQLString = SQLString & " WHERE UD05.Key1 = '" & MyRecord.SONumber & "' "
182        SQLString = SQLString & " AND UD05.Key2 = '" & MyRecord.SOLine & "' "

183        With rstRecordSet
184            .Close
185            cmdCommand.CommandText = SQLString
186            .Open cmdCommand
187        End With

188        If rstRecordSet.RecordCount > 0 Then
189            rstRecordSet.MoveFirst
190            MyRecord.RPM = rstRecordSet.Fields("RPM")
191            MyRecord.Voltage = rstRecordSet.Fields("Voltage")
192            MyRecord.StatorFill = rstRecordSet.Fields("StatorFill")
193            MyRecord.Frequency = rstRecordSet.Fields("Frequency")
194            MyRecord.Phases = rstRecordSet.Fields("Phases")
195            MyRecord.RatedInputPower = rstRecordSet.Fields("RatedInputPower")
196            MyRecord.FLCurrent = rstRecordSet.Fields("FLCurrent")
197            MyRecord.ThermalClass = rstRecordSet.Fields("ThermalClass")
198        End If

           'get ud07 data
199        SQLString = "SELECT"
200        SQLString = SQLString & " UD07.Character01      AS CirculationPath"
201        SQLString = SQLString & " FROM ice.UD07"
202        SQLString = SQLString & " WHERE UD07.Key1 = '" & MyRecord.SONumber & "' "
203        SQLString = SQLString & " AND UD07.Key2 = '" & MyRecord.SOLine & "' "

204        With rstRecordSet
205            .Close
206            cmdCommand.CommandText = SQLString
207            .Open cmdCommand
208        End With

209        If rstRecordSet.RecordCount > 0 Then
210            rstRecordSet.MoveFirst
211            MyRecord.CirculationPath = rstRecordSet.Fields("CirculationPath")
212        End If

           'get ud08 data
213        SQLString = "SELECT"
214        SQLString = SQLString & " UD08.ShortChar01      AS EXPRating"
215        SQLString = SQLString & " FROM ice.UD08"
216        SQLString = SQLString & " WHERE UD08.Key1 = '" & MyRecord.SONumber & "' "
217        SQLString = SQLString & " AND UD08.Key2 = '" & MyRecord.SOLine & "' "

218        With rstRecordSet
219            .Close
220            cmdCommand.CommandText = SQLString
221            .Open cmdCommand
222        End With


223        If rstRecordSet.RecordCount > 0 Then
224            rstRecordSet.MoveFirst
225            MyRecord.ExpClass = rstRecordSet.Fields("EXPRating")
226        End If

           'get ud09 data
227        SQLString = "SELECT"
228        SQLString = SQLString & " UD09.Character01      AS TestProcedure"
229        SQLString = SQLString & " FROM ice.UD09"
230        SQLString = SQLString & " WHERE UD09.Key1 = '" & MyRecord.SONumber & "' "
231        SQLString = SQLString & " AND UD09.Key2 = '" & MyRecord.SOLine & "' "

232        With rstRecordSet
233            .Close
234            cmdCommand.CommandText = SQLString
235            .Open cmdCommand
236        End With

237        If rstRecordSet.RecordCount > 0 Then
238            rstRecordSet.MoveFirst
239            MyRecord.TestProcedure = rstRecordSet.Fields("TestProcedure")
240        End If

           'get part data
241        SQLString = "SELECT"
242        SQLString = SQLString & " Part.Character06      AS MotorSize"
243        SQLString = SQLString & " FROM dbo.Part"
244        SQLString = SQLString & " WHERE Part.PartNum = '" & MyRecord.PartNum & "' "

245        With rstRecordSet
246            .Close
247            cmdCommand.CommandText = SQLString
248            .Open cmdCommand
249        End With

250        If rstRecordSet.RecordCount > 0 Then
251            rstRecordSet.MoveFirst
252            MyRecord.MotorSize = rstRecordSet.Fields("MotorSize")
253        End If

           'get Customer
254        SQLString = "SELECT"
255        SQLString = SQLString & " Customer.Name      AS Customer"
256        SQLString = SQLString & " FROM Customer"
257        SQLString = SQLString & " WHERE Customer.CustNum = '" & MyRecord.CustNum & "' "

258        With rstRecordSet
259            .Close
260            cmdCommand.CommandText = SQLString
261            .Open cmdCommand
262        End With

263        If rstRecordSet.RecordCount > 0 Then
264            rstRecordSet.MoveFirst
265            MyRecord.Customer = rstRecordSet.Fields("Customer")
266        End If

           'get ShipTo
267        If MyRecord.ShipToNum <> "" Then
268            SQLString = "SELECT"
269            SQLString = SQLString & " ShipTo.Name      AS ShipTo"
270            SQLString = SQLString & " FROM ShipTo"
271            SQLString = SQLString & " WHERE ShipTo.ShipToNum = '" & MyRecord.ShipToNum & "' "

272            With rstRecordSet
273                .Close
274                cmdCommand.CommandText = SQLString
275                .Open cmdCommand
276            End With

277            If rstRecordSet.RecordCount > 0 Then
278                rstRecordSet.MoveFirst
279                MyRecord.ShipTo = rstRecordSet.Fields("ShipTo")
280            End If
281        End If

           'close the recordset and connection
282        rstRecordSet.Close
283        conConn.Close

284        Set rstRecordSet = Nothing
285        Set conConn = Nothing

286        If MyRecord.ModelNo = "" Then
287            MyRecord.ModelNo = MyRecord.PartNum
288        End If

289        GetEpicorODBCData = MyRecord
' <VB WATCH>
290        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
291        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetEpicorODBCData"

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
            vbwReportVariable "EpicorConnectionString", EpicorConnectionString
            vbwReportVariable "SQLString", SQLString
            vbwReport_EpicorRoutines_SNRecord "MyRecord", MyRecord
            vbwReportVariable "conConn", conConn
            vbwReportVariable "cmdCommand", cmdCommand
            vbwReportVariable "rstRecordSet", rstRecordSet
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function



'   From E9
'    Option Explicit
'    Public Type SNRecord
'        SONumber As String
'        SOLine As String
'        ModelNo As String
'        MotorSize As String
'        PartNum As String
'        Customer As String
'        ShipTo As String
'        CustNum As String
'        ShipToNum As String
'        TDH As String
'        Flow As String
'        ImpellerDiameter As String
'        SuctionPressure As String
'        SpGr As String
'        Fluid As String
'        PumpTemperature As String
'        Viscosity As String
'        VaporPressure As String
'        SuctFlangeSize As String
'        DischFlangeSize As String
'        RPM As String
'        Voltage As String
'        StatorFill As String
'        CirculationPath As String
'        TestProcedure As String
'        DesignPressure As String
'        Frequency As String
'        '
'        Phases As String
'        NPSHr As String
'        RatedOutput As String
'        FLCurrent As String
'        ThermalClass As String
'        ExpClass As String
'        LiquidTemp As String
'        JobNumber As String
'    End Type
'
'Public Function GetEpicorODBCData(SerialNumber As String, EpicorConnectionString As String) As SNRecord
'    Dim conConn As New ADODB.Connection
'    Dim cmdCommand As New ADODB.Command
'    Dim rstRecordSet As New ADODB.Recordset
'    Dim SQLString As String
'
'    Dim MyRecord As SNRecord
'
'    'construct connection string
'    conConn.Open EpicorConnectionString

'
'
'    'get job number, order number, order line and misc data from serial number  and order detail tables
'    SQLString = "SELECT"
'    SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
'    SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
'    'SQLString = SQLString & " SerialNo.CustNum    AS CustNum,"
'    'SQLString = SQLString & " SerialNo.ShipToNum    AS ShipToNum,"
'    SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
'    SQLString = SQLString & " JobProd.OrderNum     AS SONumber,"
'    SQLString = SQLString & " JobProd.OrderLine    AS SOLine,"
'    SQLString = SQLString & " OrderDtl.Character01  AS ModelNo, "
'    SQLString = SQLString & " OrderHed.CustNum  AS CustNum, "
'    SQLString = SQLString & " OrderHed.ShiptoNum AS ShipToNum, "
'    SQLString = SQLString & " Customer.Name AS CustomerName,  "
'    SQLString = SQLString & " ShipTo.Name AS ShipToName  "
'    SQLString = SQLString & " FROM OrderDtl, SerialNo, JobProd, OrderHed, Customer, ShipTo "
'    SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
'    SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
'    SQLString = SQLString & " AND OrderDtl.OrderNum = JobProd.OrderNum"
'    SQLString = SQLString & " AND OrderHed.OrderNum = JobProd.OrderNum"
'    SQLString = SQLString & " AND Customer.CustNum = OrderHed.CustNum"
'    SQLString = SQLString & " AND ShipTo.ShipToNum = OrderHed.ShipToNum"
'    SQLString = SQLString & ";"
'
'    With cmdCommand
'        .ActiveConnection = conConn
'        .CommandText = SQLString
'        .CommandType = adCmdText
'    End With
'
'    With rstRecordSet
'       .CursorType = adOpenStatic
'       .CursorLocation = adUseClient
'       .LockType = adLockBatchOptimistic
'       .Open cmdCommand
'     End With
'
'    'if we have a record, save the data, else tell user and leave
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.SONumber = rstRecordSet.Fields("SONumber")
'        MyRecord.SOLine = rstRecordSet.Fields("SOLine")
'        MyRecord.ModelNo = rstRecordSet.Fields("ModelNo")
'        MyRecord.PartNum = rstRecordSet.Fields("PartNum")
'        MyRecord.CustNum = rstRecordSet.Fields("CustNum")
'        MyRecord.ShipToNum = rstRecordSet.Fields("ShipToNum")
'        MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
'        MyRecord.Customer = rstRecordSet.Fields("CustomerName")
'        MyRecord.ShipTo = IIf(MyRecord.ShipToNum = "", rstRecordSet.Fields("CustomerName"), rstRecordSet.Fields("ShipToName"))
'    Else
'        MsgBox ("No Records found for Serial Number = " & SerialNumber)
'        Exit Function
'    End If
'
'    'get ud02 data
'    SQLString = "SELECT"
'    SQLString = SQLString & " UD02.Number01         AS TDH,"
'    SQLString = SQLString & " UD02.Number02         AS Flow,"
'    SQLString = SQLString & " UD02.Number07         AS ImpellerDiameter,"
'    SQLString = SQLString & " UD02.Number03         AS SuctionPressure,"
'    SQLString = SQLString & " UD02.Number17         AS DesignPressure,"
'    SQLString = SQLString & " UD02.Number14         AS NPSHr"
'    SQLString = SQLString & " FROM UD02"
'    SQLString = SQLString & " WHERE UD02.Key1 = '" & MyRecord.SONumber & "' "
'    SQLString = SQLString & " AND UD02.Key2 = '" & MyRecord.SOLine & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.TDH = rstRecordSet.Fields("TDH")
'        MyRecord.Flow = rstRecordSet.Fields("Flow")
'        MyRecord.ImpellerDiameter = rstRecordSet.Fields("ImpellerDiameter")
'        MyRecord.SuctionPressure = rstRecordSet.Fields("SuctionPressure")
'        MyRecord.DesignPressure = rstRecordSet.Fields("DesignPressure")
'        MyRecord.NPSHr = rstRecordSet.Fields("NPSHr")
'    End If
'
'    'get ud03 data
'    SQLString = "SELECT"
'    SQLString = SQLString & " UD03.Number09         AS SpGr,"
'    SQLString = SQLString & " UD03.Character02      AS Fluid,"
'    SQLString = SQLString & " UD03.Number07         AS PumpTemperature,"
'    SQLString = SQLString & " UD03.Number11         AS Viscosity,"
'    SQLString = SQLString & " UD03.Number13         AS VaporPressure,"
'    SQLString = SQLString & " UD03.Number07           As LiquidTemp"
'    SQLString = SQLString & " FROM UD03"
'    SQLString = SQLString & " WHERE UD03.Key1 = '" & MyRecord.SONumber & "' "
'    SQLString = SQLString & " AND UD03.Key2 = '" & MyRecord.SOLine & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.SpGr = rstRecordSet.Fields("SpGr")
'        MyRecord.Fluid = rstRecordSet.Fields("Fluid")
'        MyRecord.PumpTemperature = rstRecordSet.Fields("PumpTemperature")
'        MyRecord.Viscosity = rstRecordSet.Fields("Viscosity")
'        MyRecord.VaporPressure = rstRecordSet.Fields("VaporPressure")
'        MyRecord.LiquidTemp = rstRecordSet.Fields("LiquidTemp")
'    End If
'
'    'get ud04 data
'    SQLString = "SELECT"
'    SQLString = SQLString & " UD04.Character01      AS SuctFlangeSize,"
'    SQLString = SQLString & " UD04.Character04      AS DischFlangeSize"
'    SQLString = SQLString & " FROM UD04"
'    SQLString = SQLString & " WHERE UD04.Key1 = '" & MyRecord.SONumber & "' "
'    SQLString = SQLString & " AND UD04.Key2 = '" & MyRecord.SOLine & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.SuctFlangeSize = rstRecordSet.Fields("SuctFlangeSize")
'        MyRecord.DischFlangeSize = rstRecordSet.Fields("DischFlangeSize")
'    End If
'
'    'get ud05 data
'    SQLString = "SELECT"
'    SQLString = SQLString & " UD05.Character05      AS RPM,"
'    SQLString = SQLString & " UD05.Character01      AS Voltage,"
'    SQLString = SQLString & " UD05.Character08      AS StatorFill,"
'    SQLString = SQLString & " UD05.Character02      AS Frequency,"
'    SQLString = SQLString & " UD05.Character03      AS Phases,"
'    SQLString = SQLString & " UD05.Character06        As ThermalClass,"
'    SQLString = SQLString & " UD05.Number01         AS RatedOutput,"
'    SQLString = SQLString & " UD05.Number02         AS FLCurrent"
'    SQLString = SQLString & " FROM UD05"
'    SQLString = SQLString & " WHERE UD05.Key1 = '" & MyRecord.SONumber & "' "
'    SQLString = SQLString & " AND UD05.Key2 = '" & MyRecord.SOLine & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.RPM = rstRecordSet.Fields("RPM")
'        MyRecord.Voltage = rstRecordSet.Fields("Voltage")
'        MyRecord.StatorFill = rstRecordSet.Fields("StatorFill")
'        MyRecord.Frequency = rstRecordSet.Fields("Frequency")
'        MyRecord.Phases = rstRecordSet.Fields("Phases")
'        MyRecord.RatedOutput = rstRecordSet.Fields("RatedOutput")
'        MyRecord.FLCurrent = rstRecordSet.Fields("FLCurrent")
'        MyRecord.ThermalClass = rstRecordSet.Fields("ThermalClass")
'    End If
'
'    'get ud07 data
'    SQLString = "SELECT"
'    SQLString = SQLString & " UD07.Character01      AS CirculationPath"
'    SQLString = SQLString & " FROM UD07"
'    SQLString = SQLString & " WHERE UD07.Key1 = '" & MyRecord.SONumber & "' "
'    SQLString = SQLString & " AND UD07.Key2 = '" & MyRecord.SOLine & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.CirculationPath = rstRecordSet.Fields("CirculationPath")
'    End If
'
'    'get ud08 data
'    SQLString = "SELECT"
'    SQLString = SQLString & " UD08.ShortChar01      AS EXPRating"
'    SQLString = SQLString & " FROM UD08"
'    SQLString = SQLString & " WHERE UD08.Key1 = '" & MyRecord.SONumber & "' "
'    SQLString = SQLString & " AND UD08.Key2 = '" & MyRecord.SOLine & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.ExpClass = rstRecordSet.Fields("EXPRating")
'    End If
'
'    'get ud09 data
'    SQLString = "SELECT"
'    SQLString = SQLString & " UD09.Character01      AS TestProcedure"
'    SQLString = SQLString & " FROM UD09"
'    SQLString = SQLString & " WHERE UD09.Key1 = '" & MyRecord.SONumber & "' "
'    SQLString = SQLString & " AND UD09.Key2 = '" & MyRecord.SOLine & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.TestProcedure = rstRecordSet.Fields("TestProcedure")
'    End If
'
'    'get part data
'    SQLString = "SELECT"
'    SQLString = SQLString & " Part.Character06      AS MotorSize"
'    SQLString = SQLString & " FROM Part"
'    SQLString = SQLString & " WHERE Part.PartNum = '" & MyRecord.PartNum & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.MotorSize = rstRecordSet.Fields("MotorSize")
'    End If
'
'    'get Customer
'    SQLString = "SELECT"
'    SQLString = SQLString & " Customer.Name      AS Customer"
'    SQLString = SQLString & " FROM Customer"
'    SQLString = SQLString & " WHERE Customer.CustNum = '" & MyRecord.CustNum & "' "
'
'    With rstRecordSet
'        .Close
'        cmdCommand.CommandText = SQLString
'        .Open cmdCommand
'    End With
'
'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.Customer = rstRecordSet.Fields("Customer")
'    End If
'
'    'get ShipTo
'    If MyRecord.ShipToNum <> "" Then
'        SQLString = "SELECT"
'        SQLString = SQLString & " ShipTo.Name      AS ShipTo"
'        SQLString = SQLString & " FROM ShipTo"
'        SQLString = SQLString & " WHERE ShipTo.ShipToNum = '" & MyRecord.ShipToNum & "' "
'
'        With rstRecordSet
'            .Close
'            cmdCommand.CommandText = SQLString
'            .Open cmdCommand
'        End With
'
'        If rstRecordSet.RecordCount > 0 Then
'            rstRecordSet.MoveFirst
'            MyRecord.ShipTo = rstRecordSet.Fields("ShipTo")
'        End If
'    End If
'
'    'close the recordset and connection
'    rstRecordSet.Close
'    conConn.Close
'
'    Set rstRecordSet = Nothing
'    Set conConn = Nothing
'
'    GetEpicorODBCData = MyRecord
'End Function
'
    

' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
