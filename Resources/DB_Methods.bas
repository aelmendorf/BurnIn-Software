Attribute VB_Name = "Module1"
Function Exist(wafer As String) As Boolean
    Dim Cmd As New ADODB.Command
    Dim Param As ADODB.Parameter
    Dim connect As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
     conStr = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
        "SERVER=172.20.4.20;" & _
        "DATABASE=epi;" & _
        "USER=aelmendorf;" & _
        "PASSWORD=Drizzle123!;" & _
        "Option=3"
        
    With connect
    .ConnectionString = conStr
    .CommandTimeout = 0
    .Open
    End With

    Set Cmd.ActiveConnection = connect
    
    Cmd.CommandText = "checkNew"
    Cmd.CommandType = CommandTypeEnum.adCmdStoredProc
    
    Set Param = Cmd.CreateParameter("wafer", adVarChar, adParamInput, 255, wafer)
    Cmd.Parameters.Append Param
    
    Set Param = Cmd.CreateParameter("test", adInteger, adParamInput, 255, 1)
    Cmd.Parameters.Append Param
    
    Set Param = Cmd.CreateParameter("?isentry", adInteger, adParamOutput, 255, isentry)
    Cmd.Parameters.Append Param
    
    Cmd.Execute Options:=adExecuteNoRecords
    
    If Cmd.Parameters("?isentry").Value = 1 Then
        Exist = True
    Else
        Exist = False
    End If
    connect.Close
    'Debug.Print Cmd.Parameters("?isentry").Value
    
End Function

Sub LastRow()
    Dim Last_Row As Integer
    Last_Row = Range("e2").Value
    Cells(Last_Row, 2).Select
    
End Sub

Sub InputSelected()
Dim dir As String
dir = "\\172.20.4.11\Data\Characterization Raw Data\Quick EL Test\B01-0956-08"

End Sub


Sub LastRowSummary()
    Dim Last_Row As Integer
    Last_Row = Range("a1").Value
    Cells(Last_Row, 1).Select
    
End Sub

Sub ImportData()
    Dim connect As New ADODB.Connection
    Dim Cmd As String
    Dim rg As Range
    Dim fld As ADODB.Field
    Dim rs As New ADODB.Recordset
    Dim conStr As String
    Dim cel As Range
    Dim selectedRange As Range
    conStr = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
        "SERVER=172.20.4.20;" & _
        "DATABASE=epi;" & _
        "USER=aelmendorf;" & _
        "PASSWORD=Drizzle123!;" & _
        "Option=3"
        
    Set selectedRange = Application.Selection
    With connect
    .ConnectionString = conStr
    .CommandTimeout = 0
    .Open
    End With

    For Each cel In selectedRange.Cells
        If cel.Column = 2 Then
            If Exist(cel.Value) Then
                Cmd = "Call getWaferData('" & cel.Value & ",1)"
                rs.Open "Call getWaferData('" & cel.Value & "',1)", connect, adOpenKeyset, adLockReadOnly, ADODB.adCmdText
                Range("M" & cel.Row).CopyFromRecordset rs
                rs.Close
                Set rg = Range("M" & cel.Row & ":AJ" & cel.Row)
                For Each dcel In rg.Cells
                    If dcel.Value = 0 Then
                        dcel.Value = ""
                    End If
                Next dcel
            End If
        End If
    Next cel
    MsgBox "Import Done"
    connect.Close

End Sub


Sub test()
Debug.Print Exist("11-2655")
End Sub
