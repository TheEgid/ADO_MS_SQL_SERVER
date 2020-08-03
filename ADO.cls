VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Option Explicit
'
'Sub ADOExample()
'
'Dim ADO As New ADO 'класс ado
'
'Dim Zs As String
'
'ADO.Header = False
'ADO.DataSource = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
'
'
'Zs = "SELECT ProductId, TRIM(ProductName), Price FROM DEMO.dbo.Goods;"
'ADO.Query Trim(Zs)
'
'Range("A2").CopyFromRecordset ADO.Recordset 'вывод
'
'ADO.Destroy '«акрываем соединение, чтобы не висело : )
'
'End Sub


'******************************************************************************************
'класс дл¤ подключени¤ к ms sql server c авторизацией windows
'******************************************************************************************


Option Explicit

Public Connection As Object
Public Recordset As Object
Public DataSource As String
Public Header As Boolean


Public Sub Create()
    Set Me.Connection = CreateObject("ADODB.Connection")
    Set Me.Recordset = CreateObject("ADODB.Recordset")
    
    Me.Recordset.CursorLocation = 3 ' adUseClient
End Sub


Public Sub Destroy()
    Call Me.Disconnect
    
    Set Me.Recordset = Nothing
    Set Me.Connection = Nothing
End Sub


Public Sub Connect(Optional ByVal ConnectionString As String)
    If Me.Connection Is Nothing Then
        Call Me.Create
    End If

    If ConnectionString = "" Then
        Me.Connection.Open GetExcelConnectionString()
    Else
        Me.Connection.Open ConnectionString
    End If
End Sub


Public Sub Disconnect()
    If Not Me.Recordset Is Nothing Then
        If Me.Recordset.State = 1 Then
            Me.Recordset.Close
        End If
    End If
    
    If Not Me.Connection Is Nothing Then
        If Me.Connection.State = 1 Then
            Me.Connection.Close
        End If
    End If
End Sub


Public Function Query(ParamArray QueryString() As Variant) As Date
    If Me.Connection Is Nothing Then
        Call Me.Create
    End If
    
    If Me.Connection.State = 0 Then
        Call Me.Connect
    End If
    
    If Me.Recordset.State = 1 Then
        Me.Recordset.Close
    End If
    
    Me.Recordset.Open Join(QueryString, " "), Me.Connection
    
    Query = Now
End Function


Public Function ToArray() As Variant
    Dim Field As Object
    Dim Arr() As Variant
    Dim Column As Integer
    Dim Row As Long
    
    If Me.Recordset Is Nothing Then
        Exit Function
    End If
    
    If Me.Recordset.BOF And Me.Recordset.EOF Then
        Exit Function
    End If
    
    ReDim Arr(1 To Me.Recordset.RecordCount, 1 To Me.Recordset.Fields.Count)
    
    Me.Recordset.MoveFirst
    
    Do Until Me.Recordset.EOF
        Row = Row + 1
        
        For Each Field In Me.Recordset.Fields
            Column = Column + 1
            Arr(Row, Column) = Field.Value
        Next
        
        Me.Recordset.MoveNext
        Column = 0
    Loop
    
    ToArray = Arr
End Function


Private Sub Class_Initialize()
    Call Me.Create
End Sub


Private Sub Class_Terminate()
    Call Me.Destroy
End Sub

Private Function GetExcelConnectionString() As String
    Dim strCon As String
    strCon = vbNullString
    strCon = strCon & "Provider=SQLOLEDB;"
    strCon = strCon & "Integrated Security=SSPI;"
    strCon = strCon & "Data Source='WIN-1BC8LPB9B96';"
    strCon = strCon & "Use Procedure for Prepare=1;"
    strCon = strCon & "Auto Translate=True;"
    strCon = strCon & "Packet Size=4096;"
    strCon = strCon & "Workstation ID=W-TPL-3275;"
    strCon = strCon & "Use Encryption for Data=False;"
    strCon = strCon & "Tag with column collation when possible=False;"
    GetExcelConnectionString = strCon
End Function

'Private Function GetExcelConnectionString() As String
'    Dim Params As String
'    Params = IIf(Val(Application.Version) < 12, _
'        "Provider='Microsoft.Jet.OLEDB.4.0';Data Source=':1';Extended Properties='Excel 4.0;HDR=:2;IMEX=1';", _
'        "Provider='Microsoft.ACE.OLEDB.12.0';Data Source=':1';Extended Properties='Excel 12.0;HDR=:2;IMEX=1';")
'    Params = Replace(Params, ":1", IIf(Me.DataSource = "", ThisWorkbook.FullName, Me.DataSource), , 1)
'    Params = Replace(Params, ":2", IIf(Me.Header, "Yes", "No"), , 1)
'    GetExcelConnectionString = Params
'End Function



