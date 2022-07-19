Attribute VB_Name = "Module1"
'Macro created by Nicolas Sousa (U589310 | nmsousa@dow.com) 07-19-2022

Option Explicit

Sub Input_Data()
'input data to access database

    Dim conn As ADODB.Connection
    Dim databaseDir As String
    Dim wbName As String
    Dim shName As String
    Dim sqlCommand As String

    Sheets("Input").Select

    Set conn = New ADODB.Connection

    databaseDir = ThisWorkbook.Path & "\DATABASE.accdb"
    
    wbName = Application.ActiveWorkbook.FullName
    
    shName = "[" & Sheet1.Name & "$]"
    
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & databaseDir & ";"
    
    conn.Execute ("DELETE * FROM tbl_Fruit_Price")
    
    sqlCommand = "INSERT INTO tbl_Fruit_Price "
    sqlCommand = sqlCommand & "SELECT * FROM [Excel 8.0;HDR=YES;DATABASE=" & wbName & "]." & shName

    conn.Execute (sqlCommand)
    
    conn.Close
    
    Set conn = Nothing

End Sub

Sub Output_Data()
'get data from access database

Dim databaseDir As String
Dim sqlCommand As String
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim col As Integer

    Sheets("Output").Select
    
    Selection.Cells.Clear

    databaseDir = ThisWorkbook.Path & "\DATABASE.accdb"
    
    Set conn = New ADODB.Connection
    
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & databaseDir & ";"

    Set rst = New ADODB.Recordset
    
    With rst
        
        sqlCommand = "SELECT * FROM tbl_Fruit_Price"
        .Open Source:=sqlCommand, ActiveConnection:=conn

        For col = 0 To rst.Fields.Count - 1
            Range("A1").Offset(0, col).Value = rst.Fields(col).Name
        Next

        Range("A1").Offset(1, 0).CopyFromRecordset rst
    
    End With
    
    Set rst = Nothing
    
    conn.Close
    
    Set conn = Nothing

End Sub
