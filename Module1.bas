Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    With ActiveWorkbook.Connections("Report_KSPI_IFL").OLEDBConnection
        .BackgroundQuery = True
        .CommandText = Array("001_tbl_kspi_report")
        .CommandType = xlCmdTable
        .Connection = Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\FIN_045\Report_KSPI_IFL.accdb;Mode=Share Deny Write;E" _
        , _
        "xtended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Ty" _
        , _
        "pe=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Da" _
        , _
        "tabase Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compac" _
        , _
        "t=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass" _
        , _
        " UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        )
        .RefreshOnFileOpen = False
        .SavePassword = False
        .SourceConnectionFile = ""
        .SourceDataFile = "C:\FIN_045\Report_KSPI_IFL.accdb"
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
        .AlwaysUseConnectionFile = False
    End With
    With ActiveWorkbook.Connections("Report_KSPI_IFL")
        .Name = "Report_KSPI_IFL"
        .Description = ""
    End With
    ActiveWorkbook.Connections("Report_KSPI_IFL").Refresh
    With Selection.ListObject.QueryTable
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
    End With
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.ListObjects("Table_Report_KSPI_IFL.accdb").Unlist
End Sub

Sub ChangeFileFormat()

    Dim strCurrentFileExt   As String
    Dim strNewFileExt       As String
    Dim objFSO              As Object
    Dim objFolder           As Object
    Dim objFile             As Object
    Dim xlFile              As Workbook
    Dim strNewName          As String
    Dim strFolderPath       As String

    strCurrentFileExt = ".xls"
    strNewFileExt = ".xlsx"

    strFolderPath = "C:\Users\Scorpio\Desktop\New folder"
    If Right(strFolderPath, 1) <> "\" Then
        strFolderPath = strFolderPath & "\"
    End If

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.getfolder(strFolderPath)
    For Each objFile In objFolder.Files
        strNewName = objFile.Name
        If Right(strNewName, Len(strCurrentFileExt)) = strCurrentFileExt Then
            Set xlFile = Workbooks.Open(objFile.Path, , True)
            strNewName = Replace(strNewName, strCurrentFileExt, strNewFileExt)
            Application.DisplayAlerts = False
            Select Case strNewFileExt
            Case ".xlsx"
                xlFile.SaveAs strFolderPath & strNewName, XlFileFormat.xlOpenXMLWorkbook
            Case ".xlsm"
                xlFile.SaveAs strFolderPath & strNewName, XlFileFormat.xlOpenXMLWorkbookMacroEnabled
            End Select
            xlFile.Close
            Application.DisplayAlerts = True
        End If
    Next objFile

ClearMemory:
    strCurrentFileExt = vbNullString
    strNewFileExt = vbNullString
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
    Set xlFile = Nothing
    strNewName = vbNullString
    strFolderPath = vbNullString
End Sub
