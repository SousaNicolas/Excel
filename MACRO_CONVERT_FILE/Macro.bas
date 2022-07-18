Attribute VB_Name = "Macro"
'Macro adapted from Stack Overflow
'Link: https://stackoverflow.com/questions/29167539/batch-convert-xls-to-xlsx-with-vba-without-opening-the-workbooks
'07-18-2022

Option Explicit

Sub ChangeFileFormat()
'Macro to convert file from xls to xlsx

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
