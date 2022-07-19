Attribute VB_Name = "Module1"
'-Begin-----------------------------------------------------------------

Option Explicit

Dim gColl() As String
Dim j As Integer

Sub GetAll(Obj As Object) '---------------------------------------------
'-
'- Recursively called sub routine to get the IDs of all UI elements
'-
'-----------------------------------------------------------------------

  Dim cntObj As Integer
  Dim i As Integer
  Dim Child As Object

  On Error Resume Next
  cntObj = Obj.Children.Count()
  If cntObj > 0 Then
    For i = 0 To cntObj - 1
      Set Child = Obj.Children.item(CLng(i))
      GetAll Child
      ReDim Preserve gColl(j)
      gColl(j) = CStr(Child.ID)
      j = j + 1
    Next
  End If
  On Error GoTo 0

End Sub

Sub Start() '-----------------------------------------------------------
'-
'- Sub routine to get all UI elements of the SAP GUI for Windows
'- with connection 0 and session 0
'-
'-----------------------------------------------------------------------

  Dim SapGuiAuto As Object
  Dim app As SAPFEWSELib.GuiApplication
  Dim connection As SAPFEWSELib.GuiConnection
  Dim session As SAPFEWSELib.GuiSession
  Dim i As Integer

  Set SapGuiAuto = GetObject("SAPGUI")
  If Not IsObject(SapGuiAuto) Then
    Exit Sub
  End If

  Set app = SapGuiAuto.GetScriptingEngine
  If Not IsObject(app) Then
    Exit Sub
  End If

  Set connection = app.Children(0)
  If Not IsObject(connection) Then
    Exit Sub
  End If

  If connection.DisabledByServer = True Then
    Exit Sub
  End If

  Set session = connection.Children(0)
  If Not IsObject(session) Then
    Exit Sub
  End If

  If session.Info.IsLowSpeedConnection = True Then
    Exit Sub
  End If

  GetAll session
  
  For i = LBound(gColl) To UBound(gColl)
    Cells(i + 1, 1) = gColl(i)
  Next

End Sub

'-End-------------------------------------------------------------------
