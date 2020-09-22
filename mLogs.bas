Attribute VB_Name = "mLogs"
Public Sub AddAlertToLog(vTextBox As TextBox, vText As String)
  
  Dim S As String
  
  S = "*** ALERT (" & Now & ")]" & vbCrLf
  S = S & "*** " & vText & vbCrLf & vbCrLf
  
  vTextBox = S & vTextBox
  
  DoEvents
  
End Sub

Public Sub CompileToLog(vTextBox As TextBox, vProject As String, vExe As String, vItemNumber As Integer, vTotalItems As Integer)
  
  Dim S As String
   
  S = "[PROJECT COMPILED: " & vItemNumber & " of " & vTotalItems & "  (" & Now & ")]" & vbCrLf
  S = S & vProject & vbCrLf
  S = S & vExe & vbCrLf & vbCrLf
  
  vTextBox = S & vTextBox
  
  DoEvents
  
End Sub
