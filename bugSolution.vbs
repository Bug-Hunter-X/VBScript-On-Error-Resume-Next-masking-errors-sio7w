Function GetValue(key)
  Dim result
  On Error GoTo ErrorHandler
  SetValue key, "some value"
  result = someVariable
  Exit Function
ErrorHandler:
  MsgBox "Error setting or getting value: " & Err.Description, vbCritical
  Err.Clear
  GetValue = Null ' Or handle the error appropriately
End Function