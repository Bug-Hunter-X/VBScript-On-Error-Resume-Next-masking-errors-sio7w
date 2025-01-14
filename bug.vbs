Function GetValue(key)
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Resume Next
  SetValue key, "some value"
  If Err.Number <> 0 Then
    Err.Clear
  End If
  GetValue = someVariable
End Function