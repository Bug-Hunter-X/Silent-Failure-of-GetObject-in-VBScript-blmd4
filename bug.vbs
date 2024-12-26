Function GetValue(key)
  On Error Resume Next
  Set obj = GetObject("Script:" & key)
  If Err.Number <> 0 Then
    Err.Clear
    GetValue = Null
  Else
    GetValue = obj.Value
    Set obj = Nothing
  End If
End Function