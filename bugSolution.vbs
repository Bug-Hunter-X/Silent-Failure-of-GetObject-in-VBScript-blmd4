Function GetValue(key)
  Dim obj, errNum
  On Error Resume Next
  Set obj = GetObject("Script:" & key)
  errNum = Err.Number
  On Error GoTo 0 ' Reset error handling
  If errNum <> 0 Then
    'Handle Error Explicitly. Log, Throw Exception, or Return an appropriate value 
    WScript.Echo "Error accessing key: " & key & ". Error Number: " & errNum
    GetValue = Null  ' Or a more appropriate default value
  Else
    GetValue = obj.Value
    Set obj = Nothing
  End If
End Function