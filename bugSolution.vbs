Function MyFunction(param1)
  On Error Resume Next
  'Check if param1 is a string
  If VarType(param1) <> vbString Then
    Err.Raise 13, , "Parameter must be a string"
  ElseIf param1 = "" Then
    Err.Raise 13, , "Empty string not allowed"
  End If
  On Error GoTo 0
  'Rest of function code
End Function