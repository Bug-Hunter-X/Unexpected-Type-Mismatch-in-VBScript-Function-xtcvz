Function MyFunction(param1)
  'Some code here that might cause an error
  If param1 = "" Then
    Err.Raise 13, , "Type mismatch"
  End If
End Function