Function MyFunction(param1, param2)
  On Error GoTo ErrorHandler
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbError, , "Parameters cannot be empty"
    Exit Function ' Crucial to prevent further execution after raising an error
  End If
  ' ...rest of function...
  Exit Function
ErrorHandler:
  MsgBox "Error: " & Err.Number & " - " & Err.Description
  Err.Clear
End Function