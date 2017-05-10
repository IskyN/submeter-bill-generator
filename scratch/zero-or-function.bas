Attribute VB_Name = "Module3"
Public Function ZeroOrFunction(val, func As String, arg2) As String
    Debug.Print "original = "; val
    Dim result As String
    If val = 0 Or val = "" Then
        Debug.Print "zero"
        result = ""
    Else
        Debug.Print "not zero"
        Dim c As String
        c = func & "(" & val & ", " & """" & arg2 & """" & ")"
        result = Application.Evaluate(c)
    End If
    Debug.Print "result = "; result
    ZeroOrFunction = result
End Function


