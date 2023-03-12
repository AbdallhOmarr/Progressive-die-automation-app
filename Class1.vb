Imports Microsoft.VisualBasic

Public Class Class1
    Dim input As String = "AXIS2_PLACEMENT_3D ( 'NONE', #27, #113, #154 ) ;"

    ' Split the input string by spaces and parentheses
    Dim tokens As String() = input.Split(New Char() {" "c, "("c, ")"c, ","c}, StringSplitOptions.RemoveEmptyEntries)

    ' Iterate over the resulting array to extract the required information
    For Each token As String In tokens
    If token.StartsWith("#") Then
    Dim number As Integer = Integer.Parse(token.Substring(1))
        Console.WriteLine($"Found #{number}")
    ElseIf token <> ";" Then
        Console.WriteLine($"Found '{token}'")
    End If
    Next

End Class
