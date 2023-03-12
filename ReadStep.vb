Imports Microsoft.VisualBasic
Imports System.IO

Public Class ReadStepClass


    Private Shared Sub ReadMatrixFromFile(ByVal filePath As String, ByRef matrix As Object(,))
        Dim rowIndex As Integer = 0

        Using reader As New StreamReader(filePath)
            Dim line As String = ""

            ' read each line of the file
            While Not reader.EndOfStream
                line = reader.ReadLine()
                Console.WriteLine(line)
                ' Split the line by spaces and parentheses
                Dim tokens As String() = line.Split(New Char() {" "c, "("c, ")"c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                Console.WriteLine(tokens)
                ' Add each number to the matrix in the same row
                Dim columnIndex As Integer = 0
                Dim c As Integer = 0


                For Each token As String In tokens

                    If token.StartsWith("#") Then
                        Console.WriteLine(token)
                        Dim number As Integer = Integer.Parse(token.Substring(1))
                        matrix(rowIndex, columnIndex) = number

                        columnIndex = columnIndex + 1
                    End If
                    If token = "=" Then
                        matrix(rowIndex, columnIndex) = tokens(c + 1)
                        columnIndex = columnIndex + 1
                    End If
                    c += 1
                Next
                ' Move to the next row of the matrix
                rowIndex += 1
            End While
        End Using
    End Sub

    Private Shared Sub WriteMatrixToFile(ByVal filePath As String, ByVal matrix As Object(,))
        Using writer As New StreamWriter(filePath)
            For i As Integer = 0 To matrix.GetUpperBound(0)
                For j As Integer = 0 To matrix.GetUpperBound(1)
                    If matrix(i, j) IsNot Nothing Then
                        writer.Write(matrix(i, j).ToString() & " ")
                    Else
                        writer.Write("")
                    End If
                Next
                writer.WriteLine()
            Next
        End Using
    End Sub

    Private Shared Sub PromptUserForInput()
        Dim userInput As String
        userInput = InputBox("Please enter your name:")
    End Sub
End Class
