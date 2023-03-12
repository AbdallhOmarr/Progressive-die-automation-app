Imports Microsoft.VisualBasic
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Windows.Forms.LinkLabel

Public Class ReadStepClass
    Public Shared Function GetFileRowCount(ByVal filePath As String)
        Dim lineCount As Integer = 0

        If File.Exists(filePath) Then
            Using reader As New StreamReader(filePath)

                ' Read each line in the file
                While reader.ReadLine() IsNot Nothing
                    ' Increment the line count for each line
                    lineCount += 1
                End While

            End Using
            Return lineCount
        End If
        ' Open the file for reading
        ' Initialize a line count variable

    End Function
    Public Shared Function ReadMatrixFromFile(ByVal filePath As String)
        Dim rowIndex As Integer = 0
        Dim lineCount As Integer = GetFileRowCount(filePath)
        Dim matrix As Object(,) = New Object(lineCount - 1, 100) {}

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
        Return matrix
    End Function

    Public Shared Function WriteMatrixToFile(ByVal filePath As String, ByVal matrix As Object(,))
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
    End Function
    Public Shared Function WriteSingleMatrixToFile(ByVal filePath As String, ByVal matrix As String())
        Using writer As New StreamWriter(filePath)
            For i As Integer = 0 To matrix.GetUpperBound(0)
                writer.Write(matrix(i) & " ")
                writer.WriteLine()
            Next
        End Using
    End Function


    Public Shared Function GetClosedShel(ByVal filePath As String)
        '1. Read file 
        '2. Loop overeach line
        '3. if entity name == ClosedShell 
        '   1. get row 
        Dim closedShell As Object
        Using reader As New StreamReader(filePath)
            Dim line As String = ""
            While Not reader.EndOfStream
                line = reader.ReadLine()
                Console.WriteLine(line)
                If line.Contains("CLOSED_SHELL") = True Then
                    closedShell = line
                End If
            End While

        End Using
        If Not closedShell Is Nothing Then
            Return closedShell
        End If
        If closedShell Is Nothing Then
            Return ""
        End If
    End Function

    Public Shared Function GetAdvancedFaces(ByVal ClosedShell As Object, ByVal FilePath As String)
        ' getting advanced faces matrix 
        '1. ClosedShell is string 
        '2. separate string by spaces 
        '3. store it in matrix 
        '4. escape idx 0,1 
        '5. this is a matrix of advanced faces 
        Dim words As String() = ClosedShell.split(" "c)
        Dim faces As Integer() = New Integer(words.Length - 3) {}

        Dim counter As Integer = 0
        For Each word In words
            Console.WriteLine(word)
            If counter <> 0 AndAlso counter <> 1 Then
                Dim value As Integer
                If Integer.TryParse(word, value) Then
                    Console.WriteLine(word)
                    faces(counter - 2) = value
                End If

            End If
            counter += 1
        Next
        Console.WriteLine(faces.Length.ToString & "faces matrix length")
        Dim facesMatrix As String() = New String(faces.Length) {}
        '1. loop on all entities
        '2. if first number in row in faces 
        '3. add the line to facesMatrix
        Dim counter2 As Integer = 0

        Using reader As New StreamReader(FilePath)
            Dim line As String = ""
            While Not reader.EndOfStream
                line = reader.ReadLine()
                Console.WriteLine(line)
                'get first num
                Dim firstNum As Integer
                Dim value As Integer
                If Integer.TryParse(line.Split(" "c)(0), value) Then
                    Console.WriteLine(line.Split(" "c)(0))
                    firstNum = value
                    If faces.Cast(Of Integer).Contains(firstNum) AndAlso Not String.IsNullOrEmpty(line) Then 'use String.IsNullOrEmpty to check if string is not empty
                        facesMatrix(counter2) = line
                        counter2 += 1

                    End If
                End If


                'check if firstNum in faces

            End While

        End Using


        Return facesMatrix
    End Function

    Public Shared Function getBoundry(ByVal AdvancdFace As Object, ByVal FilePath As String, ByVal savedPath As String)
        ' getting boundry matrix matrix 
        '1. Advanced Face is string 
        '2. separate string by spaces 
        '3. store it in matrix 
        '4. escape idx 0,1 
        '5. this is a matrix of boundries
        If Not AdvancdFace Is Nothing Then

            Dim words As String() = AdvancdFace.split(" "c)
            Dim boundry As Integer() = New Integer(words.Length - 3) {}

            Dim counter As Integer = 0
            For Each word In words
                Console.WriteLine(word)
                If counter <> 0 AndAlso counter <> 1 Then
                    Dim value As Integer
                    If Integer.TryParse(word, value) Then
                        Console.WriteLine(word)
                        boundry(counter - 2) = value
                    End If

                End If
                counter += 1
            Next
            Console.WriteLine(boundry.Length.ToString & "faces matrix length")
            Dim boundryMatrix As String() = New String(boundry.Length) {}
            '1. loop on all entities
            '2. if first number in row in faces 
            '3. add the line to facesMatrix
            Dim counter2 As Integer = 0

            Using reader As New StreamReader(FilePath)
                Dim line As String = ""
                While Not reader.EndOfStream
                    line = reader.ReadLine()
                    Console.WriteLine(line)
                    'get first num
                    Dim firstNum As Integer
                    Dim value As Integer
                    If Integer.TryParse(line.Split(" "c)(0), value) Then
                        Console.WriteLine(line.Split(" "c)(0))
                        firstNum = value
                        If boundry.Cast(Of Integer).Contains(firstNum) AndAlso Not String.IsNullOrEmpty(line) Then 'use String.IsNullOrEmpty to check if string is not empty

                            boundryMatrix(counter2) = line
                            counter2 += 1

                        End If
                    End If


                    'check if firstNum in faces

                End While

            End Using

            WriteSingleMatrixToFile(savedPath, boundryMatrix)


            Return boundryMatrix
        End If

    End Function

    Public Shared Function GetEdgeLoop(ByVal AdvancedFaceFilePath As String, ByVal FilePath As String)
        'open advanced face txt file
        'loop on entities
        ' if entity is face bound 
        ' get edge loop 
        ' loop on the whole file 
        ' get line where edge loop is 
        ' store the line after its facebound in the advanced face txt file
        Dim edgeLoop As Integer()
        Dim lines As String()

        Using reader As New StreamReader(AdvancedFaceFilePath)
            Dim line As String = ""
            Dim idx1 As Integer = 0
            While Not reader.EndOfStream

                line = reader.ReadLine()
                MsgBox(line.GetType.ToString)
                lines(idx1) = line

                Dim words As String() = line.Split(" "c)
                edgeLoop = New Integer(words.Length - 3) {}

                If line.Contains("FACE") Or line.Contains("BOUND") Or line.Contains("OUTER") Or line.Contains("INNER") Then
                    Dim counter As Integer = 0
                    For Each word In words
                        Console.WriteLine(word)
                        If counter <> 0 AndAlso counter <> 1 Then
                            Dim value As Integer
                            If Integer.TryParse(word, value) Then
                                Console.WriteLine(word)
                                edgeLoop(counter - 2) = value
                                Using reader2 As New StreamReader(FilePath)
                                    Dim line2 As String = ""
                                    While Not reader2.EndOfStream
                                        line = reader2.ReadLine()
                                        Dim firstNum As Integer
                                        Dim value2 As Integer
                                        If Integer.TryParse(line2.Split(" "c)(0), value2) Then
                                            Console.WriteLine(line2.Split(" "c)(0))
                                            firstNum = value2

                                        End If
                                        If value = firstNum Then
                                            lines(idx1 + 1) = line2
                                            ' line2 is the required line
                                            ' we need to add it to the file at this index 
                                            ' add it at location idx1 +1
                                            ' shift the rest of the file at the next index 
                                        End If
                                    End While

                                End Using

                            End If

                        End If
                        counter += 1
                    Next

                End If
                idx1 += 1
            End While


        End Using
        File.WriteAllLines(AdvancedFaceFilePath, lines.ToArray())

    End Function

End Class

