Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports SolidWorks.Interop
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Public Class Form1

    Dim SLDPRT_FileLocation As String
    '
    Dim SolidWorks As SldWorks
    Dim Part As ModelDoc2
    Dim FeatMgr As FeatureManager
    Dim SketchMgr As SketchManager
    Dim SelectionMgr As SelectionMgr
    Dim FeatStat As FeatureStatistics
    Dim SheetMetalFeature As ISheetMetalFeatureData
    Dim Feature As Feature
    Dim BendFeature As Feature
    Dim ModView As ModelView

    Dim Material As String
    Dim Thikness As Double
    Dim UTS As Double
    Dim Image_COUNTER As Integer = 0
    Dim Planes_COUNTER As Integer = 1

    Dim application_State As Integer = 0

    Dim CircularCheckBox As Integer = 0
    Dim BlankingCheckBox As Integer = 0


    Function delay(ByVal delayTime As Integer) As Boolean
        System.Threading.Thread.Sleep(delayTime)
    End Function

    Function GetSLDPRTLocationFromUser() As String
        Dim FileBrowser As FileDialog = New OpenFileDialog()
        FileBrowser.Title = "Open Part File"
        FileBrowser.InitialDirectory = "D:\"
        FileBrowser.Filter = "SLDPRT files (*.SLDPRT)|*.SLDPRT|All files (*.*)|*.*"
        If FileBrowser.ShowDialog() = DialogResult.OK Then
            Return FileBrowser.FileName
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Clear()

        SLDPRT_FileLocation = GetSLDPRTLocationFromUser()

        If SLDPRT_FileLocation <> "" Then
            TextBox1.Text = SLDPRT_FileLocation
            'ProgressBar1.Value = 100
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
        SolidWorks = CreateObject("SldWorks.application") 'create object model 
        SolidWorks.Visible = True


        ' to save opened documents before close
        Dim opened_documents As Object
        opened_documents = SolidWorks.GetDocuments

        If Not opened_documents Is Nothing AndAlso opened_documents.Length > 0 Then

            Dim i As Integer
            For i = 0 To opened_documents.Length - 1
                Dim opened_doc As Object
                opened_doc = opened_documents(i)
                opened_doc.Save
            Next
        End If
        SolidWorks.CloseAllDocuments(True) 'close all documents opened 

        'open selected part

        Dim SLDPRT_InvalidPathFlag As Byte = False

        If SLDPRT_FileLocation <> "" Then
            Part = SolidWorks.OpenDoc(SLDPRT_FileLocation, 1)
        Else
            MsgBox("Invalid Path!")
            SLDPRT_InvalidPathFlag = True

        End If


        If SLDPRT_InvalidPathFlag = False Then
            ' if the selected part path is correct do the following

            'choose isometric view (7 refer to the view type)
            Part.ShowNamedView2("*Isometric", 7)

            'create model view object and assigning its value to be the active view
            ModView = Part.ActiveView
            ModView.DisplayMode = swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges
            Part.Extension.InsertScene("\scenes\01 basic scenes\11 white kitchen.p2s")
            Part.ViewZoomtofit2()




            ' save part image 

            'define extension string as an array of character
            Dim MyChar() As Char = (".SLDPRT")

            'define part name and trim the extension 
            Dim NAME As String = Part.GetPathName().TrimEnd(MyChar)

            'define image name and add its extension
            Dim Path = NAME & "-PartImage" & ".JPG"

            'save image at path defined
            Part.SaveAs3(Path, 2, 0)

            'view image in the picture box
            PictureBox1.Image = New Bitmap(Path)
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

            'assign values to [selection manager, feature manager, feature statistics]
            SelectionMgr = Part.SelectionManager
            FeatMgr = Part.FeatureManager
            FeatStat = FeatMgr.FeatureStatistics

            'define new variables 
            ' count of all features 
            Dim FeatureCount As Integer = FeatStat.FeatureCount

            ' array of all feature names
            Dim FeatNames(FeatureCount) As String

            'array of bends 
            Dim SketchedBend(FeatureCount) As String


            'count of bends
            Dim SketchedBendCount As Integer = 0



            'getting feature names
            For i As Integer = 0 To FeatureCount - 1
                FeatNames(i) = FeatStat.FeatureNames(i)
            Next

            For i As Integer = 0 To FeatureCount - 1
                ListBox1.Items.Add(FeatNames(i))
            Next


            'For i As Integer = 0 To FeatureCount - 1
            '    Dim Index As Integer = -1
            '    Index = FeatNames(i).IndexOf("Edge-Flange")
            '    If Index >= 0 Then
            '        SketchedBend(SketchedBendCount) = FeatNames(i)
            '        SketchedBendCount = SketchedBendCount + 1
            '    End If
            'Next


            ListBox1.Items.Add("-----------------")


            'model doc > first feature > 
            Dim bendCount As Integer = 0
            Dim feat As Feature = Part.FirstFeature
            Do While Not feat Is Nothing
                'Do something with the feature
                Console.WriteLine(feat.Name)
                If feat.Name.Contains("Flat-Pattern") Then
                    Console.WriteLine("--------")
                    Dim subFeature As Feature = feat.GetFirstSubFeature
                    Console.WriteLine("in Flat Pattern")
                    Console.WriteLine(subFeature)
                    Do While Not subFeature Is Nothing
                        ' Print the name of the sub-feature
                        Console.WriteLine("Feature Name:" & subFeature.Name & " Feature Type: " & subFeature.GetType().Name)
                        'If subFeature.Name.Contains("Bend") And subFeature.Name.Contains("Lines") Then

                        '    Dim bendLinesFeature As = CType(subFeature.GetSpecificFeature2(), SketchSegment)

                        '    Dim bendLines As Object = bendLinesFeature.GetBendLines

                        '    For i As Integer = 0 To UBound(bendLines)
                        '        Dim bendLine As SketchLine = CType(bendLines(i), SketchLine)
                        '        Dim startPoint As Double() = bendLine.StartPoint
                        '        Dim endPoint As Double() = bendLine.EndPoint
                        '        Dim length As Double = bendLine.Length
                        '        Console.WriteLine("Bend line " & (i + 1) & ": Start point=(" & startPoint(0) & "," & startPoint(1) & "), End point=(" & endPoint(0) & "," & endPoint(1) & "), Length=" & length)
                        '    Next

                        'End If
                        If subFeature.Name.Contains("Bend") And Not subFeature.Name.Contains("Lines") Then
                            SketchedBend(SketchedBendCount) = FeatNames(bendCount)
                            SketchedBendCount = SketchedBendCount + 1

                            bendCount += 1
                        End If
                        ' Get the next sub-feature of the parent feature
                        subFeature = subFeature.GetNextSubFeature

                    Loop
                    Console.WriteLine("--------")

                End If
                feat = feat.GetNextFeature
            Loop
            ListBox1.Items.Add("Number of bends = " & bendCount)



            ListBox1.Items.Add("-----------------")

            Dim BaseFlange As Byte = False
            Dim value As System.Double

            BaseFlange = Part.Extension.SelectByID2("Base-Flange1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)


            BaseFlange = Part.Extension.SelectByID2("Boss-Extrude1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)

            If BaseFlange = 255 Then
                Part.Extension.SelectByID2("Sheet-Metal", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                Feature = SelectionMgr.GetSelectedObject6(1, -1)
                SheetMetalFeature = Feature.GetDefinition
                value = SheetMetalFeature.Thickness * 1000
                Thikness = value

            Else
                Dim message, title, defaultValue As String
                Dim InputValue As Object
                message = "Enter Thickness value"
                title = "Input Thickness"
                defaultValue = "1"
                InputValue = InputBox(message, title, defaultValue)
                If InputValue Is "" Then InputValue = defaultValue
                'InputValue = InputBox(message, title, defaultValue, 100, 100)
                If InputValue Is "" Then InputValue = defaultValue
                Thikness = InputValue
                delay(500)
            End If

            ListBox1.Items.Add("Thikness " & Thikness)

            ListBox1.Items.Add("-----------------")
            Material = Part.MaterialIdName

            If Material <> "" And Material <> "Steel" Then
                Dim MaterialWordArray As String() = Material.Split(New Char() {"|"c})
                Material = MaterialWordArray(MaterialWordArray.Length - 2)
            Else
                delay(500)
                Dim message, title, defaultValue As String
                Dim InputValue As Object
                message = "Enter Part Material"
                title = "Input Material"
                defaultValue = "1.0037"
                InputValue = InputBox(message, title, defaultValue)
                If InputValue Is "" Then InputValue = defaultValue
                'InputValue = InputBox(message, title, defaultValue, 100, 100)
                If InputValue Is "" Then InputValue = defaultValue
                Material = InputValue
            End If

            ListBox1.Items.Add("Material " & Material)
            ListBox1.Items.Add("-----------------")


            Dim LinesNumber As Integer = 0
            Dim ArcsNumber As Integer = 0

            Dim Sketch As Sketch
            Dim SketchSegmentArray As Object
            Dim SegmentCounter As Object
            Dim CurrentSketchSegment As SketchSegment

            Dim State As Boolean = False
            Dim ModelState As Boolean = False

            Dim LineState As Boolean = False
            Dim ArcState As Boolean = False

            Dim SketchCounter As Integer = 1

            Dim Flag As Integer = 1

            Dim PartName As String

            Dim ExtensionChar() As Char = (".SLDPRT")
            PartName = Part.GetPathName().TrimEnd(ExtensionChar) + ".txt"

            If File.Exists(PartName) Then
                File.Delete(PartName)
            End If

            Dim SearchesNumber As Integer = 100

            While SketchCounter < SearchesNumber

                Flag = 1
                'selecting entity and if it selected will return True 
                State = Part.Extension.SelectByID2("Model", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)

                ModelState = Part.Extension.SelectByID2("Model", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                Console.WriteLine("state")
                Console.WriteLine(State)


                If State = True Then
                    Flag = 0
                    ListBox1.Items.Add("Model")
                End If

                If State <> True Then
                    While State <> True
                        State = Part.Extension.SelectByID2("Sketch" & SketchCounter, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
                        SketchCounter = SketchCounter + 1

                        If State = True Then
                            Flag = 0
                            ListBox1.Items.Add("Sketch" & SketchCounter - 1)
                        End If

                        If SketchCounter > SearchesNumber Then
                            Exit While
                            Flag = 1
                        End If

                    End While
                End If

                If SketchCounter > SearchesNumber Then
                    Exit While
                    Flag = 1
                End If

                If Flag = 0 Then

                    'edit sketch 
                    Part.EditSketch()
                    'clear selection
                    Part.ClearSelection2(True)
                    'getting active sketch to access its segments(lines,circles,..)
                    Sketch = Part.GetActiveSketch2()
                    'getting sketch segment
                    SketchSegmentArray = Sketch.GetSketchSegments

                    'arrays to save line with startpoint and end point
                    Dim LinePointsArray(5) As Double

                    Dim ConstructionLinePointsArray(5) As Double
                    'array to save circle center point(x,y)
                    Dim CirclePointsArray(2) As Double
                    'array to store radius of array
                    Dim Radius As Double = 0

                    Dim ArcPointsArray(6) As Double
                    'array 
                    Dim PointArray As System.Object
                    'sketch segment
                    Dim SelectedSegment As SketchSegment
                    Dim SelectedLine As SketchLine
                    Dim SelectedArc As SketchArc

                    If Not IsNothing(SketchSegmentArray) Then
                        For Each SegmentCounter In SketchSegmentArray
                            CurrentSketchSegment = SegmentCounter
                            Console.WriteLine(CurrentSketchSegment.GetType)
                            If swSketchSegments_e.swSketchLINE = CurrentSketchSegment.GetType Then
                                LinesNumber = LinesNumber + 1
                            End If
                        Next
                    End If

                    Dim SelectedLines As Integer = 0
                    Dim LineCounter As Integer = 1
                    Dim ConstructionLineFlag As Boolean = False

                    While SelectedLines <> LinesNumber
                        LineState = Part.Extension.SelectByID2("Line" & LineCounter, "SKETCHSEGMENT", 0, 0, 0, False, 2, Nothing, 0)
                        If LineState = True Then
                            SelectedSegment = SelectionMgr.GetSelectedObject(1)
                            SelectedLine = SelectedSegment

                            ConstructionLineFlag = False

                            If SelectedSegment.ConstructionGeometry = True Then
                                ConstructionLineFlag = True
                            End If

                            If ConstructionLineFlag = False Then
                                PointArray = SelectedLine.GetStartPoint()
                                LinePointsArray(0) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                LinePointsArray(1) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                PointArray = SelectedLine.GetEndPoint()
                                LinePointsArray(2) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                LinePointsArray(3) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                Console.WriteLine("Line Length")
                                LinePointsArray(4) = SelectedLine.GetLength * 1000

                                ListBox1.Items.Add("Line " & LinePointsArray(0) & " " & LinePointsArray(1) & " " & LinePointsArray(2) & " " & LinePointsArray(3) & " " & LinePointsArray(4))

                                File.AppendAllText(PartName, "Line " & LinePointsArray(0) & " " & LinePointsArray(1) & " " & LinePointsArray(2) & " " & LinePointsArray(3) & " " & LinePointsArray(4) & System.Environment.NewLine)

                                SelectedLines = SelectedLines + 1
                            End If

                            If ConstructionLineFlag = True Then
                                PointArray = SelectedLine.GetStartPoint()
                                ConstructionLinePointsArray(0) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                ConstructionLinePointsArray(1) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                PointArray = SelectedLine.GetEndPoint()
                                ConstructionLinePointsArray(2) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                ConstructionLinePointsArray(3) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                Console.WriteLine("Z: " & PointArray(2))
                                ListBox1.Items.Add("ConstructionLine " & LinePointsArray(0) & " " & LinePointsArray(1) & " " & LinePointsArray(2) & " " & LinePointsArray(3) & " " & LinePointsArray(4))

                                File.AppendAllText(PartName, "ConstructionLine " & LinePointsArray(0) & " " & LinePointsArray(1) & " " & LinePointsArray(2) & " " & LinePointsArray(3) & " " & LinePointsArray(4) & System.Environment.NewLine)


                                SelectedLines = SelectedLines + 1
                            End If


                        End If

                        LineCounter = LineCounter + 1
                    End While

                    If Not IsNothing(SketchSegmentArray) Then
                        For Each SegmentCounter In SketchSegmentArray
                            CurrentSketchSegment = SegmentCounter
                            If swSketchSegments_e.swSketchARC = CurrentSketchSegment.GetType Then
                                ArcsNumber = ArcsNumber + 1
                            End If
                        Next
                    End If

                    Dim SelectedArcs As Integer = 0
                    Dim ArcCounter As Integer = 1

                    Dim ConstructionArcFlag As Boolean = False

                    While SelectedArcs <> ArcsNumber

                        ArcState = Part.Extension.SelectByID2("Arc" & ArcCounter, "SKETCHSEGMENT", 0, 0, 0, False, 2, Nothing, 0)
                        If ArcState = True Then
                            SelectedSegment = SelectionMgr.GetSelectedObject(1)
                            SelectedArc = SelectedSegment

                            ConstructionArcFlag = False

                            If SelectedSegment.ConstructionGeometry = True Then
                                ConstructionArcFlag = True
                            End If

                            If SelectedArc.IsCircle() Then

                                If ConstructionArcFlag = False Then
                                    PointArray = SelectedArc.GetCenterPoint()
                                    CirclePointsArray(0) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    CirclePointsArray(1) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                    Radius = SelectedArc.GetRadius()
                                    CirclePointsArray(2) = Math.Round(Radius * 1000, 1, MidpointRounding.AwayFromZero)

                                    ListBox1.Items.Add("Circle " & CirclePointsArray(0) & " " & CirclePointsArray(1) & " " & CirclePointsArray(2))

                                    File.AppendAllText(PartName, "Circle " & CirclePointsArray(0) & " " & CirclePointsArray(1) & " " & CirclePointsArray(2) & System.Environment.NewLine)
                                End If

                                If ConstructionArcFlag = True Then
                                    PointArray = SelectedArc.GetCenterPoint()
                                    CirclePointsArray(0) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    CirclePointsArray(1) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                    Radius = SelectedArc.GetRadius()
                                    CirclePointsArray(2) = Math.Round(Radius * 1000, 1, MidpointRounding.AwayFromZero)

                                    ListBox1.Items.Add("ConstructionCircle " & CirclePointsArray(0) & " " & CirclePointsArray(1) & " " & CirclePointsArray(2))

                                    File.AppendAllText(PartName, "ConstructionCircle " & CirclePointsArray(0) & " " & CirclePointsArray(1) & " " & CirclePointsArray(2) & System.Environment.NewLine)
                                End If
                            Else

                                If ConstructionArcFlag = False Then

                                    PointArray = SelectedArc.GetCenterPoint()
                                    ArcPointsArray(0) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    ArcPointsArray(1) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                    PointArray = SelectedArc.GetStartPoint()
                                    ArcPointsArray(2) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    ArcPointsArray(3) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                    PointArray = SelectedArc.GetEndPoint()
                                    ArcPointsArray(4) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    ArcPointsArray(5) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)

                                    ListBox1.Items.Add("Arc " & ArcPointsArray(0) & " " & ArcPointsArray(1) & " " & ArcPointsArray(2) & " " & ArcPointsArray(3) & " " & ArcPointsArray(4) & " " & ArcPointsArray(5))

                                    File.AppendAllText(PartName, "Arc " & ArcPointsArray(0) & " " & ArcPointsArray(1) & " " & ArcPointsArray(2) & " " & ArcPointsArray(3) & " " & ArcPointsArray(4) & " " & ArcPointsArray(5) & System.Environment.NewLine)
                                End If

                                If ConstructionArcFlag = True Then

                                    PointArray = SelectedArc.GetCenterPoint()
                                    ArcPointsArray(0) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    ArcPointsArray(1) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                    PointArray = SelectedArc.GetStartPoint()
                                    ArcPointsArray(2) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    ArcPointsArray(3) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                                    PointArray = SelectedArc.GetEndPoint()
                                    ArcPointsArray(4) = Math.Round(PointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                                    ArcPointsArray(5) = Math.Round(PointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)

                                    ListBox1.Items.Add("ConstructionArc " & ArcPointsArray(0) & " " & ArcPointsArray(1) & " " & ArcPointsArray(2) & " " & ArcPointsArray(3) & " " & ArcPointsArray(4) & " " & ArcPointsArray(5))

                                    File.AppendAllText(PartName, "ConstructionArc " & ArcPointsArray(0) & " " & ArcPointsArray(1) & " " & ArcPointsArray(2) & " " & ArcPointsArray(3) & " " & ArcPointsArray(4) & " " & ArcPointsArray(5) & System.Environment.NewLine)
                                End If

                            End If

                            SelectedArcs = SelectedArcs + 1
                        End If

                        ArcCounter = ArcCounter + 1
                    End While
                End If

                Part.SketchManager.InsertSketch(True)
                State = False
                LinesNumber = 0
                ArcsNumber = 0

                If ModelState = True Then
                    Exit While
                End If

            End While

            For i As Integer = 0 To SketchedBendCount - 1

                ListBox1.Items.Add(SketchedBend(i))

                Part.Extension.SelectByID2(SketchedBend(i), "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
                Part.FeatEdit()

                Dim BendSketch As Sketch = Part.GetActiveSketch()
                Dim BendLinesArray As Object = BendSketch.GetSketchSegments()
                Dim BendLineCounterFor As Object
                Dim BendLine As SketchLine
                Dim BendLinePointArray As System.Object

                Dim CurrentBendLine As SketchSegment
                Dim SelectedBendLine As SketchSegment

                Dim BendLineLinesNumber As Integer = 0

                If Not IsNothing(BendLinesArray) Then
                    For Each BendLineCounterFor In BendLinesArray
                        CurrentBendLine = BendLineCounterFor
                        If swSketchSegments_e.swSketchLINE = CurrentBendLine.GetType Then
                            BendLineLinesNumber = BendLineLinesNumber + 1
                        End If
                    Next
                End If

                Dim SelectedBendLines As Integer = 0
                Dim BendLineState As Boolean = False
                Dim BendLineCounter As Integer = 1
                Dim BendLinePointsArray(4) As Double

                While SelectedBendLines <> BendLineLinesNumber
                    BendLineState = Part.Extension.SelectByID2("Line" & BendLineCounter, "SKETCHSEGMENT", 0, 0, 0, False, 2, Nothing, 0)

                    If BendLineState = True Then
                        SelectedBendLine = SelectionMgr.GetSelectedObject(1)
                        BendLine = SelectedBendLine

                        Dim ConstructionBendLineFlag = False

                        If SelectedBendLine.ConstructionGeometry = True Then
                            ConstructionBendLineFlag = True
                            SelectedBendLines = SelectedBendLines + 1
                        End If

                        If SelectedBendLine.ConstructionGeometry = False Then

                            BendLinePointArray = BendLine.GetStartPoint()
                            BendLinePointsArray(0) = Math.Round(BendLinePointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                            BendLinePointsArray(1) = Math.Round(BendLinePointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)
                            BendLinePointArray = BendLine.GetEndPoint()
                            BendLinePointsArray(2) = Math.Round(BendLinePointArray(0) * 1000, 1, MidpointRounding.AwayFromZero)
                            BendLinePointsArray(3) = Math.Round(BendLinePointArray(1) * 1000, 1, MidpointRounding.AwayFromZero)

                            ListBox1.Items.Add("Bend Line " & BendLinePointsArray(0) & " " & BendLinePointsArray(1) & " " & BendLinePointsArray(2) & " " & BendLinePointsArray(3))

                            File.AppendAllText(PartName, "BendLine " & BendLinePointsArray(0) & " " & BendLinePointsArray(1) & " " & BendLinePointsArray(2) & " " & BendLinePointsArray(3) & System.Environment.NewLine)

                            SelectedBendLines = SelectedBendLines + 1
                        End If

                    End If

                    BendLineCounter = BendLineCounter + 1
                End While

                Part.SketchManager.InsertSketch(True)
            Next

            File.AppendAllText(PartName, "Material " & Material & System.Environment.NewLine)

            'TextBox3.Clear()
            'TextBox3.Text = Material

            File.AppendAllText(PartName, "Thikness " & Thikness) '& System.Environment.NewLine)

            'ProgressBar1.Value = 100
            MsgBox("Part Data has been extracted!")










        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class
