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


            For i As Integer = 0 To FeatureCount - 1
                Dim Index As Integer = -1
                Index = FeatNames(i).IndexOf("Edge-Flange")
                If Index >= 0 Then
                    SketchedBend(SketchedBendCount) = FeatNames(i)
                    SketchedBendCount = SketchedBendCount + 1
                End If
            Next


            ListBox1.Items.Add("-----------------")
            ListBox1.Items.Add("Sketched Bend Count " & SketchedBendCount)

            For i As Integer = 0 To SketchedBendCount - 1
                ListBox1.Items.Add(SketchedBend(i))
            Next

            ListBox1.Items.Add("-----------------")

            Dim FlatPatternFolder As Object = FeatMgr.GetFlatPatternFolder()
            ' trying to get bending count 
            Dim FlatPatternFeature As Feature = FlatPatternFolder.GetFeature
            'Dim FlatPatternStats As Object = FlatPatternFeature.GetStatistics()
            Dim FlatPatternCount As Integer = FlatPatternFolder.GetFlatPatternCount
            Dim featArray() As Object = FlatPatternFolder.GetFlatPatterns
            Dim FlatPatternNames(FlatPatternCount) As String

            'getting feature names
            For i As Integer = 0 To FlatPatternCount - 1
                FlatPatternNames(i) = featArray(i).Name
            Next

            For i As Integer = 0 To FlatPatternCount - 1
                ListBox1.Items.Add(FlatPatternNames(i))
            Next
            Console.WriteLine("feature name")
            Console.WriteLine(FlatPatternFeature.Name)

            Dim MacroFeatureNames As New List(Of String)
            Dim SubFeature As Feature = FlatPatternFeature.feature
            While Not SubFeature Is Nothing
                Dim featName2 As String = SubFeature.Name
                Console.WriteLine(featName2)
                If featName2.Contains("Flatten") Then
                    MacroFeatureNames.Add(featName2)
                    ListBox1.Items.Add(featName2)
                End If
                SubFeature = SubFeature.GetNextSubFeature()
            End While





        End If



    End Sub
End Class
