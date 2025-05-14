Imports System.Drawing.Imaging
Imports System.IO

Public Class Form1
    Private oAppSetting As New AppSetting()
    Private startPoint As PointF
    Private isDragging As Boolean = False
    Private currentRect As RectangleF
    Private rectangleList As New Hashtable

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oAppSetting.LoadData()
        SetComboBox(cbAiService, oAppSetting.GetValue("AiService", "gpt-4o"))
        txtAnthropicApiKey.Text = oAppSetting.GetValue("AnthropicApiKey")
        txtOpenAIKey.Text = oAppSetting.GetValue("OpenAIKey")
        txtGhostscriptPath.Text = oAppSetting.GetValue("txtGhostscriptPath", txtGhostscriptPath.Text)
        txtInputFile.Text = oAppSetting.GetValue("InputFile", txtInputFile.Text)
        txtDotsPerInch.Text = oAppSetting.GetValue("DotsPerInch", txtDotsPerInch.Text)
        SplitContainer1.SplitterDistance = oAppSetting.GetValue("SplitterDistance", SplitContainer1.SplitterDistance)

        If txtInputFile.Text <> "" Then
            If IO.File.Exists(txtInputFile.Text) = False Then
                'PDF File No longet exists
                txtInputFile.Text = ""
            End If
        End If

        LoadPageList()
        SelectGridCell()
        ListExtractedImages()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        oAppSetting.SetValue("AiService", GetComboBoxVal(cbAiService, "gpt-4o"))
        oAppSetting.SetValue("AnthropicApiKey", txtAnthropicApiKey.Text)
        oAppSetting.SetValue("OpenAIKey", txtOpenAIKey.Text)
        oAppSetting.SetValue("GhostscriptPath", txtGhostscriptPath.Text)
        oAppSetting.SetValue("InputFile", txtInputFile.Text)
        oAppSetting.SetValue("DotsPerInch", txtDotsPerInch.Text)
        oAppSetting.SetValue("SplitterDistance", SplitContainer1.SplitterDistance)
        oAppSetting.SaveData()
    End Sub


    Private Function GetComboBoxVal(ByRef oComboBox As ComboBox, sDefaultValue As String) As String
        If oComboBox.SelectedIndex = -1 Then
            Return sDefaultValue
        End If

        Return oComboBox.Items(oComboBox.SelectedIndex)
    End Function

    Private Sub SetComboBox(ByRef oComboBox As ComboBox, sValue As String)
        For i As Integer = 0 To oComboBox.Items.Count - 1
            If oComboBox.Items(i) = sValue Then
                oComboBox.SelectedIndex = i
                Exit Sub
            End If
        Next
    End Sub

    Private Sub btnApiKeyShow_Click(sender As Object, e As EventArgs) Handles btnApiKeyShow.Click
        If txtAnthropicApiKey.PasswordChar = "*" Then
            txtAnthropicApiKey.PasswordChar = ""
        Else
            txtAnthropicApiKey.PasswordChar = "*"
        End If
    End Sub

    Private Sub btnOpenAIKeyShow_Click(sender As Object, e As EventArgs) Handles btnOpenAIKeyShow.Click
        If txtOpenAIKey.PasswordChar = "*" Then
            txtOpenAIKey.PasswordChar = ""
        Else
            txtOpenAIKey.PasswordChar = "*"
        End If
    End Sub

    Private Sub btnInputFile_Click(sender As Object, e As EventArgs) Handles btnInputFile.Click
        OpenFileDialog1.FileName = txtInputFile.Text
        OpenFileDialog1.Title = "Open PDF File"
        OpenFileDialog1.Filter = "pdf files|*.pdf"
        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName = "" Then
            Exit Sub
        End If

        txtInputFile.Text = OpenFileDialog1.FileName
        LoadPageList()
        lbImages.Items.Clear()
    End Sub

    Private Sub btnGhostscriptPath_Click(sender As Object, e As EventArgs) Handles btnGhostscriptPath.Click
        OpenFileDialog1.FileName = txtGhostscriptPath.Text
        OpenFileDialog1.Title = "Open EXE File"
        OpenFileDialog1.Filter = "EXE files|*.exe"
        OpenFileDialog1.ShowDialog()

        If OpenFileDialog1.FileName <> "" Then
            txtGhostscriptPath.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub btnLoadPdf_Click(sender As Object, e As EventArgs) Handles btnLoadPdf.Click

        If txtGhostscriptPath.Text = "" Then
            MsgBox("Ghost script file is blank")
            Exit Sub
        End If

        Dim inputFilePath As String = txtInputFile.Text
        Dim inputExtension As String = Path.GetExtension(txtInputFile.Text)
        If inputExtension <> ".pdf" Then
            MsgBox("File is Not PDF")
            Exit Sub
        End If

        Dim outputfolder As String = GetOutputFolder()
        If IO.Directory.Exists(outputfolder) = False Then
            IO.Directory.CreateDirectory(outputfolder)
        End If

        Log("Converting PDF to jpeg files")

        Dim sDotsPerInch As String = txtDotsPerInch.Text
        Dim sArguments As String = "-sDEVICE=jpeg -dBATCH -dNOPAUSE -r" & sDotsPerInch & " -sOutputFile=""" & outputfolder & "\%03d.jpg"" """ & inputFilePath & """ -c quit"
        Dim sError As String = RunDosCommandAsynch(txtGhostscriptPath.Text, sArguments, 60 * 10)
        If sError <> "" Then
            Log(sError)
            Exit Sub
        End If

        LoadPageList()
        SelectGridCell()

        Log("Created " & DataGridView1.RowCount & " jpeg files")

    End Sub

    Function GetSelectedRowIndex() As Integer
        If DataGridView1.SelectedRows.Count > 0 Then
            Return DataGridView1.SelectedRows(0).Index
        ElseIf DataGridView1.SelectedCells.Count > 0 Then
            Return DataGridView1.SelectedCells(0).RowIndex
        End If
        Return -1
    End Function

    Function GetSelectedPage() As String
        Dim iSelectedRowIndex As Integer = GetSelectedRowIndex()
        If iSelectedRowIndex = -1 Then
            Return ""
        End If

        Dim oRow As DataGridViewRow = DataGridView1.Rows(iSelectedRowIndex)
        Return oRow.Cells("Page").Value & ""
    End Function

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        SelectGridCell()
    End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        SelectGridCell()
    End Sub

    Sub SelectGridCell()
        Dim sPage As String = GetSelectedPage()
        If sPage = "" Then
            Exit Sub
        End If

        Dim outputfolder As String = GetOutputFolder()
        If outputfolder = "" Then
            Exit Sub
        End If

        Dim sImageFilePath As String = Path.Combine(outputfolder, sPage) & ".jpg"
        If IO.File.Exists(sImageFilePath) = False Then
            Exit Sub
        End If

        If Not PictureBox1.Image Is Nothing Then
            PictureBox1.Image.Dispose()
        End If

        Dim oImage = Image.FromFile(sImageFilePath)
        PictureBox1.Image = oImage
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

        If zoomFactor <> 1.0 Then
            ResetZoom()
        End If

    End Sub

    Function GetOutputFolder() As String
        Dim inputFilePath As String = txtInputFile.Text
        If inputFilePath = "" Then Return ""
        Dim inputfolder As String = Path.GetDirectoryName(inputFilePath)
        Dim inputFileName As String = Path.GetFileNameWithoutExtension(inputFilePath)
        Return Path.Combine(inputfolder, inputFileName)
    End Function

    Sub LoadPageList()
        Dim iRowIndex As Integer = GetSelectedRowIndex()

        Dim oTable As Data.DataTable = GetDataTable()
        DataGridView1.DataSource = oTable
        DataGridView1.Update()
        DataGridView1.AllowUserToAddRows = False

        If iRowIndex <> -1 And iRowIndex < DataGridView1.RowCount Then
            DataGridView1.MultiSelect = False
            DataGridView1.Rows(iRowIndex).Cells(0).Selected = True
        End If

    End Sub

    Private Function GetDataTable() As Data.DataTable

        Dim outputfolder As String = GetOutputFolder()

        If IO.Directory.Exists(outputfolder) = False Then
            Return Nothing
        End If

        Dim oTable As New Data.DataTable
        oTable.Columns.Add(New Data.DataColumn("Page"))
        oTable.Columns.Add(New Data.DataColumn("Selection", System.Type.GetType("System.Int64")))

        For Each file As String In Directory.GetFiles(outputfolder, "*.jpg")
            Dim oDataRow As DataRow = oTable.NewRow()
            Dim sPage As String = Path.GetFileNameWithoutExtension(file)
            oDataRow("Page") = sPage

            If rectangleList.ContainsKey(sPage) Then
                oDataRow("Selection") = rectangleList(sPage).Count
            End If

            oTable.Rows.Add(oDataRow)
        Next

        Return oTable
    End Function

    Function RunDosCommandAsynch(sExeFilePath As String, sArguments As String, iTimeOutSec As Integer) As String

        Dim sError As String = ""
        Dim oProcess As Process = New Process()
        oProcess.StartInfo.UseShellExecute = False
        oProcess.StartInfo.RedirectStandardOutput = True
        oProcess.StartInfo.RedirectStandardError = True
        oProcess.StartInfo.FileName = sExeFilePath
        oProcess.StartInfo.Arguments = sArguments
        oProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        oProcess.StartInfo.CreateNoWindow = True

        oProcess.Start()
        oProcess.WaitForExit(1000 * iTimeOutSec)
        If oProcess.HasExited = False Then
            oProcess.Kill()
            sError = "Timeout after " + iTimeOutSec + " seconds."
        End If

        If String.IsNullOrEmpty(sError) Then
            sError = oProcess.StandardError.ReadToEnd()
        End If

        If oProcess.ExitCode <> 0 AndAlso String.IsNullOrEmpty(sError) Then
            sError = "ExitCode: " + oProcess.ExitCode
        End If

        oProcess.Close()

        Return sError
    End Function

    ' Event handler for mouse down event
    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        If e.Button = MouseButtons.Left Then
            ' Store the starting point as a ratio
            startPoint = New PointF(e.X / PictureBox1.Width, e.Y / PictureBox1.Height)
            isDragging = True
            currentRect = New RectangleF()
        End If
    End Sub

    ' Event handler for mouse move event
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If isDragging Then
            ' Calculate the current rectangle as ratios
            currentRect = New RectangleF(
                Math.Min(startPoint.X, e.X / PictureBox1.Width),
                Math.Min(startPoint.Y, e.Y / PictureBox1.Height),
                Math.Abs(startPoint.X - e.X / PictureBox1.Width),
                Math.Abs(startPoint.Y - e.Y / PictureBox1.Height))

            ' Refresh the PictureBox to trigger the Paint event
            PictureBox1.Invalidate()
        End If
    End Sub

    ' Event handler for mouse up event
    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        If isDragging Then
            isDragging = False

            If CInt(currentRect.Width * PictureBox1.Width) > 0 AndAlso Int(currentRect.Height * PictureBox1.Height) > 0 Then

                Dim rectangles As New List(Of RectangleF)()
                Dim sPage As String = GetSelectedPage()
                If rectangleList.ContainsKey(sPage) Then
                    rectangles = rectangleList(sPage)
                End If

                ' Add the current rectangle to the list as ratios
                rectangles.Add(currentRect)
                rectangleList(sPage) = rectangles
                LoadPageList()

                ' Refresh the PictureBox to trigger the Paint event
                PictureBox1.Invalidate()

            End If
        End If
    End Sub

    ' Redraw the rectangles if PictureBox is resized or invalidated
    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint

        Dim rectangles As New List(Of RectangleF)()
        Dim sPage As String = GetSelectedPage()
        If rectangleList.ContainsKey(sPage) Then
            rectangles = rectangleList(sPage)
        End If

        ' Draw all rectangles from the list with scaling
        For Each rect As RectangleF In rectangles
            Dim scaledRect As New RectangleF(
                rect.X * PictureBox1.Width,
                rect.Y * PictureBox1.Height,
                rect.Width * PictureBox1.Width,
                rect.Height * PictureBox1.Height)
            e.Graphics.DrawRectangle(Pens.Red, Rectangle.Round(scaledRect))
        Next

        ' Draw the current rectangle if dragging
        If isDragging Then
            Dim scaledRect As New RectangleF(
                currentRect.X * PictureBox1.Width,
                currentRect.Y * PictureBox1.Height,
                currentRect.Width * PictureBox1.Width,
                currentRect.Height * PictureBox1.Height)
            e.Graphics.DrawRectangle(Pens.Red, Rectangle.Round(scaledRect))
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        Dim sPage As String = GetSelectedPage()
        If rectangleList.ContainsKey(sPage) Then
            rectangleList.Remove(sPage)
            LoadPageList()
        End If

        PictureBox1.Invalidate()
    End Sub

    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        rectangleList = New Hashtable
        LoadPageList()
        PictureBox1.Invalidate()
    End Sub



    Public Sub ExtractSubImage(originalImagePath As String, outputImagePath As String, rectangle As RectangleF)
        ' Load the original image
        Dim originalImage As Image = Image.FromFile(originalImagePath)

        ' Calculate the actual rectangle based on the image size
        Dim actualRect As New Rectangle(
        CInt(rectangle.X * originalImage.Width),
        CInt(rectangle.Y * originalImage.Height),
        CInt(rectangle.Width * originalImage.Width),
        CInt(rectangle.Height * originalImage.Height))

        ' Extract the subimage
        Dim subImage As New Bitmap(actualRect.Width, actualRect.Height)

        Using g As Graphics = Graphics.FromImage(subImage)
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            g.DrawImage(originalImage, New Rectangle(0, 0, actualRect.Width, actualRect.Height), actualRect, GraphicsUnit.Pixel)
        End Using

        Dim sAiService As String = GetComboBoxVal(cbAiService, "gpt-4o")
        If L(sAiService, 4) = "gpt-" Then
            'short side of the image should be less than 768px and the long side should be less than 2,000px.
            If subImage.Width > subImage.Height Then
                If subImage.Width > 2000 Or subImage.Height > 768 Then
                    subImage = ResizeImage(subImage, 2000, 768, originalImagePath)
                End If
            Else
                If subImage.Width > 768 Or subImage.Height > 2000 Then
                    subImage = ResizeImage(subImage, 768, 2000, originalImagePath)
                End If
            End If

        Else 'Anthropic has a list of acceptable image sizes
            'https://docs.anthropic.com/en/docs/build-with-claude/vision#how-to-use-vision
            subImage = ResizeAnthropicImage(subImage, originalImagePath)
        End If

        ' Save the subimage as a new jpg file
        subImage.Save(outputImagePath, ImageFormat.Jpeg)

        ' Dispose of the original image and subimage
        originalImage.Dispose()
        subImage.Dispose()
    End Sub

    Function ResizeAnthropicImage(originalImage As Bitmap, originalImagePath As String) As Bitmap

        Dim avaibleSizes As New Dictionary(Of String, Size) From {
                {"1:1", New Size(1092, 1092)},
                {"3:4", New Size(951, 1268)},
                {"2:3", New Size(896, 1344)},
                {"9:16", New Size(819, 1456)},
                {"1:2", New Size(784, 1568)}
                }

        Dim sizes As New Dictionary(Of String, Size)
        For Each kvp As KeyValuePair(Of String, Size) In avaibleSizes
            Dim size As Size = kvp.Value
            If originalImage.Width > size.Width AndAlso originalImage.Height > size.Height Then
                sizes.Add(kvp.Key, size)
            End If
        Next

        If sizes.Count = 0 Then
            'All sized are too small - image needs to be resized
            sizes = avaibleSizes
        End If

        ' Load the image
        Dim originalAspectRatio As Double = originalImage.Width / originalImage.Height

        ' Calculate the aspect ratio and size differences
        Dim aspectRatioDifferences As New List(Of Double)
        Dim widthDifferences As New List(Of Double)
        Dim heightDifferences As New List(Of Double)

        For Each kvp As KeyValuePair(Of String, Size) In sizes
            Dim size As Size = kvp.Value
            Dim targetAspectRatio As Double = size.Width / size.Height
            Dim aspectRatioDifference As Double = Math.Abs(originalAspectRatio - targetAspectRatio)
            Dim widthDifference As Double = Math.Abs(originalImage.Width - size.Width)
            Dim heightDifference As Double = Math.Abs(originalImage.Height - size.Height)

            aspectRatioDifferences.Add(aspectRatioDifference)
            widthDifferences.Add(widthDifference)
            heightDifferences.Add(heightDifference)
        Next

        ' Normalize the differences using softmax
        Dim normalizedAspectRatios As List(Of Double) = Softmax(aspectRatioDifferences)
        Dim normalizedWidths As List(Of Double) = Softmax(widthDifferences)
        Dim normalizedHeights As List(Of Double) = Softmax(heightDifferences)

        ' Combine the normalized scores
        Dim combinedScores As New Dictionary(Of String, Double)
        Dim index As Integer = 0
        For Each kvp As KeyValuePair(Of String, Size) In sizes
            combinedScores(kvp.Key) = normalizedAspectRatios(index) * 10.0 + normalizedWidths(index) + normalizedHeights(index)
            index += 1
        Next

        ' Find the size with the highest combined score
        Dim bestSize As Size = sizes(combinedScores.OrderBy(Function(x) x.Value).First().Key)
        Dim sProblems As String = ""

        If originalImage.Width > bestSize.Width Then
            sProblems = "Width will be reduced from " & originalImage.Width & " to " & bestSize.Width
        End If

        If originalImage.Height > bestSize.Height Then
            If sProblems <> "" Then sProblems += ". "
            sProblems += "Height will be reduced from " & originalImage.Height & " to " & bestSize.Height
        End If

        If sProblems <> "" Then
            Log(Path.GetFileName(originalImagePath) & " " & sProblems)
        End If

        ' Resize the image
        Dim resizedImage As New Bitmap(bestSize.Width, bestSize.Height)
        Using graphics As Graphics = Graphics.FromImage(resizedImage)
            graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            graphics.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            graphics.DrawImage(originalImage, 0, 0, bestSize.Width, bestSize.Height)
        End Using

        originalImage.Dispose()

        Return resizedImage
    End Function

    Function Softmax(values As List(Of Double)) As List(Of Double)
        Dim maxVal As Double = values.Max()
        Dim expValues As List(Of Double) = values.Select(Function(x) Math.Exp(x - maxVal)).ToList()
        Dim sumExpValues As Double = expValues.Sum()
        Return expValues.Select(Function(x) x / sumExpValues).ToList()
    End Function

    Public Function ResizeImage(originalImage As Bitmap, maxWidth As Integer, maxHeight As Integer, originalImagePath As String) As Bitmap
        ' Get the original width and height of the image
        Dim originalWidth As Integer = originalImage.Width
        Dim originalHeight As Integer = originalImage.Height
        Dim sProblems As String = ""

        If originalImage.Width > maxWidth Then
            sProblems = "Width will be reduced from " & originalImage.Width & " to " & maxWidth
        End If

        If originalImage.Height > maxHeight Then
            If sProblems <> "" Then sProblems += ". "
            sProblems += "Height will be reduced from " & originalImage.Height & " to " & maxHeight
        End If

        If sProblems <> "" Then
            Log(Path.GetFileName(originalImagePath) & " " & sProblems)
        End If

        ' Calculate the ratio of the width and height
        Dim ratioX As Double = maxWidth / originalWidth
        Dim ratioY As Double = maxHeight / originalHeight

        ' Determine the ratio that will allow the image to fit within the specified dimensions while maintaining the aspect ratio
        Dim ratio As Double = Math.Min(ratioX, ratioY)

        ' Calculate the new width and height based on the ratio
        Dim newWidth As Integer = CInt(originalWidth * ratio)
        Dim newHeight As Integer = CInt(originalHeight * ratio)

        ' Create a new Bitmap object with the new dimensions
        Dim resizedImage As New Bitmap(newWidth, newHeight)

        ' Use a Graphics object to draw the resized image
        Using graphics As Graphics = Graphics.FromImage(resizedImage)
            ' Set the quality of the resize operation
            graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            graphics.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

            ' Draw the original image onto the resized image
            graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight)
        End Using

        originalImage.Dispose()

        ' Return the resized image
        Return resizedImage
    End Function

    Private Function L(s As String, i As Integer) As String
        Return Microsoft.VisualBasic.Left(s, i)
    End Function

    Private Sub llAnthropic_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles llAnthropic.LinkClicked
        Process.Start(New ProcessStartInfo("https://console.anthropic.com"))
    End Sub

    Private Sub llOpenAI_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles llOpenAI.LinkClicked
        Process.Start(New ProcessStartInfo("https://platform.openai.com"))
    End Sub

    Private Sub llGS_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles llGS.LinkClicked
        Process.Start(New ProcessStartInfo("https://ghostscript.com/releases/gsdnld.html"))
    End Sub


    Private Sub btnCreateExcel_Click(sender As Object, e As EventArgs) Handles btnCreateExcel.Click

        Dim outputfolder As String = GetOutputFolder()
        If IO.Directory.Exists(outputfolder) = False Then
            MsgBox("PDF is not loaded")
            Exit Sub
        End If

        If Directory.GetFiles(outputfolder, "*.jpg").Length = 0 Then
            MsgBox("PDF is not loaded")
            Exit Sub
        End If

        Log("Extracting images from selections...")

        Try
            ExtractSubImages()
        Catch ex As Exception
            Log("Error Extracting images from selections " & ex.Message)
            Exit Sub
        End Try

        Log("Extracted images from selections.")
        Log("Creating html files...")

        Try
            CreateHtmlFromJpg()
        Catch ex As Exception
            Log("Error Creating html files " & ex.Message)
            Exit Sub
        End Try

        Log("Created html files.")
        Log("Creating Excel file")

        Try
            CreateExcelFromHtml()
        Catch ex As Exception
            Log("Error Creating Excel file " & ex.Message)
            Exit Sub
        End Try

        Log("Created Excel file")

    End Sub

    Private Sub Log(s As String)

        If txtMsg.Text = "" Then
            txtMsg.Text = s
        Else
            txtMsg.AppendText(vbCrLf & s)
        End If

        txtMsg.Visible = True
        txtMsg.ScrollToCaret()
        txtMsg.Refresh()
        Application.DoEvents()
    End Sub

    Sub ExtractSubImages()
        Dim outputfolder As String = GetOutputFolder()
        If outputfolder = "" Then
            Return
        End If

        Dim sTempFolderPath As String = Path.Combine(outputfolder, "Temp")
        If Directory.Exists(sTempFolderPath) = False Then
            Directory.CreateDirectory(sTempFolderPath)
        Else
            For Each sFilePath In Directory.GetFiles(sTempFolderPath, "*.jpg")
                File.Delete(sFilePath)
            Next
        End If

        For Each oEntry As DictionaryEntry In rectangleList
            Dim sPage As String = oEntry.Key
            Dim rectangles As List(Of RectangleF) = rectangleList(sPage)
            Dim iIndex As Integer = 0

            For Each rect As RectangleF In rectangles
                iIndex += 1

                Dim sInputFilePath As String = Path.Combine(outputfolder, sPage & ".jpg")
                Dim sOutputFilePath As String = Path.Combine(sTempFolderPath, sPage & "_" & iIndex & ".jpg")

                ExtractSubImage(sInputFilePath, sOutputFilePath, rect)
            Next
        Next

        ListExtractedImages()
    End Sub

    Sub ListExtractedImages()
        Dim outputfolder As String = GetOutputFolder()
        If outputfolder = "" Then
            Return
        End If

        Dim sTempFolderPath As String = Path.Combine(outputfolder, "Temp")
        If Directory.Exists(sTempFolderPath) = False Then
            Exit Sub
        End If

        lbImages.Items.Clear()

        For Each sFilePath In Directory.GetFiles(sTempFolderPath, "*.jpg")
            lbImages.Items.Add(Path.GetFileNameWithoutExtension(sFilePath))
        Next
    End Sub

    Private Sub lbImages_DoubleClick(sender As Object, e As EventArgs) Handles lbImages.DoubleClick

        If lbImages.SelectedIndex = -1 Then
            Exit Sub
        End If

        Dim outputfolder As String = GetOutputFolder()
        If outputfolder = "" Then
            Return
        End If

        Dim sTempFolderPath As String = Path.Combine(outputfolder, "Temp")
        If Directory.Exists(sTempFolderPath) = False Then
            Exit Sub
        End If

        Dim sFilePath As String = Path.Combine(sTempFolderPath, lbImages.Items(lbImages.SelectedIndex)) & ".jpg"

        Process.Start(New ProcessStartInfo(sFilePath))
    End Sub

    Sub CreateHtmlFromJpg()

        Dim outputfolder As String = GetOutputFolder()
        If outputfolder = "" Then
            Return
        End If

        Dim sTempFolderPath As String = Path.Combine(outputfolder, "Temp")
        If Directory.Exists(sTempFolderPath) = False Then
            Return
        Else
            For Each sFilePath In Directory.GetFiles(sTempFolderPath, "*.htm")
                File.Delete(sFilePath)
            Next
        End If

        Dim sPrompt As String = "Convert image to HTML. Do not provide any comments - just provide HTML."
        Dim sAiService As String = GetComboBoxVal(cbAiService, "gpt-4o")
        Dim oAiCaller As New AICaller()
        oAiCaller.sOpenAiKey = txtOpenAIKey.Text
        oAiCaller.sAnthropicApiKey = txtAnthropicApiKey.Text
        oAiCaller.sModel = sAiService

        For Each sImageFilePath In Directory.GetFiles(sTempFolderPath, "*.jpg")
            Dim inputFileName As String = Path.GetFileNameWithoutExtension(sImageFilePath)
            Dim outputHtmlFile As String = Path.Combine(sTempFolderPath, inputFileName) & ".htm"
            Dim sOutputHtml As String = ""

            Try
                sOutputHtml = oAiCaller.SendImg(sImageFilePath, sPrompt, "high")
            Catch ex As Exception
                MsgBox("SendImg Error: " & ex.Message & " - " & sOutputHtml)
                Exit Sub
            End Try

            sOutputHtml = Replace(sOutputHtml, "```html", "")
            sOutputHtml = Replace(sOutputHtml, "```", "")

            IO.File.WriteAllText(outputHtmlFile, sOutputHtml)
        Next

    End Sub


    Sub CreateExcelFromHtml()
        Dim outputfolder As String = GetOutputFolder()
        If outputfolder = "" Then
            Return
        End If

        Dim sTempFolderPath As String = Path.Combine(outputfolder, "Temp")
        If Directory.Exists(sTempFolderPath) = False Then
            Return
        End If

        Dim oExcel, oWorkBook0
        oExcel = CreateObject("Excel.Application")
        oExcel.Visible = True
        oExcel.DisplayAlerts = False
        oWorkBook0 = oExcel.Workbooks.Add

        Dim oFileList As New List(Of String)

        For Each sFilePath In Directory.GetFiles(sTempFolderPath, "*.htm")
            oFileList.Add(sFilePath)
        Next

        'descending order
        oFileList.Sort(Function(x, y) y.CompareTo(x))

        Dim oWorkBook, oSheet
        For Each sFilePath As String In oFileList
            oWorkBook = oExcel.Workbooks.Open(sFilePath)
            oSheet = oWorkBook.Worksheets(1)
            oSheet.Move(, oWorkBook0.Worksheets(1))
        Next

        Try
            oWorkBook0.Worksheets(1).Delete()
        Catch ex As Exception
            'Delete Sheet1
        End Try

        Dim inputFileName As String = Path.GetFileNameWithoutExtension(txtInputFile.Text)
        Dim sExcelFilePath As String = Path.Combine(outputfolder, inputFileName) & ".xlsx"

        If IO.File.Exists(sExcelFilePath) Then
            TryDeleteFile(sExcelFilePath)

            If IO.File.Exists(sExcelFilePath) Then
                For i As Integer = 1 To 1000
                    sExcelFilePath = Path.Combine(outputfolder, inputFileName & "_" & i) & ".xlsx"

                    If IO.File.Exists(sExcelFilePath) Then
                        TryDeleteFile(sExcelFilePath)
                    End If

                    If IO.File.Exists(sExcelFilePath) = False Then
                        Exit For
                    End If
                Next
            End If

        End If

        oWorkBook0.SaveAs(sExcelFilePath)
        'oWorkBook0.Close()
        'oExcel.Quit()
    End Sub

    Sub TryDeleteFile(sFilePath As String)
        Try
            IO.File.Delete(sFilePath)
        Catch ex As Exception

        End Try
    End Sub

    Private zoomFactor As Double = 1.0
    Private Const zoomStep As Double = 0.1

    Private Sub PictureBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseWheel
        Dim oldWidth As Integer = PictureBox1.Width
        Dim oldHeight As Integer = PictureBox1.Height

        If e.Delta > 0 Then
            ' Zoom In
            zoomFactor += zoomStep
        ElseIf e.Delta < 0 Then
            ' Zoom Out
            zoomFactor -= zoomStep
        End If

        ' Limit the zoom factor to prevent it from becoming too small or too large
        zoomFactor = Math.Max(zoomFactor, 1.0)
        zoomFactor = Math.Min(zoomFactor, 10.0)

        lbZoom.Text = (zoomFactor * 100).ToString() & "%"

        ' Calculate new size of PictureBox
        PictureBox1.Width = CInt(SplitContainer1.Panel2.Width * zoomFactor)
        PictureBox1.Height = CInt(SplitContainer1.Panel2.Height * zoomFactor)

        ' Calculate the mouse position relative to the PictureBox before the zoom
        Dim mouseX As Integer = e.X
        Dim mouseY As Integer = e.Y

        ' Calculate the new position of the PictureBox to keep the zoom centered on the mouse
        PictureBox1.Left -= CInt((mouseX / oldWidth) * (PictureBox1.Width - oldWidth))
        PictureBox1.Top -= CInt((mouseY / oldHeight) * (PictureBox1.Height - oldHeight))
    End Sub

    Private Sub PictureBox1_Resize(sender As Object, e As EventArgs) Handles PictureBox1.Resize
        PictureBox1.Invalidate()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        ' Ensure the PictureBox receives focus to capture the MouseWheel event
        PictureBox1.Focus()
    End Sub

    Private Sub btnResetZoom_Click(sender As Object, e As EventArgs) Handles btnResetZoom.Click
        ResetZoom()
    End Sub

    Sub ResetZoom()
        zoomFactor = 1.0
        lbZoom.Text = (zoomFactor * 100).ToString() & "%"
        PictureBox1.Left = 2
        PictureBox1.Top = 2
        PictureBox1.Width = SplitContainer1.Panel2.Width 'originalWidth
        PictureBox1.Height = SplitContainer1.Panel2.Height 'originalHeight
    End Sub


End Class
