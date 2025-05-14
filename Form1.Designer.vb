<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.lbImages = New System.Windows.Forms.ListBox()
        Me.txtMsg = New System.Windows.Forms.TextBox()
        Me.btnResetZoom = New System.Windows.Forms.Button()
        Me.lbZoom = New System.Windows.Forms.Label()
        Me.btnCreateExcel = New System.Windows.Forms.Button()
        Me.btnLoadPdf = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.llGS = New System.Windows.Forms.LinkLabel()
        Me.llOpenAI = New System.Windows.Forms.LinkLabel()
        Me.llAnthropic = New System.Windows.Forms.LinkLabel()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnInputFile = New System.Windows.Forms.Button()
        Me.btnGhostscriptPath = New System.Windows.Forms.Button()
        Me.btnOpenAIKeyShow = New System.Windows.Forms.Button()
        Me.btnApiKeyShow = New System.Windows.Forms.Button()
        Me.txtOpenAIKey = New System.Windows.Forms.TextBox()
        Me.txtGhostscriptPath = New System.Windows.Forms.TextBox()
        Me.txtDotsPerInch = New System.Windows.Forms.TextBox()
        Me.txtInputFile = New System.Windows.Forms.TextBox()
        Me.txtAnthropicApiKey = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbAiService = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label7 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lbImages)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtMsg)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnResetZoom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lbZoom)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnCreateExcel)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnLoadPdf)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataGridView1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClearAll)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llGS)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llOpenAI)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llAnthropic)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClear)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnInputFile)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnGhostscriptPath)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnOpenAIKeyShow)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnApiKeyShow)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtOpenAIKey)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtGhostscriptPath)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtDotsPerInch)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtInputFile)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtAnthropicApiKey)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cbAiService)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel2.Controls.Add(Me.PictureBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1693, 925)
        Me.SplitContainer1.SplitterDistance = 613
        Me.SplitContainer1.TabIndex = 0
        '
        'lbImages
        '
        Me.lbImages.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbImages.FormattingEnabled = True
        Me.lbImages.ItemHeight = 20
        Me.lbImages.Location = New System.Drawing.Point(527, 589)
        Me.lbImages.Name = "lbImages"
        Me.lbImages.Size = New System.Drawing.Size(74, 124)
        Me.lbImages.TabIndex = 26
        '
        'txtMsg
        '
        Me.txtMsg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMsg.Location = New System.Drawing.Point(16, 731)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.Size = New System.Drawing.Size(585, 116)
        Me.txtMsg.TabIndex = 25
        '
        'btnResetZoom
        '
        Me.btnResetZoom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnResetZoom.Location = New System.Drawing.Point(527, 517)
        Me.btnResetZoom.Name = "btnResetZoom"
        Me.btnResetZoom.Size = New System.Drawing.Size(74, 66)
        Me.btnResetZoom.TabIndex = 24
        Me.btnResetZoom.Text = "Zoom 100%"
        Me.btnResetZoom.UseVisualStyleBackColor = True
        '
        'lbZoom
        '
        Me.lbZoom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbZoom.AutoSize = True
        Me.lbZoom.Location = New System.Drawing.Point(537, 484)
        Me.lbZoom.Name = "lbZoom"
        Me.lbZoom.Size = New System.Drawing.Size(50, 20)
        Me.lbZoom.TabIndex = 1
        Me.lbZoom.Text = "100%"
        '
        'btnCreateExcel
        '
        Me.btnCreateExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCreateExcel.Location = New System.Drawing.Point(306, 853)
        Me.btnCreateExcel.Name = "btnCreateExcel"
        Me.btnCreateExcel.Size = New System.Drawing.Size(139, 48)
        Me.btnCreateExcel.TabIndex = 23
        Me.btnCreateExcel.Text = "Export to Excel"
        Me.btnCreateExcel.UseVisualStyleBackColor = True
        '
        'btnLoadPdf
        '
        Me.btnLoadPdf.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLoadPdf.Location = New System.Drawing.Point(153, 853)
        Me.btnLoadPdf.Name = "btnLoadPdf"
        Me.btnLoadPdf.Size = New System.Drawing.Size(123, 48)
        Me.btnLoadPdf.TabIndex = 22
        Me.btnLoadPdf.Text = "Load PDF"
        Me.btnLoadPdf.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(16, 339)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.RowTemplate.Height = 28
        Me.DataGridView1.Size = New System.Drawing.Size(505, 386)
        Me.DataGridView1.TabIndex = 21
        '
        'btnClearAll
        '
        Me.btnClearAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearAll.Location = New System.Drawing.Point(527, 407)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(74, 62)
        Me.btnClearAll.TabIndex = 20
        Me.btnClearAll.Text = "Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'llGS
        '
        Me.llGS.AutoSize = True
        Me.llGS.Location = New System.Drawing.Point(149, 180)
        Me.llGS.Name = "llGS"
        Me.llGS.Size = New System.Drawing.Size(18, 20)
        Me.llGS.TabIndex = 19
        Me.llGS.TabStop = True
        Me.llGS.Text = "?"
        '
        'llOpenAI
        '
        Me.llOpenAI.AutoSize = True
        Me.llOpenAI.Location = New System.Drawing.Point(149, 125)
        Me.llOpenAI.Name = "llOpenAI"
        Me.llOpenAI.Size = New System.Drawing.Size(18, 20)
        Me.llOpenAI.TabIndex = 18
        Me.llOpenAI.TabStop = True
        Me.llOpenAI.Text = "?"
        '
        'llAnthropic
        '
        Me.llAnthropic.AutoSize = True
        Me.llAnthropic.Location = New System.Drawing.Point(149, 68)
        Me.llAnthropic.Name = "llAnthropic"
        Me.llAnthropic.Size = New System.Drawing.Size(18, 20)
        Me.llAnthropic.TabIndex = 17
        Me.llAnthropic.TabStop = True
        Me.llAnthropic.Text = "?"
        '
        'btnClear
        '
        Me.btnClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClear.Location = New System.Drawing.Point(527, 339)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(74, 62)
        Me.btnClear.TabIndex = 16
        Me.btnClear.Text = "Clear Page"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnInputFile
        '
        Me.btnInputFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInputFile.Location = New System.Drawing.Point(527, 277)
        Me.btnInputFile.Name = "btnInputFile"
        Me.btnInputFile.Size = New System.Drawing.Size(74, 40)
        Me.btnInputFile.TabIndex = 15
        Me.btnInputFile.Text = "..."
        Me.btnInputFile.UseVisualStyleBackColor = True
        '
        'btnGhostscriptPath
        '
        Me.btnGhostscriptPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGhostscriptPath.Location = New System.Drawing.Point(527, 167)
        Me.btnGhostscriptPath.Name = "btnGhostscriptPath"
        Me.btnGhostscriptPath.Size = New System.Drawing.Size(74, 33)
        Me.btnGhostscriptPath.TabIndex = 14
        Me.btnGhostscriptPath.Text = "..."
        Me.btnGhostscriptPath.UseVisualStyleBackColor = True
        '
        'btnOpenAIKeyShow
        '
        Me.btnOpenAIKeyShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOpenAIKeyShow.Location = New System.Drawing.Point(527, 115)
        Me.btnOpenAIKeyShow.Name = "btnOpenAIKeyShow"
        Me.btnOpenAIKeyShow.Size = New System.Drawing.Size(74, 34)
        Me.btnOpenAIKeyShow.TabIndex = 13
        Me.btnOpenAIKeyShow.Text = "*"
        Me.btnOpenAIKeyShow.UseVisualStyleBackColor = True
        '
        'btnApiKeyShow
        '
        Me.btnApiKeyShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApiKeyShow.Location = New System.Drawing.Point(527, 57)
        Me.btnApiKeyShow.Name = "btnApiKeyShow"
        Me.btnApiKeyShow.Size = New System.Drawing.Size(74, 36)
        Me.btnApiKeyShow.TabIndex = 12
        Me.btnApiKeyShow.Text = "*"
        Me.btnApiKeyShow.UseVisualStyleBackColor = True
        '
        'txtOpenAIKey
        '
        Me.txtOpenAIKey.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOpenAIKey.Location = New System.Drawing.Point(173, 119)
        Me.txtOpenAIKey.Name = "txtOpenAIKey"
        Me.txtOpenAIKey.Size = New System.Drawing.Size(348, 26)
        Me.txtOpenAIKey.TabIndex = 11
        '
        'txtGhostscriptPath
        '
        Me.txtGhostscriptPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGhostscriptPath.Location = New System.Drawing.Point(173, 174)
        Me.txtGhostscriptPath.Name = "txtGhostscriptPath"
        Me.txtGhostscriptPath.Size = New System.Drawing.Size(348, 26)
        Me.txtGhostscriptPath.TabIndex = 10
        Me.txtGhostscriptPath.Text = "C:\Program Files\gs\gs10.01.1\bin\gswin64.exe"
        '
        'txtDotsPerInch
        '
        Me.txtDotsPerInch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDotsPerInch.Location = New System.Drawing.Point(173, 231)
        Me.txtDotsPerInch.Name = "txtDotsPerInch"
        Me.txtDotsPerInch.Size = New System.Drawing.Size(348, 26)
        Me.txtDotsPerInch.TabIndex = 9
        Me.txtDotsPerInch.Text = "1200"
        '
        'txtInputFile
        '
        Me.txtInputFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInputFile.Location = New System.Drawing.Point(173, 284)
        Me.txtInputFile.Name = "txtInputFile"
        Me.txtInputFile.Size = New System.Drawing.Size(348, 26)
        Me.txtInputFile.TabIndex = 8
        '
        'txtAnthropicApiKey
        '
        Me.txtAnthropicApiKey.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAnthropicApiKey.Location = New System.Drawing.Point(173, 62)
        Me.txtAnthropicApiKey.Name = "txtAnthropicApiKey"
        Me.txtAnthropicApiKey.Size = New System.Drawing.Size(348, 26)
        Me.txtAnthropicApiKey.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 290)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(107, 20)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "PDF File Path"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(135, 20)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Anthropic API key"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 125)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(122, 20)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "OpenAI API key"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 180)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(131, 20)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Ghost script path"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 237)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 20)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Resolution (DPI)"
        '
        'cbAiService
        '
        Me.cbAiService.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbAiService.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbAiService.FormattingEnabled = True
        Me.cbAiService.Items.AddRange(New Object() {"gpt-4o", "gpt-4o-mini", "claude-3-5-sonnet-20240620"})
        Me.cbAiService.Location = New System.Drawing.Point(173, 16)
        Me.cbAiService.Name = "cbAiService"
        Me.cbAiService.Size = New System.Drawing.Size(348, 28)
        Me.cbAiService.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "AI Service"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(1070, 884)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 896)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(786, 20)
        Me.Label7.TabIndex = 27
        Me.Label7.Text = "Select a page and them select a table to export (click drag and release).  Use mo" &
    "use wheel to zoom in and out."
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1693, 925)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "Form1"
        Me.Text = "PDF to Excel"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents btnCreateExcel As Button
    Friend WithEvents btnLoadPdf As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents btnClearAll As Button
    Friend WithEvents llGS As LinkLabel
    Friend WithEvents llOpenAI As LinkLabel
    Friend WithEvents llAnthropic As LinkLabel
    Friend WithEvents btnClear As Button
    Friend WithEvents btnInputFile As Button
    Friend WithEvents btnGhostscriptPath As Button
    Friend WithEvents btnOpenAIKeyShow As Button
    Friend WithEvents btnApiKeyShow As Button
    Friend WithEvents txtOpenAIKey As TextBox
    Friend WithEvents txtGhostscriptPath As TextBox
    Friend WithEvents txtDotsPerInch As TextBox
    Friend WithEvents txtInputFile As TextBox
    Friend WithEvents txtAnthropicApiKey As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents cbAiService As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents lbZoom As Label
    Friend WithEvents btnResetZoom As Button
    Friend WithEvents txtMsg As TextBox
    Friend WithEvents lbImages As ListBox
    Friend WithEvents Label7 As Label
End Class
