<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PowerBlueServerApp
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PowerBlueServerApp))
        Me.BrowsePptButton = New System.Windows.Forms.Button()
        Me.StartServerButton = New System.Windows.Forms.Button()
        Me.PptFileSelectionDialog = New System.Windows.Forms.OpenFileDialog()
        Me.PowerBlueTitle = New System.Windows.Forms.Label()
        Me.CopyRightLabel = New System.Windows.Forms.Label()
        Me.ClearLogsButton = New System.Windows.Forms.Button()
        Me.VersionningLabel = New System.Windows.Forms.Label()
        Me.LogsLabel = New System.Windows.Forms.Label()
        Me.PowerBlueLogTextBox = New System.Windows.Forms.TextBox()
        Me.PowerBlueServerBackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.StopServerButton = New System.Windows.Forms.Button()
        Me.Step1 = New System.Windows.Forms.Label()
        Me.Step2 = New System.Windows.Forms.Label()
        Me.Step3 = New System.Windows.Forms.Label()
        Me.ServerHelp = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'BrowsePptButton
        '
        Me.BrowsePptButton.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.BrowsePptButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BrowsePptButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BrowsePptButton.Location = New System.Drawing.Point(12, 136)
        Me.BrowsePptButton.Name = "BrowsePptButton"
        Me.BrowsePptButton.Size = New System.Drawing.Size(139, 47)
        Me.BrowsePptButton.TabIndex = 0
        Me.BrowsePptButton.Text = "Browse PPT"
        Me.BrowsePptButton.UseVisualStyleBackColor = True
        '
        'StartServerButton
        '
        Me.StartServerButton.Enabled = False
        Me.StartServerButton.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.StartServerButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StartServerButton.Location = New System.Drawing.Point(157, 136)
        Me.StartServerButton.Name = "StartServerButton"
        Me.StartServerButton.Size = New System.Drawing.Size(136, 47)
        Me.StartServerButton.TabIndex = 1
        Me.StartServerButton.Text = "Start Server"
        Me.StartServerButton.UseVisualStyleBackColor = True
        '
        'PptFileSelectionDialog
        '
        Me.PptFileSelectionDialog.FileName = "SelectedPptFile"
        '
        'PowerBlueTitle
        '
        Me.PowerBlueTitle.AutoSize = True
        Me.PowerBlueTitle.Font = New System.Drawing.Font("Jokerman", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PowerBlueTitle.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.PowerBlueTitle.Location = New System.Drawing.Point(93, 44)
        Me.PowerBlueTitle.Name = "PowerBlueTitle"
        Me.PowerBlueTitle.Size = New System.Drawing.Size(262, 58)
        Me.PowerBlueTitle.TabIndex = 4
        Me.PowerBlueTitle.Text = "Power Blue"
        '
        'CopyRightLabel
        '
        Me.CopyRightLabel.AutoSize = True
        Me.CopyRightLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CopyRightLabel.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.CopyRightLabel.Location = New System.Drawing.Point(72, 377)
        Me.CopyRightLabel.Name = "CopyRightLabel"
        Me.CopyRightLabel.Size = New System.Drawing.Size(292, 13)
        Me.CopyRightLabel.TabIndex = 6
        Me.CopyRightLabel.Text = "Developer: Sreedhar Reddy V    Mail: Srib4ufrnd@gmail.com"
        Me.CopyRightLabel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ClearLogsButton
        '
        Me.ClearLogsButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClearLogsButton.Location = New System.Drawing.Point(185, 330)
        Me.ClearLogsButton.Name = "ClearLogsButton"
        Me.ClearLogsButton.Size = New System.Drawing.Size(46, 23)
        Me.ClearLogsButton.TabIndex = 7
        Me.ClearLogsButton.Text = "Clear"
        Me.ClearLogsButton.UseVisualStyleBackColor = True
        '
        'VersionningLabel
        '
        Me.VersionningLabel.AutoSize = True
        Me.VersionningLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VersionningLabel.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.VersionningLabel.Location = New System.Drawing.Point(349, 77)
        Me.VersionningLabel.Name = "VersionningLabel"
        Me.VersionningLabel.Size = New System.Drawing.Size(37, 13)
        Me.VersionningLabel.TabIndex = 8
        Me.VersionningLabel.Text = "V 1.0"
        '
        'LogsLabel
        '
        Me.LogsLabel.AutoSize = True
        Me.LogsLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LogsLabel.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.LogsLabel.Location = New System.Drawing.Point(192, 204)
        Me.LogsLabel.Name = "LogsLabel"
        Me.LogsLabel.Size = New System.Drawing.Size(34, 13)
        Me.LogsLabel.TabIndex = 9
        Me.LogsLabel.Text = "Logs"
        '
        'PowerBlueLogTextBox
        '
        Me.PowerBlueLogTextBox.CausesValidation = False
        Me.PowerBlueLogTextBox.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.PowerBlueLogTextBox.Location = New System.Drawing.Point(12, 220)
        Me.PowerBlueLogTextBox.MaxLength = 999999999
        Me.PowerBlueLogTextBox.Multiline = True
        Me.PowerBlueLogTextBox.Name = "PowerBlueLogTextBox"
        Me.PowerBlueLogTextBox.ReadOnly = True
        Me.PowerBlueLogTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.PowerBlueLogTextBox.Size = New System.Drawing.Size(417, 104)
        Me.PowerBlueLogTextBox.TabIndex = 5
        '
        'PowerBlueServerBackgroundWorker
        '
        Me.PowerBlueServerBackgroundWorker.WorkerSupportsCancellation = True
        '
        'StopServerButton
        '
        Me.StopServerButton.Enabled = False
        Me.StopServerButton.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.StopServerButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StopServerButton.Location = New System.Drawing.Point(299, 136)
        Me.StopServerButton.Name = "StopServerButton"
        Me.StopServerButton.Size = New System.Drawing.Size(130, 47)
        Me.StopServerButton.TabIndex = 10
        Me.StopServerButton.Text = "Stop Server"
        Me.StopServerButton.UseVisualStyleBackColor = True
        '
        'Step1
        '
        Me.Step1.AutoSize = True
        Me.Step1.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Step1.Location = New System.Drawing.Point(63, 120)
        Me.Step1.Name = "Step1"
        Me.Step1.Size = New System.Drawing.Size(41, 13)
        Me.Step1.TabIndex = 11
        Me.Step1.Text = "Step: 1"
        '
        'Step2
        '
        Me.Step2.AutoSize = True
        Me.Step2.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Step2.Location = New System.Drawing.Point(204, 120)
        Me.Step2.Name = "Step2"
        Me.Step2.Size = New System.Drawing.Size(41, 13)
        Me.Step2.TabIndex = 12
        Me.Step2.Text = "Step: 2"
        '
        'Step3
        '
        Me.Step3.AutoSize = True
        Me.Step3.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Step3.Location = New System.Drawing.Point(349, 120)
        Me.Step3.Name = "Step3"
        Me.Step3.Size = New System.Drawing.Size(41, 13)
        Me.Step3.TabIndex = 13
        Me.Step3.Text = "Step: 3"
        '
        'ServerHelp
        '
        Me.ServerHelp.AutoSize = True
        Me.ServerHelp.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ServerHelp.Location = New System.Drawing.Point(396, 9)
        Me.ServerHelp.Name = "ServerHelp"
        Me.ServerHelp.Size = New System.Drawing.Size(38, 13)
        Me.ServerHelp.TabIndex = 15
        Me.ServerHelp.Text = "Help ?"
        '
        'PowerBlueServerApp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(80, Byte), Integer), CType(CType(179, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(446, 399)
        Me.Controls.Add(Me.ServerHelp)
        Me.Controls.Add(Me.Step3)
        Me.Controls.Add(Me.Step2)
        Me.Controls.Add(Me.Step1)
        Me.Controls.Add(Me.StopServerButton)
        Me.Controls.Add(Me.LogsLabel)
        Me.Controls.Add(Me.VersionningLabel)
        Me.Controls.Add(Me.ClearLogsButton)
        Me.Controls.Add(Me.CopyRightLabel)
        Me.Controls.Add(Me.PowerBlueLogTextBox)
        Me.Controls.Add(Me.PowerBlueTitle)
        Me.Controls.Add(Me.StartServerButton)
        Me.Controls.Add(Me.BrowsePptButton)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "PowerBlueServerApp"
        Me.Text = "PowerBlue Server"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BrowsePptButton As System.Windows.Forms.Button
    Friend WithEvents StartServerButton As System.Windows.Forms.Button
    Friend WithEvents PptFileSelectionDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents PowerBlueTitle As System.Windows.Forms.Label
    Friend WithEvents CopyRightLabel As System.Windows.Forms.Label
    Friend WithEvents ClearLogsButton As System.Windows.Forms.Button
    Friend WithEvents VersionningLabel As System.Windows.Forms.Label
    Friend WithEvents LogsLabel As System.Windows.Forms.Label
    Friend WithEvents PowerBlueLogTextBox As System.Windows.Forms.TextBox
    Friend WithEvents PowerBlueServerBackgroundWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents StopServerButton As System.Windows.Forms.Button
    Friend WithEvents Step1 As System.Windows.Forms.Label
    Friend WithEvents Step2 As System.Windows.Forms.Label
    Friend WithEvents Step3 As System.Windows.Forms.Label
    Friend WithEvents ServerHelp As System.Windows.Forms.Label

End Class
