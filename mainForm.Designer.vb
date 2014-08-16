<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mainForm
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
        Me.fileLabel = New System.Windows.Forms.Label()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.statusTxt = New System.Windows.Forms.Label()
        Me.fileTxt = New System.Windows.Forms.Label()
        Me.rcBackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.progressBar = New System.Windows.Forms.ProgressBar()
        Me.analyzeButton = New System.Windows.Forms.Button()
        Me.aboutButton = New System.Windows.Forms.Button()
        Me.exitButton = New System.Windows.Forms.Button()
        Me.cancelWorkButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'fileLabel
        '
        Me.fileLabel.AutoSize = True
        Me.fileLabel.Location = New System.Drawing.Point(13, 13)
        Me.fileLabel.Name = "fileLabel"
        Me.fileLabel.Size = New System.Drawing.Size(23, 13)
        Me.fileLabel.TabIndex = 0
        Me.fileLabel.Text = "File"
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Location = New System.Drawing.Point(12, 38)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(37, 13)
        Me.statusLabel.TabIndex = 1
        Me.statusLabel.Text = "Status"
        '
        'statusTxt
        '
        Me.statusTxt.AutoSize = True
        Me.statusTxt.Location = New System.Drawing.Point(52, 38)
        Me.statusTxt.Name = "statusTxt"
        Me.statusTxt.Size = New System.Drawing.Size(47, 13)
        Me.statusTxt.TabIndex = 3
        Me.statusTxt.Text = "<status>"
        '
        'fileTxt
        '
        Me.fileTxt.AutoSize = True
        Me.fileTxt.Location = New System.Drawing.Point(53, 13)
        Me.fileTxt.Name = "fileTxt"
        Me.fileTxt.Size = New System.Drawing.Size(32, 13)
        Me.fileTxt.TabIndex = 2
        Me.fileTxt.Text = "<file>"
        '
        'rcBackgroundWorker
        '
        Me.rcBackgroundWorker.WorkerReportsProgress = True
        Me.rcBackgroundWorker.WorkerSupportsCancellation = True
        '
        'progressBar
        '
        Me.progressBar.Location = New System.Drawing.Point(13, 55)
        Me.progressBar.Name = "progressBar"
        Me.progressBar.Size = New System.Drawing.Size(318, 11)
        Me.progressBar.TabIndex = 4
        '
        'analyzeButton
        '
        Me.analyzeButton.Location = New System.Drawing.Point(13, 72)
        Me.analyzeButton.Name = "analyzeButton"
        Me.analyzeButton.Size = New System.Drawing.Size(75, 23)
        Me.analyzeButton.TabIndex = 5
        Me.analyzeButton.Text = "Analyze"
        Me.analyzeButton.UseVisualStyleBackColor = True
        '
        'aboutButton
        '
        Me.aboutButton.Location = New System.Drawing.Point(175, 72)
        Me.aboutButton.Name = "aboutButton"
        Me.aboutButton.Size = New System.Drawing.Size(75, 23)
        Me.aboutButton.TabIndex = 6
        Me.aboutButton.Text = "About"
        Me.aboutButton.UseVisualStyleBackColor = True
        '
        'exitButton
        '
        Me.exitButton.Location = New System.Drawing.Point(256, 72)
        Me.exitButton.Name = "exitButton"
        Me.exitButton.Size = New System.Drawing.Size(75, 23)
        Me.exitButton.TabIndex = 7
        Me.exitButton.Text = "Exit"
        Me.exitButton.UseVisualStyleBackColor = True
        '
        'cancelWorkButton
        '
        Me.cancelWorkButton.Location = New System.Drawing.Point(94, 72)
        Me.cancelWorkButton.Name = "cancelWorkButton"
        Me.cancelWorkButton.Size = New System.Drawing.Size(75, 23)
        Me.cancelWorkButton.TabIndex = 8
        Me.cancelWorkButton.Text = "Cancel"
        Me.cancelWorkButton.UseVisualStyleBackColor = True
        '
        'mainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(343, 109)
        Me.Controls.Add(Me.cancelWorkButton)
        Me.Controls.Add(Me.exitButton)
        Me.Controls.Add(Me.aboutButton)
        Me.Controls.Add(Me.analyzeButton)
        Me.Controls.Add(Me.progressBar)
        Me.Controls.Add(Me.statusTxt)
        Me.Controls.Add(Me.fileTxt)
        Me.Controls.Add(Me.statusLabel)
        Me.Controls.Add(Me.fileLabel)
        Me.Name = "mainForm"
        Me.Text = "RangeChecker"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents fileLabel As System.Windows.Forms.Label
    Friend WithEvents statusLabel As System.Windows.Forms.Label
    Friend WithEvents statusTxt As System.Windows.Forms.Label
    Friend WithEvents fileTxt As System.Windows.Forms.Label
    Friend WithEvents rcBackgroundWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents progressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents analyzeButton As System.Windows.Forms.Button
    Friend WithEvents aboutButton As System.Windows.Forms.Button
    Friend WithEvents exitButton As System.Windows.Forms.Button
    Friend WithEvents cancelWorkButton As System.Windows.Forms.Button

End Class
