<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.BGWL7P = New System.ComponentModel.BackgroundWorker
        Me.BGWG4P = New System.ComponentModel.BackgroundWorker
        Me.BGWL6P = New System.ComponentModel.BackgroundWorker
        Me.BGWGBP = New System.ComponentModel.BackgroundWorker
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.SuspendLayout()
        '
        'BGWL7P
        '
        Me.BGWL7P.WorkerReportsProgress = True
        Me.BGWL7P.WorkerSupportsCancellation = True
        '
        'BGWG4P
        '
        Me.BGWG4P.WorkerReportsProgress = True
        Me.BGWG4P.WorkerSupportsCancellation = True
        '
        'BGWL6P
        '
        Me.BGWL6P.WorkerReportsProgress = True
        Me.BGWL6P.WorkerSupportsCancellation = True
        '
        'BGWGBP
        '
        Me.BGWGBP.WorkerReportsProgress = True
        Me.BGWGBP.WorkerSupportsCancellation = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(1, 2)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(388, 65)
        Me.ProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee
        Me.ProgressBar1.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 72)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Name = "Form1"
        Me.Text = "BR NCM Changes [09/04/2013]"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BGWL7P As System.ComponentModel.BackgroundWorker
    Friend WithEvents BGWG4P As System.ComponentModel.BackgroundWorker
    Friend WithEvents BGWL6P As System.ComponentModel.BackgroundWorker
    Friend WithEvents BGWGBP As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar

End Class
