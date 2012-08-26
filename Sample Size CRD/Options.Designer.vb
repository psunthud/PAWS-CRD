<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Options
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
        Me.components = New System.ComponentModel.Container
        Me.Label2 = New System.Windows.Forms.Label
        Me.Okay = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.Directory = New System.Windows.Forms.TextBox
        Me.SetReplication = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(33, 31)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(155, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Number of Replications"
        '
        'Okay
        '
        Me.Okay.Location = New System.Drawing.Point(67, 171)
        Me.Okay.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Okay.Name = "Okay"
        Me.Okay.Size = New System.Drawing.Size(100, 28)
        Me.Okay.TabIndex = 5
        Me.Okay.Text = "OK"
        Me.Okay.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(175, 171)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 28)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Cancel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(33, 108)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(134, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Reset User Settings"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(215, 102)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(100, 28)
        Me.Button4.TabIndex = 8
        Me.Button4.Text = "Reset"
        Me.ToolTip1.SetToolTip(Me.Button4, "Reset settings to default value")
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(33, 69)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(106, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Mplus Directory"
        '
        'Directory
        '
        Me.Directory.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.PAWS_CRD.My.MySettings.Default, "MplusDirectory", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.Directory.Location = New System.Drawing.Point(216, 66)
        Me.Directory.Name = "Directory"
        Me.Directory.Size = New System.Drawing.Size(99, 22)
        Me.Directory.TabIndex = 10
        Me.Directory.Text = Global.PAWS_CRD.My.MySettings.Default.MplusDirectory
        '
        'SetReplication
        '
        Me.SetReplication.DataBindings.Add(New System.Windows.Forms.Binding("Text", Global.PAWS_CRD.My.MySettings.Default, "StorageNRep", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.SetReplication.Location = New System.Drawing.Point(216, 27)
        Me.SetReplication.Margin = New System.Windows.Forms.Padding(4)
        Me.SetReplication.Name = "SetReplication"
        Me.SetReplication.Size = New System.Drawing.Size(99, 22)
        Me.SetReplication.TabIndex = 4
        Me.SetReplication.Text = Global.PAWS_CRD.My.MySettings.Default.StorageNRep
        Me.SetReplication.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ToolTip1.SetToolTip(Me.SetReplication, "Number of replications used in a priori Monte Carlo Simulation")
        '
        'Options
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(340, 212)
        Me.Controls.Add(Me.Directory)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Okay)
        Me.Controls.Add(Me.SetReplication)
        Me.Controls.Add(Me.Label2)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "Options"
        Me.Text = "Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents SetReplication As System.Windows.Forms.TextBox
    Friend WithEvents Okay As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Directory As System.Windows.Forms.TextBox
End Class
