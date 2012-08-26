<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Criteria
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Criteria))
        Me.InputCriteria = New System.Windows.Forms.GroupBox
        Me.Option3 = New System.Windows.Forms.RadioButton
        Me.Option2 = New System.Windows.Forms.RadioButton
        Me.Option1 = New System.Windows.Forms.RadioButton
        Me.PCheckBox = New System.Windows.Forms.CheckBox
        Me.RPowerWidth2 = New System.Windows.Forms.RadioButton
        Me.RPowerWidth1 = New System.Windows.Forms.RadioButton
        Me.PowerGroupBox = New System.Windows.Forms.GroupBox
        Me.DegreeOfCertaintyLabel2 = New System.Windows.Forms.Label
        Me.IsDegreeOfCertainty = New System.Windows.Forms.ComboBox
        Me.DegreeOfCertainty = New System.Windows.Forms.TextBox
        Me.DegreeOfCertaintyLabel = New System.Windows.Forms.Label
        Me.DesiredWidth = New System.Windows.Forms.TextBox
        Me.Power = New System.Windows.Forms.TextBox
        Me.CILevel = New System.Windows.Forms.ComboBox
        Me.CIText = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.NIndiv = New System.Windows.Forms.TextBox
        Me.NGroups = New System.Windows.Forms.TextBox
        Me.PT = New System.Windows.Forms.TextBox
        Me.SSLabel = New System.Windows.Forms.Label
        Me.NIndivCheckBox = New System.Windows.Forms.CheckBox
        Me.NGCheckBox = New System.Windows.Forms.CheckBox
        Me.CostInfo = New System.Windows.Forms.GroupBox
        Me.TotalCost = New System.Windows.Forms.TextBox
        Me.CIndivCost = New System.Windows.Forms.TextBox
        Me.CGroupCost = New System.Windows.Forms.TextBox
        Me.TIndivCost = New System.Windows.Forms.TextBox
        Me.TGroupCost = New System.Windows.Forms.TextBox
        Me.TotalCostText = New System.Windows.Forms.Label
        Me.CostIndiv = New System.Windows.Forms.Label
        Me.CostGroup = New System.Windows.Forms.Label
        Me.CostCtrl = New System.Windows.Forms.Label
        Me.CostTreatment = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.InputCriteria.SuspendLayout()
        Me.PowerGroupBox.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.CostInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'InputCriteria
        '
        Me.InputCriteria.Controls.Add(Me.Option3)
        Me.InputCriteria.Controls.Add(Me.Option2)
        Me.InputCriteria.Controls.Add(Me.Option1)
        Me.InputCriteria.Location = New System.Drawing.Point(12, 12)
        Me.InputCriteria.Name = "InputCriteria"
        Me.InputCriteria.Size = New System.Drawing.Size(414, 108)
        Me.InputCriteria.TabIndex = 32
        Me.InputCriteria.TabStop = False
        Me.InputCriteria.Text = "Input Criterion"
        '
        'Option3
        '
        Me.Option3.AutoSize = True
        Me.Option3.Location = New System.Drawing.Point(30, 65)
        Me.Option3.Name = "Option3"
        Me.Option3.Size = New System.Drawing.Size(327, 17)
        Me.Option3.TabIndex = 2
        Me.Option3.Text = "Maximize the Power or Minimize the Width with a Limited Budget"
        Me.Option3.UseVisualStyleBackColor = True
        '
        'Option2
        '
        Me.Option2.AutoSize = True
        Me.Option2.Location = New System.Drawing.Point(30, 42)
        Me.Option2.Name = "Option2"
        Me.Option2.Size = New System.Drawing.Size(261, 17)
        Me.Option2.TabIndex = 1
        Me.Option2.Text = "Minimize the Cost with a Specified Power or Width"
        Me.Option2.UseVisualStyleBackColor = True
        '
        'Option1
        '
        Me.Option1.AutoSize = True
        Me.Option1.Location = New System.Drawing.Point(30, 19)
        Me.Option1.Name = "Option1"
        Me.Option1.Size = New System.Drawing.Size(369, 17)
        Me.Option1.TabIndex = 0
        Me.Option1.Text = "Minimize the Total Number of Individuals with a Specified Power or Width"
        Me.Option1.UseVisualStyleBackColor = True
        '
        'PCheckBox
        '
        Me.PCheckBox.AutoSize = True
        Me.PCheckBox.Enabled = False
        Me.PCheckBox.Location = New System.Drawing.Point(30, 56)
        Me.PCheckBox.Name = "PCheckBox"
        Me.PCheckBox.Size = New System.Drawing.Size(177, 17)
        Me.PCheckBox.TabIndex = 3
        Me.PCheckBox.Text = "Proportion of Treatment Clusters"
        Me.PCheckBox.UseVisualStyleBackColor = True
        '
        'RPowerWidth2
        '
        Me.RPowerWidth2.AutoSize = True
        Me.RPowerWidth2.Enabled = False
        Me.RPowerWidth2.Location = New System.Drawing.Point(30, 46)
        Me.RPowerWidth2.Name = "RPowerWidth2"
        Me.RPowerWidth2.Size = New System.Drawing.Size(144, 17)
        Me.RPowerWidth2.TabIndex = 16
        Me.RPowerWidth2.Text = "Width of CI of Effect Size"
        Me.RPowerWidth2.UseVisualStyleBackColor = True
        '
        'RPowerWidth1
        '
        Me.RPowerWidth1.AutoSize = True
        Me.RPowerWidth1.Enabled = False
        Me.RPowerWidth1.Location = New System.Drawing.Point(30, 20)
        Me.RPowerWidth1.Name = "RPowerWidth1"
        Me.RPowerWidth1.Size = New System.Drawing.Size(55, 17)
        Me.RPowerWidth1.TabIndex = 14
        Me.RPowerWidth1.Text = "Power"
        Me.RPowerWidth1.UseVisualStyleBackColor = True
        '
        'PowerGroupBox
        '
        Me.PowerGroupBox.Controls.Add(Me.DegreeOfCertaintyLabel2)
        Me.PowerGroupBox.Controls.Add(Me.IsDegreeOfCertainty)
        Me.PowerGroupBox.Controls.Add(Me.DegreeOfCertainty)
        Me.PowerGroupBox.Controls.Add(Me.DegreeOfCertaintyLabel)
        Me.PowerGroupBox.Controls.Add(Me.DesiredWidth)
        Me.PowerGroupBox.Controls.Add(Me.Power)
        Me.PowerGroupBox.Controls.Add(Me.CILevel)
        Me.PowerGroupBox.Controls.Add(Me.CIText)
        Me.PowerGroupBox.Controls.Add(Me.RPowerWidth2)
        Me.PowerGroupBox.Controls.Add(Me.RPowerWidth1)
        Me.PowerGroupBox.Enabled = False
        Me.PowerGroupBox.Location = New System.Drawing.Point(13, 426)
        Me.PowerGroupBox.Name = "PowerGroupBox"
        Me.PowerGroupBox.Size = New System.Drawing.Size(413, 130)
        Me.PowerGroupBox.TabIndex = 33
        Me.PowerGroupBox.TabStop = False
        Me.PowerGroupBox.Text = "Desired Power or Width of Confidence Interval of Effect Size"
        '
        'DegreeOfCertaintyLabel2
        '
        Me.DegreeOfCertaintyLabel2.AutoSize = True
        Me.DegreeOfCertaintyLabel2.Enabled = False
        Me.DegreeOfCertaintyLabel2.Location = New System.Drawing.Point(254, 101)
        Me.DegreeOfCertaintyLabel2.Name = "DegreeOfCertaintyLabel2"
        Me.DegreeOfCertaintyLabel2.Size = New System.Drawing.Size(63, 13)
        Me.DegreeOfCertaintyLabel2.TabIndex = 22
        Me.DegreeOfCertaintyLabel2.Text = "with level of"
        '
        'IsDegreeOfCertainty
        '
        Me.IsDegreeOfCertainty.Enabled = False
        Me.IsDegreeOfCertainty.FormattingEnabled = True
        Me.IsDegreeOfCertainty.Location = New System.Drawing.Point(154, 98)
        Me.IsDegreeOfCertainty.Name = "IsDegreeOfCertainty"
        Me.IsDegreeOfCertainty.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.IsDegreeOfCertainty.Size = New System.Drawing.Size(94, 21)
        Me.IsDegreeOfCertainty.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.IsDegreeOfCertainty, resources.GetString("IsDegreeOfCertainty.ToolTip"))
        '
        'DegreeOfCertainty
        '
        Me.DegreeOfCertainty.Enabled = False
        Me.DegreeOfCertainty.Location = New System.Drawing.Point(322, 98)
        Me.DegreeOfCertainty.Name = "DegreeOfCertainty"
        Me.DegreeOfCertainty.Size = New System.Drawing.Size(85, 20)
        Me.DegreeOfCertainty.TabIndex = 20
        '
        'DegreeOfCertaintyLabel
        '
        Me.DegreeOfCertaintyLabel.AutoSize = True
        Me.DegreeOfCertaintyLabel.Enabled = False
        Me.DegreeOfCertaintyLabel.Location = New System.Drawing.Point(50, 101)
        Me.DegreeOfCertaintyLabel.Name = "DegreeOfCertaintyLabel"
        Me.DegreeOfCertaintyLabel.Size = New System.Drawing.Size(98, 13)
        Me.DegreeOfCertaintyLabel.TabIndex = 19
        Me.DegreeOfCertaintyLabel.Text = "Degree of Certainty"
        '
        'DesiredWidth
        '
        Me.DesiredWidth.Enabled = False
        Me.DesiredWidth.Location = New System.Drawing.Point(307, 45)
        Me.DesiredWidth.Name = "DesiredWidth"
        Me.DesiredWidth.Size = New System.Drawing.Size(100, 20)
        Me.DesiredWidth.TabIndex = 17
        '
        'Power
        '
        Me.Power.Enabled = False
        Me.Power.Location = New System.Drawing.Point(307, 19)
        Me.Power.Name = "Power"
        Me.Power.Size = New System.Drawing.Size(100, 20)
        Me.Power.TabIndex = 15
        '
        'CILevel
        '
        Me.CILevel.Enabled = False
        Me.CILevel.FormattingEnabled = True
        Me.CILevel.Location = New System.Drawing.Point(307, 71)
        Me.CILevel.Name = "CILevel"
        Me.CILevel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CILevel.Size = New System.Drawing.Size(100, 21)
        Me.CILevel.TabIndex = 18
        '
        'CIText
        '
        Me.CIText.AutoSize = True
        Me.CIText.Enabled = False
        Me.CIText.Location = New System.Drawing.Point(50, 74)
        Me.CIText.Name = "CIText"
        Me.CIText.Size = New System.Drawing.Size(90, 13)
        Me.CIText.TabIndex = 17
        Me.CIText.Text = "Confidence Level"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.NIndiv)
        Me.GroupBox1.Controls.Add(Me.NGroups)
        Me.GroupBox1.Controls.Add(Me.PT)
        Me.GroupBox1.Controls.Add(Me.SSLabel)
        Me.GroupBox1.Controls.Add(Me.NIndivCheckBox)
        Me.GroupBox1.Controls.Add(Me.NGCheckBox)
        Me.GroupBox1.Controls.Add(Me.PCheckBox)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 126)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(413, 144)
        Me.GroupBox1.TabIndex = 34
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Pre-Defined Sample Size"
        '
        'NIndiv
        '
        Me.NIndiv.Enabled = False
        Me.NIndiv.Location = New System.Drawing.Point(307, 105)
        Me.NIndiv.Name = "NIndiv"
        Me.NIndiv.Size = New System.Drawing.Size(100, 20)
        Me.NIndiv.TabIndex = 8
        '
        'NGroups
        '
        Me.NGroups.Enabled = False
        Me.NGroups.Location = New System.Drawing.Point(307, 79)
        Me.NGroups.Name = "NGroups"
        Me.NGroups.Size = New System.Drawing.Size(100, 20)
        Me.NGroups.TabIndex = 6
        '
        'PT
        '
        Me.PT.Enabled = False
        Me.PT.Location = New System.Drawing.Point(307, 53)
        Me.PT.Name = "PT"
        Me.PT.Size = New System.Drawing.Size(100, 20)
        Me.PT.TabIndex = 4
        '
        'SSLabel
        '
        Me.SSLabel.AutoSize = True
        Me.SSLabel.Location = New System.Drawing.Point(30, 26)
        Me.SSLabel.Name = "SSLabel"
        Me.SSLabel.Size = New System.Drawing.Size(180, 13)
        Me.SSLabel.TabIndex = 38
        Me.SSLabel.Text = "Please Choose Input Criterion Above"
        '
        'NIndivCheckBox
        '
        Me.NIndivCheckBox.AutoSize = True
        Me.NIndivCheckBox.Enabled = False
        Me.NIndivCheckBox.Location = New System.Drawing.Point(30, 108)
        Me.NIndivCheckBox.Name = "NIndivCheckBox"
        Me.NIndivCheckBox.Size = New System.Drawing.Size(202, 17)
        Me.NIndivCheckBox.TabIndex = 7
        Me.NIndivCheckBox.Text = "Number of Individuals in Each Cluster"
        Me.NIndivCheckBox.UseVisualStyleBackColor = True
        '
        'NGCheckBox
        '
        Me.NGCheckBox.AutoSize = True
        Me.NGCheckBox.Enabled = False
        Me.NGCheckBox.Location = New System.Drawing.Point(30, 82)
        Me.NGCheckBox.Name = "NGCheckBox"
        Me.NGCheckBox.Size = New System.Drawing.Size(142, 17)
        Me.NGCheckBox.TabIndex = 5
        Me.NGCheckBox.Text = "Total Number of Clusters"
        Me.NGCheckBox.UseVisualStyleBackColor = True
        '
        'CostInfo
        '
        Me.CostInfo.Controls.Add(Me.TotalCost)
        Me.CostInfo.Controls.Add(Me.CIndivCost)
        Me.CostInfo.Controls.Add(Me.CGroupCost)
        Me.CostInfo.Controls.Add(Me.TIndivCost)
        Me.CostInfo.Controls.Add(Me.TGroupCost)
        Me.CostInfo.Controls.Add(Me.TotalCostText)
        Me.CostInfo.Controls.Add(Me.CostIndiv)
        Me.CostInfo.Controls.Add(Me.CostGroup)
        Me.CostInfo.Controls.Add(Me.CostCtrl)
        Me.CostInfo.Controls.Add(Me.CostTreatment)
        Me.CostInfo.Enabled = False
        Me.CostInfo.Location = New System.Drawing.Point(13, 276)
        Me.CostInfo.Name = "CostInfo"
        Me.CostInfo.Size = New System.Drawing.Size(413, 130)
        Me.CostInfo.TabIndex = 35
        Me.CostInfo.TabStop = False
        Me.CostInfo.Text = "Cost Information"
        '
        'TotalCost
        '
        Me.TotalCost.Enabled = False
        Me.TotalCost.Location = New System.Drawing.Point(198, 95)
        Me.TotalCost.Name = "TotalCost"
        Me.TotalCost.Size = New System.Drawing.Size(209, 20)
        Me.TotalCost.TabIndex = 13
        '
        'CIndivCost
        '
        Me.CIndivCost.Enabled = False
        Me.CIndivCost.Location = New System.Drawing.Point(307, 69)
        Me.CIndivCost.Name = "CIndivCost"
        Me.CIndivCost.Size = New System.Drawing.Size(100, 20)
        Me.CIndivCost.TabIndex = 12
        '
        'CGroupCost
        '
        Me.CGroupCost.Enabled = False
        Me.CGroupCost.Location = New System.Drawing.Point(307, 43)
        Me.CGroupCost.Name = "CGroupCost"
        Me.CGroupCost.Size = New System.Drawing.Size(100, 20)
        Me.CGroupCost.TabIndex = 10
        '
        'TIndivCost
        '
        Me.TIndivCost.Enabled = False
        Me.TIndivCost.Location = New System.Drawing.Point(198, 69)
        Me.TIndivCost.Name = "TIndivCost"
        Me.TIndivCost.Size = New System.Drawing.Size(100, 20)
        Me.TIndivCost.TabIndex = 11
        '
        'TGroupCost
        '
        Me.TGroupCost.Enabled = False
        Me.TGroupCost.Location = New System.Drawing.Point(198, 43)
        Me.TGroupCost.Name = "TGroupCost"
        Me.TGroupCost.Size = New System.Drawing.Size(100, 20)
        Me.TGroupCost.TabIndex = 9
        '
        'TotalCostText
        '
        Me.TotalCostText.AutoSize = True
        Me.TotalCostText.Enabled = False
        Me.TotalCostText.Location = New System.Drawing.Point(30, 98)
        Me.TotalCostText.Name = "TotalCostText"
        Me.TotalCostText.Size = New System.Drawing.Size(101, 13)
        Me.TotalCostText.TabIndex = 7
        Me.TotalCostText.Text = "Total Cost Available"
        '
        'CostIndiv
        '
        Me.CostIndiv.AutoSize = True
        Me.CostIndiv.Enabled = False
        Me.CostIndiv.Location = New System.Drawing.Point(30, 72)
        Me.CostIndiv.Name = "CostIndiv"
        Me.CostIndiv.Size = New System.Drawing.Size(101, 13)
        Me.CostIndiv.TabIndex = 2
        Me.CostIndiv.Text = "New Individual Cost"
        '
        'CostGroup
        '
        Me.CostGroup.AutoSize = True
        Me.CostGroup.Enabled = False
        Me.CostGroup.Location = New System.Drawing.Point(30, 46)
        Me.CostGroup.Name = "CostGroup"
        Me.CostGroup.Size = New System.Drawing.Size(113, 13)
        Me.CostGroup.TabIndex = 2
        Me.CostGroup.Text = "New Group Fixed Cost"
        '
        'CostCtrl
        '
        Me.CostCtrl.AutoSize = True
        Me.CostCtrl.Location = New System.Drawing.Point(337, 19)
        Me.CostCtrl.Name = "CostCtrl"
        Me.CostCtrl.Size = New System.Drawing.Size(40, 13)
        Me.CostCtrl.TabIndex = 1
        Me.CostCtrl.Text = "Control"
        '
        'CostTreatment
        '
        Me.CostTreatment.AutoSize = True
        Me.CostTreatment.Location = New System.Drawing.Point(221, 19)
        Me.CostTreatment.Name = "CostTreatment"
        Me.CostTreatment.Size = New System.Drawing.Size(55, 13)
        Me.CostTreatment.TabIndex = 0
        Me.CostTreatment.Text = "Treatment"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(270, 562)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "OK"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(351, 562)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 20
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Criteria
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(438, 597)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CostInfo)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.PowerGroupBox)
        Me.Controls.Add(Me.InputCriteria)
        Me.Name = "Criteria"
        Me.Text = "Criteria for Finding Sample Size"
        Me.InputCriteria.ResumeLayout(False)
        Me.InputCriteria.PerformLayout()
        Me.PowerGroupBox.ResumeLayout(False)
        Me.PowerGroupBox.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.CostInfo.ResumeLayout(False)
        Me.CostInfo.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents InputCriteria As System.Windows.Forms.GroupBox
    Friend WithEvents Option2 As System.Windows.Forms.RadioButton
    Friend WithEvents Option1 As System.Windows.Forms.RadioButton
    Friend WithEvents PCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents RPowerWidth2 As System.Windows.Forms.RadioButton
    Friend WithEvents RPowerWidth1 As System.Windows.Forms.RadioButton
    Friend WithEvents PowerGroupBox As System.Windows.Forms.GroupBox
    Friend WithEvents CILevel As System.Windows.Forms.ComboBox
    Friend WithEvents CIText As System.Windows.Forms.Label
    Friend WithEvents Option3 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SSLabel As System.Windows.Forms.Label
    Friend WithEvents NIndivCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents NGCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents CostInfo As System.Windows.Forms.GroupBox
    Friend WithEvents CostTreatment As System.Windows.Forms.Label
    Friend WithEvents CostIndiv As System.Windows.Forms.Label
    Friend WithEvents CostGroup As System.Windows.Forms.Label
    Friend WithEvents CostCtrl As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TotalCostText As System.Windows.Forms.Label
    Friend WithEvents NGroups As System.Windows.Forms.TextBox
    Friend WithEvents PT As System.Windows.Forms.TextBox
    Friend WithEvents DesiredWidth As System.Windows.Forms.TextBox
    Friend WithEvents Power As System.Windows.Forms.TextBox
    Friend WithEvents NIndiv As System.Windows.Forms.TextBox
    Friend WithEvents TotalCost As System.Windows.Forms.TextBox
    Friend WithEvents CIndivCost As System.Windows.Forms.TextBox
    Friend WithEvents CGroupCost As System.Windows.Forms.TextBox
    Friend WithEvents TIndivCost As System.Windows.Forms.TextBox
    Friend WithEvents TGroupCost As System.Windows.Forms.TextBox
    Friend WithEvents DegreeOfCertaintyLabel As System.Windows.Forms.Label
    Friend WithEvents DegreeOfCertainty As System.Windows.Forms.TextBox
    Friend WithEvents DegreeOfCertaintyLabel2 As System.Windows.Forms.Label
    Friend WithEvents IsDegreeOfCertainty As System.Windows.Forms.ComboBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
