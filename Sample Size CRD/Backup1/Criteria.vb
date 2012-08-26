Public Class Criteria

    Dim CILevelSelected = False

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Correct Code for Checked Then Added
        If (CheckNumber() = True) Then
            Main.IsCriteriaDefined = True
            PlugInNumber()
            WriteInputCriteria()
            Me.DialogResult = DialogResult.OK
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Option1.CheckedChanged
        Option1changed()
        OptionValue = 1
    End Sub

    Public Sub Option1changed()
        SSLabel.Text = "Do you wish to specify the sample size results in advance?"
        SetSampleSizeDefault()
        ActivateNGNIndivCheckedBox(True)
        ActivateCostCriteria(False)
        ActivatePowerWidthCriteria(True)
        ActivateTotalCost(False)
        If RPowerWidth1.Checked = True Then Power.Enabled = True
        If RPowerWidth2.Checked = True Then
            DesiredWidth.Enabled = True
            CIText.Enabled = True
            CILevel.Enabled = True
        End If
        PCheckBox.Enabled = True
    End Sub


    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Option2.CheckedChanged
        Option2changed()
        OptionValue = 2
    End Sub

    Public Sub Option2changed()
        SSLabel.Text = "Do you wish to specify the sample size results in advance?"
        SetSampleSizeDefault()
        ActivateNGNIndivCheckedBox(True)
        ActivateCostCriteria(True)
        ActivatePowerWidthCriteria(True)
        ActivateTotalCost(False)
        If RPowerWidth1.Checked = True Then Power.Enabled = True
        If RPowerWidth2.Checked = True Then
            DesiredWidth.Enabled = True
            CIText.Enabled = True
            CILevel.Enabled = True
        End If
        PCheckBox.Enabled = True
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Option3.CheckedChanged
        Option3changed()
        OptionValue = 3
    End Sub

    Public Sub Option3changed()
        SSLabel.Text = "Do you wish to specify the sample size results in advance?"
        SetSampleSizeDefault()
        ActivateNGNIndivCheckedBox(True)
        ActivateCostCriteria(True)
        ActivatePowerWidthCriteria(False)
        ActivateTotalCost(True)
        Power.Enabled = False
        DesiredWidth.Enabled = False
        CIText.Enabled = False
        CILevel.Enabled = False
        PCheckBox.Enabled = True
    End Sub

    Sub SetSampleSizeDefault()
        PCheckBox.Checked = False
        NGCheckBox.Checked = False
        NIndivCheckBox.Checked = False
        NGroups.Enabled = False
        NIndiv.Enabled = False
        PT.Enabled = False
    End Sub

    Sub ActivateCostCriteria(ByVal activate As Boolean)
        TGroupCost.Enabled = activate
        TIndivCost.Enabled = activate
        CGroupCost.Enabled = activate
        CIndivCost.Enabled = activate
        CostCtrl.Enabled = activate
        CostGroup.Enabled = activate
        CostIndiv.Enabled = activate
        CostInfo.Enabled = activate
        CostTreatment.Enabled = activate
    End Sub

    Sub ActivatePowerWidthCriteria(ByVal activate As Boolean)
        RPowerWidth1.Enabled = activate
        RPowerWidth2.Enabled = activate
        PowerGroupBox.Enabled = activate
    End Sub

    Sub ActivateNGNIndivCheckedBox(ByVal activate As Boolean)
        NGCheckBox.Enabled = activate
        NIndivCheckBox.Enabled = activate
    End Sub

    Sub ActivateTotalCost(ByVal activate As Boolean)
        TotalCost.Enabled = activate
        TotalCostText.Enabled = activate
    End Sub

    Function CheckNumber()
        Dim Result As Boolean = True
        Dim ValidateEachSampleSize As Boolean = False
        Dim TempPT As Double
        Dim TempNGroup, TempNIndiv As Integer

        If Not (Option1.Checked = True Or Option2.Checked = True Or Option3.Checked = True) Then
            MsgBox("Please specified options for finding sample sizes in the top box.")
            Return False
        End If

        If (PCheckBox.Checked = True) Or (NGCheckBox.Checked = True) Or (NIndivCheckBox.Checked = True) Then
            ValidateEachSampleSize = True
        End If

        If (PCheckBox.Checked = True) Then
            If (IsDoubleInRange(PT, 0.0, 1.0, "Proportion of treatment groups")) Then
                TempPT = CDbl(PT.Text)
            Else
                Result = False
                ValidateEachSampleSize = False
            End If
        End If

        If (NGCheckBox.Checked = True) Then
            If (IsIntegerLBound(NGroups, 2, "Number of clusters")) Then
                TempNGroup = CInt(NGroups.Text)
            Else
                Result = False
                ValidateEachSampleSize = False
            End If
        End If

        If (NIndivCheckBox.Checked = True) Then
            If (IsIntegerLBound(NIndiv, 2, "Number of individuals in each cluster")) Then
                TempNIndiv = CInt(NIndiv.Text)
            Else
                Result = False
                ValidateEachSampleSize = False
            End If
        End If


        If (ValidateEachSampleSize = True) Then
            If (PCheckBox.Checked = True) And (NGCheckBox.Checked = True) And (NIndivCheckBox.Checked = True) Then
                MsgBox("Uncheck at least one of the following: proportion of treatment clusters, number of total clusters, or number of individuals in each cluster")
                Result = False
            ElseIf (PCheckBox.Checked = True) And (NGCheckBox.Checked = True) And (NIndivCheckBox.Checked = False) Then
                Dim TempNTValue As Integer = Math.Round(TempPT * TempNGroup)
                Dim TempNCValue As Integer = Math.Round((1 - TempPT) * TempNGroup)
                If Not ((TempNTValue + TempNCValue) = TempNGroup) Then
                    MsgBox("The proportion of treatment groups and number of groups do not match, such as proportion = 0.5, but number of groups = 11.")
                    Result = False
                End If
            End If
        End If

        If (Option2.Checked = True Or Option3.Checked = True) Then
            If Not IsDoubleLBound(TGroupCost, 0, "The treatment group cost") Then
                Result = False
            End If
            If Not IsDoubleLBound(TIndivCost, 0, "The treatment individual cost") Then
                Result = False
            Else
                If CDbl(TIndivCost.Text) = 0 Then
                    MsgBox("Please specify treatment individual cost more than 0 because, in reality, it has a cost. It does not matter if you specify in decimal, such as 0.01.")
                    Return False
                End If
            End If
            If Not IsDoubleLBound(CGroupCost, 0, "The control group cost") Then
                Result = False
            End If
            If Not IsDoubleLBound(CIndivCost, 0, "The control individual cost") Then
                Result = False
            Else
                If CDbl(CIndivCost.Text) = 0 Then
                    MsgBox("Please specify control individual cost more than 0 because, in reality, it has a cost. It does not matter if you specify in decimal, such as 0.01.")
                    Return False
                End If
            End If
        End If
        If (Option3.Checked = True) Then
            If Not IsDoubleLBound(TotalCost, 0, "The treatment group cost") Then
                Result = False
            Else
                Try
                    Dim PossibleTotalCost As Double = Storage.FindTotalCostExact(2, 2, 3, CDbl(TGroupCost.Text), CDbl(TIndivCost.Text), CDbl(CGroupCost.Text), CDbl(CIndivCost.Text))
                    If CDbl(TotalCost.Text < PossibleTotalCost) Then
                        MsgBox("The total cost should be greater than " & TotalCost.Text & ".")
                        Result = False
                    End If
                Catch ex As Exception

                End Try
            End If
        End If

        If (Option1.Checked = True Or Option2.Checked = True) Then
            If Not ((RPowerWidth1.Checked = True) Or (RPowerWidth2.Checked = True)) Then
                MsgBox("Since option1 or option2 was selected, the power or width choice must be chosen.")
                Return False
            End If
            If (RPowerWidth1.Checked = True) AndAlso (Not IsDoubleInRange(Power, 0, 1, "Power")) Then
                Result = False
            End If
            If (RPowerWidth2.Checked = True) Then
                If (Not IsDoubleInRange(DesiredWidth, 0, 3, "Desired Width of CI of Effect Size")) Then
                    Result = False
                End If
                If Not (CILevel.SelectedIndex = 0 Or CILevel.SelectedIndex = 1) Then
                    Result = False
                    MsgBox("Please select confidence level since the width of CI of effect size was selected.")
                End If
                If Not (IsDegreeOfCertainty.SelectedIndex = 0 Or IsDegreeOfCertainty.SelectedIndex = 1) Then
                    Result = False
                    MsgBox("Please select whether the degree of certainty in finding the sample sizes by width of CI of effect size is used.")
                End If
                If (IsDegreeOfCertainty.SelectedIndex = 1) Then
                    If Not IsDoubleInRange(DegreeOfCertainty, 0.5, 0.99, "Degree of Certainty") Then
                        Result = False
                    Else
                        If CDbl(DegreeOfCertainty.Text) = 0 Then
                            MsgBox("Please specify degree of certainty betweeen .50 to .99.")
                            Return False
                        End If
                    End If
                End If
            End If
        End If
        Return Result
    End Function

    Private Sub PlugInNumber()
        If (PCheckBox.Checked = True) And (NGCheckBox.Checked = True) And (NIndivCheckBox.Checked = False) Then
            Dim TempNTValue As Integer = Math.Round(CDbl(PT.Text) * CInt(NGroups.Text))
            Dim TempNCValue As Integer = Math.Round((1 - CDbl(PT.Text)) * CInt(NGroups.Text))
            NTValue = TempNTValue
            NCValue = TempNCValue
            NGValue = CInt(NGroups.Text)
            PTValue = CDbl(PT.Text)
        Else
            If (PCheckBox.Checked = True) Then
                PTValue = CDbl(PT.Text)
                NTValue = -1
                NCValue = -1
                NGValue = -1
            End If
            If (NGCheckBox.Checked = True) Then
                NGValue = CInt(NGroups.Text)
                NTValue = -1
                NCValue = -1
                PTValue = -1
            End If
            If (NIndivCheckBox.Checked = True) Then
                NIndivValue = CInt(NIndiv.Text)
            End If
        End If

        If (Option1.Checked = True) Then
            TGroupCostValue = 0
            CGroupCostValue = 0
            TIndivCostValue = 1
            CIndivCostValue = 1
        End If

        If (Option2.Checked = True Or Option3.Checked = True) Then
            TGroupCostValue = CDbl(TGroupCost.Text)
            TIndivCostValue = CDbl(TIndivCost.Text)
            CGroupCostValue = CDbl(CGroupCost.Text)
            CIndivCostValue = CDbl(CIndivCost.Text)
        End If

        If (Option3.Checked = True) Then
            TotalCostValue = CDbl(TotalCost.Text)
        End If

        If (Option1.Checked = True Or Option2.Checked = True) Then
            If (RPowerWidth1.Checked = True) Then
                PowerValue = CDbl(Power.Text)
                PowerOrWidthValue = 1
            End If
            If (RPowerWidth2.Checked = True) Then
                WidthValue = CDbl(DesiredWidth.Text)
                PowerOrWidthValue = 2
                If (CILevel.SelectedIndex = 0) Then
                    CILevelValue = 0.95
                ElseIf (CILevel.SelectedIndex = 1) Then
                    CILevelValue = 0.99
                End If
                If Storage.IsDegreeOfCertaintyValue = True Then
                    Storage.DegreeOfCertaintyValue = CDbl(DegreeOfCertainty.Text)
                End If
            End If

        End If
    End Sub

    Private Sub WriteInputCriteria()
        Dim text As String = ""
        If (Option1.Checked = True) Then
            text = "Min total number of individuals by "
        ElseIf (Option2.Checked = True) Then
            text = "Min cost by "
        ElseIf (Option3.Checked = True) Then
            text = "Max power/ Min width of CI of effect size with limited cost."
        Else
            text = "Option did not checked"
        End If

        If (Option1.Checked = True) Or (Option2.Checked = True) Then
            If RPowerWidth1.Checked = True Then
                text &= "power of " & PowerValue.ToString("f3") & vbCrLf
            ElseIf RPowerWidth2.Checked = True Then
                text &= vbCrLf & CILevelValue.ToString("f2") & " width of CI of Effect size of " & WidthValue.ToString("f3") & vbCrLf
                If Storage.IsDegreeOfCertaintyValue = True Then
                    text &= "The degree of certainty is specifed as " & Storage.DegreeOfCertaintyValue.ToString("f2") & vbCrLf
                Else
                    text &= "The degree of certainty is not specified (The average width is used)."
                End If
            Else
                text &= "with no criteria specified" & vbCrLf
            End If
        End If

        If (Option3.Checked = True) Then
            text &= vbCrLf & "Total cost = " & TotalCostValue.ToString("n0") & vbCrLf
        End If

        If (Option2.Checked = True) Or (Option3.Checked = True) Then
            text &= "Cost Criteria" & vbCrLf
            text &= "Treatment Group Cost = " & TGroupCostValue.ToString("f2") & vbCrLf
            text &= "Treatment Individual Cost = " & TIndivCostValue.ToString("f2") & vbCrLf
            text &= "Control Group Cost = " & CGroupCostValue.ToString("f2") & vbCrLf
            text &= "Control Individual Cost = " & CIndivCostValue.ToString("f2") & vbCrLf
            Main.TotalCostLabel.Text = "Total Cost"
        Else
            Main.TotalCostLabel.Text = "Total Number of Individuals"
        End If

        If (PCheckBox.Checked = True) Or (NGCheckBox.Checked = True) Or (NIndivCheckBox.Checked = True) Then
            text &= "Sample Size Criteria" & vbCrLf
            If (PCheckBox.Checked = True) Then
                text &= "Proportion of treatment groups = " & PTValue.ToString("f3") & vbCrLf
            End If
            If (NGCheckBox.Checked = True) Then
                text &= "Number of all groups = " & NGValue.ToString("n0") & vbCrLf
            End If
            If (NIndivCheckBox.Checked = True) Then
                text &= "Number of individuals in each group = " & NIndivValue.ToString("n0") & vbCrLf
            End If
        End If
        Main.SSCriteria.Text = text
    End Sub

    Private Sub PCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PCheckBox.CheckedChanged
        PCheckBoxChanged()
    End Sub

    Public Sub PCheckBoxChanged()
        PT.Enabled = PCheckBox.Checked
    End Sub

    Private Sub CILevel_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CILevel.SelectedIndexChanged
        Select Case CILevel.SelectedIndex
            Case 0
                CILevelValue = 0.95
            Case 1
                CILevelValue = 0.99
        End Select
        CILevelSelected = True
    End Sub

    Private Sub NGCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NGCheckBox.CheckedChanged
        NGCheckBoxChanged()
    End Sub

    Public Sub NGCheckBoxChanged()
        NGroups.Enabled = NGCheckBox.Checked
    End Sub

    Private Sub NIndivCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NIndivCheckBox.CheckedChanged
        NIndivCheckBoxChanged()
    End Sub

    Public Sub NIndivCheckBoxChanged()
        NIndiv.Enabled = NIndivCheckBox.Checked
    End Sub

    Private Sub RPowerWidth1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RPowerWidth1.CheckedChanged
        PowerOrWidthValue = 1 'For Power
        RPowerWidth1Changed()
    End Sub

    Public Sub RPowerWidth1Changed()
        enabledPowerNotWidth(True)
    End Sub

    Private Sub RPowerWidth2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RPowerWidth2.CheckedChanged
        PowerOrWidthValue = 2 'For Desired Width
        RPowerWidth2Changed()
    End Sub

    Public Sub RPowerWidth2Changed()
        enabledPowerNotWidth(False)
    End Sub

    Private Sub enabledPowerNotWidth(ByVal activate As Boolean)
        Power.Enabled = activate
        DesiredWidth.Enabled = Not activate
        CILevel.Enabled = Not activate
        CIText.Enabled = Not activate
        If Main.MonteCarlo.Checked = True Then
            DegreeOfCertaintyLabel.Enabled = Not activate
            IsDegreeOfCertainty.Enabled = Not activate
            If Not activate And (IsDegreeOfCertainty.SelectedIndex = 1) Then
                DegreeOfCertaintyLabel2.Enabled = True
                DegreeOfCertainty.Enabled = True
            Else
                DegreeOfCertaintyLabel2.Enabled = False
                DegreeOfCertainty.Enabled = False
            End If
        Else
            DegreeOfCertaintyLabel.Enabled = False
            IsDegreeOfCertainty.Enabled = False
            IsDegreeOfCertainty.SelectedIndex = 0
            DegreeOfCertaintyLabel2.Enabled = False
            DegreeOfCertainty.Enabled = False
        End If
    End Sub

    Private Sub IsDegreeOfCertainty_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IsDegreeOfCertainty.SelectedIndexChanged
        Select Case IsDegreeOfCertainty.SelectedIndex
            Case 0
                Storage.IsDegreeOfCertaintyValue = False
                DegreeOfCertaintyLabel2.Enabled = False
                DegreeOfCertainty.Enabled = False
            Case 1
                Storage.IsDegreeOfCertaintyValue = True
                DegreeOfCertaintyLabel2.Enabled = True
                DegreeOfCertainty.Enabled = True
        End Select
    End Sub

    Private Sub Criteria_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class