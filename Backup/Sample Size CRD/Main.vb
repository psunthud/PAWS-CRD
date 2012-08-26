Imports System.Diagnostics
Imports System.IO

Public Class Main

    Public IsCriteriaDefined As Boolean = False
    Public IsMplusInstalled As Boolean
    Public checkCovariateInstanceVariable As Boolean = False
    Public RunNullModelInstanceVariable As Boolean = False

    'Private RunICCCondition As Integer = 0
    'Private RunESCondition As Integer = 0
    'Private RunCostCondition As Integer = 0
    'Private RunCovariate As Integer = 0

    Private Sub Allcommand()
        TabControl1.SelectedIndex = 1
        Storage.NRep = 1000
        IsCriteriaDefined = True
    End Sub

    Private Sub Block1command()
        PowerOrWidthValue = 1
        PowerValue = 0.8
        Criteria.RPowerWidth2.Checked = False
        Criteria.RPowerWidth1.Checked = True
        Criteria.Power.Text = PowerValue.ToString("f3")
        Storage.OptionValue = 2 'Depend on condition
    End Sub

    Private Sub block2command()
        PowerOrWidthValue = 2
        WidthValue = 0.2
        CILevelValue = 0.95
        IsDegreeOfCertaintyValue = False
        Criteria.IsDegreeOfCertainty.SelectedIndex = 0
        Criteria.DegreeOfCertainty.Text = ""
        Criteria.CILevel.SelectedIndex = 0
        Criteria.RPowerWidth1.Checked = False
        Criteria.RPowerWidth2.Checked = True
        Criteria.DesiredWidth.Text = WidthValue.ToString("f3")
        Storage.OptionValue = 2 'Depend on condition
    End Sub

    Private Sub Block3command()
        PowerOrWidthValue = 2
        WidthValue = 0.5
        CILevelValue = 0.95
        IsDegreeOfCertaintyValue = False
        Criteria.IsDegreeOfCertainty.SelectedIndex = 0
        Criteria.DegreeOfCertainty.Text = ""
        Criteria.CILevel.SelectedIndex = 0
        Criteria.RPowerWidth1.Checked = False
        Criteria.RPowerWidth2.Checked = True
        Criteria.DesiredWidth.Text = WidthValue.ToString("f3")
        Storage.OptionValue = 2 'Depend on condition

    End Sub

    Private Sub Block4Command()
        TotalCostValue = 500
        Criteria.TotalCost.Text = TotalCostValue.ToString("f2")
        Storage.OptionValue = 3 'Depend on condition

    End Sub

    Private Sub Block5command()
        TotalCostValue = 1000
        Criteria.TotalCost.Text = TotalCostValue.ToString("f2")
        Storage.OptionValue = 3 'Depend on condition

    End Sub

    Private Sub Allcommand2()
        If Storage.OptionValue = 1 Then
            Criteria.Option1.Checked = True
        Else
            Criteria.Option1.Checked = False
        End If
        If Storage.OptionValue = 2 Then
            Criteria.Option2.Checked = True
        Else
            Criteria.Option2.Checked = False
        End If
        If Storage.OptionValue = 3 Then
            Criteria.Option3.Checked = True
        Else
            Criteria.Option3.Checked = False
        End If

    End Sub

    Private Sub findPowerForNonMonte()
        TabControl1.SelectedIndex() = 0
        MonteCarlo.Checked = False
    End Sub


    Function ReadSampleSizeFile(ByVal filename As String, ByVal length As Integer)
        Dim numLine As Short = 0
        Dim Lines(0 To length - 1) As String
        Dim Result(0 To length - 1, 0 To 2) As Integer
        Dim sr As StreamReader = New StreamReader(filename)

        For i = 0 To length - 1
            Lines(i) = sr.ReadLine()
        Next i

        sr.Close()


        Dim splitSentence() As String
        For i = 0 To length - 1
            splitSentence = Lines(i).Split(vbTab)
            For j = 0 To UBound(splitSentence)
                Result(i, j) = CInt(splitSentence(j))
            Next j
        Next i





        Return Result

        'Dim es(0 To 5) As Double
        'For i = 0 To 4
        '    es(i) = CDbl(cleanSentence(i + 2))
        'Next i

        'find = "Between Level"
        'numResult = 0
        'For i = 0 To numLine - 1
        '    If fileArray(i).Contains(find) Then
        '        findResult(numResult) = i
        '        numResult += 1
        '    End If
        'Next i

        'Dim powerLine As String = fileArray(findResult(0) + 3)
        'Dim powerSentence() As String = powerLine.Split(" "c)

        'currentLine = 0
        'Dim cleanPowerSentence(0 To 5) As String
        'For i = 0 To UBound(powerSentence)
        '    If powerSentence(i) <> "" AndAlso powerSentence(i) <> " " Then
        '        cleanPowerSentence(currentLine) = powerSentence(i)
        '        currentLine += 1
        '    End If
        'Next i
        'Dim sig As Double = 0.0
        'If (CDbl(cleanPowerSentence(5)) <= 0.05) Then
        '    sig = 1.0
        'End If
        'es(5) = sig
        'Return es
    End Function

    Private Sub Calculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Calculate.Click
        'Dim ICCCondition() As Double = {0.05, 0.25}
        'Dim ESCondition() As Double = {0.2, 0.5}
        'Dim CostCondition() As Double = {0, 5, 10}
        'Dim ICCZCondition() As Double = {0, 0, 0.05, 0.25, 1.0}
        'Dim output As String = ""
        ''Dim start As String
        'Dim samplesize(,) As Integer
        'Dim runIndex As Integer = 0
        'Dim desiredresult(0 To 59, 0 To 2) As String
        'Dim samplesizepath As String = "C:\Users\Sunthud\Documents\Program Results\"
        'Dim type As String = "BUD1000"
        'samplesize = ReadSampleSizeFile(samplesizepath & type & "ss.txt", 60)


        'For runICCCondition = 0 To UBound(ICCCondition)
        '    For runESCondition = 0 To UBound(ESCondition)
        '        For runCovariateType = 0 To 4

        '            For runCostCondition = 0 To UBound(CostCondition)
        '                findPowerForNonMonte()

        '                'Allcommand()

        '                ''''''''''''''' BLOCK 1 Power (.80) ''''''''''''''''''''''''''
        '                ''Block1command()
        '                ''start = "\power"

        '                '''''''''''''''''''BLOCK 2 95% CIES width = 0.2 '''''''''''''
        '                'block2command()
        '                'start = "\CI2"

        '                '''''''''''''''''BLOCK 3 95% CIES width = 0.5 ''''''''''''''''
        '                ''Block3command()
        '                ''start = "\CI5"

        '                ''''''''''''''''''BLOCK 4 BUDGET = 500 ''''''''''''''''''
        '                ''Block4Command()
        '                ''start = "\BUD500"

        '                ''''''''''''''''''''BLOCK 5 BUDGET = 1000 ''''''''''''''''
        '                ''Block5command()
        '                ''start = "\BUD1000"

        '                'Allcommand2()
        '                'ICCY2.Text = ICCCondition(runICCCondition).ToString("f3")
        '                'ES2.Text = ESCondition(runESCondition).ToString("f3")

        '                'If runCovariateType = 0 Then 'No Covariate
        '                '    checkCovariate.Checked = False
        '                '    ClusterOnlyZ.Checked = False
        '                '    ICCZ2.Text = ""
        '                '    R2ClusterZ2.Text = ""
        '                '    R2IndivZ2.Text = ""
        '                'Else
        '                '    checkCovariate.Checked = True
        '                '    If runCovariateType = 4 Then 'ClusterLevel covariate
        '                '        ClusterOnlyZ.Checked = True
        '                '        ICCZ2.Text = "1"
        '                '        R2ClusterZ2.Text = "0.13"
        '                '        R2IndivZ2.Text = "0"
        '                '    Else 'Different ICC
        '                '        ClusterOnlyZ.Checked = False
        '                '        ICCZ2.Text = ICCZCondition(runCovariateType).ToString("f3")
        '                '        If runCovariateType = 1 Then
        '                '            R2ClusterZ2.Text = "0"
        '                '        Else
        '                '            R2ClusterZ2.Text = "0.13"
        '                '        End If
        '                '        R2IndivZ2.Text = "0.13"
        '                '    End If
        '                'End If

        '                'TIndivCostValue = 1
        '                'CIndivCostValue = 1
        '                'TGroupCostValue = TIndivCostValue * CostCondition(runCostCondition)
        '                'CGroupCostValue = CIndivCostValue * CostCondition(runCostCondition)
        '                'Criteria.TIndivCost.Text = TIndivCostValue.ToString("f3")
        '                'Criteria.CIndivCost.Text = CIndivCostValue.ToString("f3")
        '                'Criteria.TGroupCost.Text = TGroupCostValue.ToString("f3")
        '                'Criteria.CGroupCost.Text = CGroupCostValue.ToString("f3")
        '                'output = start & runICCCondition.ToString("n0") & runESCondition.ToString("n0") & runCovariateType.ToString("n0") & runCostCondition.ToString("n0") & ".txt"
        '                'run()
        '                'WriteOutput(My.Application.Info.DirectoryPath & output)
        '                'Storage.Reset()


        '                NTGroups.Text = samplesize(runIndex, 0).ToString("n0")
        '                NCGroups.Text = samplesize(runIndex, 1).ToString("n0")
        '                NIndiv.Text = samplesize(runIndex, 2).ToString("n0")

        '                ICCY.Text = ICCCondition(runICCCondition).ToString("f3")
        '                ES.Text = ESCondition(runESCondition).ToString("f3")

        '                If runCovariateType = 0 Then 'No Covariate
        '                    checkCovariate.Checked = False
        '                    ClusterOnlyZ.Checked = False
        '                    ICCZ.Text = ""
        '                    R2ClusterZ.Text = ""
        '                    R2IndivZ.Text = ""
        '                Else
        '                    checkCovariate.Checked = True
        '                    If runCovariateType = 4 Then 'ClusterLevel covariate
        '                        ClusterOnlyZ.Checked = True
        '                        ICCZ.Text = "1"
        '                        R2ClusterZ.Text = "0.13"
        '                        R2IndivZ.Text = "0"
        '                    Else 'Different ICC
        '                        ClusterOnlyZ.Checked = False
        '                        ICCZ.Text = ICCZCondition(runCovariateType).ToString("f3")
        '                        If runCovariateType = 1 Then
        '                            R2ClusterZ.Text = "0"
        '                        Else
        '                            R2ClusterZ.Text = "0.13"
        '                        End If
        '                        R2IndivZ.Text = "0.13"
        '                    End If
        '                End If


        '                run()

        '                If runCovariateType = 0 Then
        '                    desiredresult(runIndex, 0) = Power.Text
        '                    desiredresult(runIndex, 1) = WidthCI95.Text
        '                    desiredresult(runIndex, 2) = WidthCI99.Text
        '                Else
        '                    desiredresult(runIndex, 0) = PowerC.Text
        '                    desiredresult(runIndex, 1) = WidthCI95C.Text
        '                    desiredresult(runIndex, 2) = WidthCI99C.Text
        '                End If
        '                Storage.Reset()
        '                runIndex += 1
        '            Next runCostCondition
        '        Next runCovariateType
        '    Next runESCondition
        'Next runICCCondition

        'Dim textOut As New StreamWriter(New FileStream(samplesizepath & type & "result.txt", FileMode.Create, FileAccess.Write))
        'For i = 0 To desiredresult.GetLength(0) - 1
        '    textOut.WriteLine(desiredresult(i, 0) & vbTab & desiredresult(i, 1) & vbTab & desiredresult(i, 2))
        'Next i
        'textOut.Close()

        run()


    End Sub

    Private Sub run()
        Dim Proceed As Boolean = ValidateInputs() 'AndAlso CheckPossibleParameter() UNABLE because all covariate inputs are okay
        While DetectProcess("Mplus")
            MsgBox("Please close the Mplus before running the program")
        End While
        If (Proceed = True) Then
            WriteInput()
            CleanOutput()
            If (TabControl1.SelectedIndex = 0) Then
                If MonteCarlo.Checked = True Then
                    status.Text = "Running Simulation"
                    RunSimulation(CInt(NTGroups.Text), CInt(NCGroups.Text), CInt(NIndiv.Text))
                Else
                    Dim pT As Double = NTValue / (NTValue + NCValue)
                    If checkCovariate.Checked = True Then
                        Dim residual() As Double = FindResidual(True, pT, NTValue + NCValue, NIndivValue, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                        Dim Var As Double = FindVar(NTValue, NCValue, NIndivValue, residual(0), residual(1))
                        PowerC.Text = FindPower(ESValue, Var).ToString("f3")
                        WidthCI95C.Text = PrintWidth(ESValue, FindWidthCIES(Var)) 'FindWidthCIES(Var).ToString("f3")
                        WidthCI99C.Text = PrintWidth(ESValue, FindWidthCIES(Var, 0.99)) 'FindWidthCIES(Var, 0.99).ToString("f3")
                        If RunNullModel.Checked = True Then
                            residual = FindResidual(False, pT, NTValue + NCValue, NIndivValue, ESValue, ICCYValue)
                            Var = FindVar(NTValue, NCValue, NIndivValue, residual(0), residual(1))
                            Power.Text = FindPower(ESValue, Var).ToString("f3")
                            WidthCI95.Text = PrintWidth(ESValue, FindWidthCIES(Var)) 'FindWidthCIES(Var).ToString("f3")
                            WidthCI99.Text = PrintWidth(ESValue, FindWidthCIES(Var, 0.99)) 'FindWidthCIES(Var, 0.99).ToString("f3")
                        End If
                    Else
                        Dim residual() As Double = FindResidual(False, pT, NTValue + NCValue, NIndivValue, ESValue, ICCYValue)
                        Dim Var As Double = FindVar(NTValue, NCValue, NIndivValue, residual(0), residual(1))
                        Power.Text = FindPower(ESValue, Var).ToString("f3")
                        WidthCI95.Text = PrintWidth(ESValue, FindWidthCIES(Var)) 'FindWidthCIES(Var).ToString("f3")
                        WidthCI99.Text = PrintWidth(ESValue, FindWidthCIES(Var, 0.99)) 'FindWidthCIES(Var, 0.99).ToString("f3")
                    End If
                End If
            ElseIf (TabControl1.SelectedIndex = 1) Then
                status.Text = "Find value algebraically"
                Dim StartArray() As Double = FindStartingValue()
                Dim StartValue() As Integer = {CInt(StartArray(0)), CInt(StartArray(1)), CInt(StartArray(2))}
                Dim pT As Double = StartArray(3)
                ProgressBar1.Value = 50
                StatusNum.Text = 5
                Storage.StartingValue = StartValue
                If MonteCarlo.Checked = True Then
                    status.Text = "Modify sample size by A Priori Monte Carlo Simulation"
                    Dim FinalValue() As Integer = FindingSampleSizeBySimulation(StartValue(0), StartValue(1), StartValue(2), pT)
                    status.Text = "Summarize the result"
                    ProgressBar1.Value = 850
                    StatusNum.Text = 85
                    If Not (StartValue(0) = -1 Or StartValue(1) = -1 Or StartValue(2) = -1 Or pT = -1) Then
                        RunSimulation(FinalValue(0), FinalValue(1), FinalValue(2))
                        NIndiv2.Text = FinalValue(2).ToString("n0")
                        NTGroups2.Text = FinalValue(0).ToString("n0")
                        NCGroups2.Text = FinalValue(1).ToString("n0")
                        TotalCost.Text = FindTotalCostExact(FinalValue(0), FinalValue(1), FinalValue(2), TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue).ToString("f0")
                    Else
                        NIndiv2.Text = "N/A"
                        NTGroups2.Text = "N/A"
                        NCGroups2.Text = "N/A"
                    End If

                Else
                    NIndiv2.Text = StartValue(2).ToString("n0")
                    NTGroups2.Text = StartValue(0).ToString("n0")
                    NCGroups2.Text = StartValue(1).ToString("n0")
                    TotalCost.Text = FindTotalCostExact(StartValue(0), StartValue(1), StartValue(2), TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue).ToString("f0")
                    If checkCovariate.Checked = True Then
                        Dim residual() As Double = FindResidual(True, pT, StartValue(0) + StartValue(1), StartValue(2), ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                        Dim Var As Double = FindVar(StartValue(0), StartValue(1), StartValue(2), residual(0), residual(1))
                        ObtainedPowerC.Text = FindPower(ESValue, Var).ToString("f3")
                        ObtainedWidth95C.Text = PrintWidth(ESValue, FindWidthCIES(Var)) 'FindWidthCIES(Var).ToString("f3")
                        ObtainedWidth99C.Text = PrintWidth(ESValue, FindWidthCIES(Var, 0.99)) 'FindWidthCIES(Var, 0.99).ToString("f3")
                        If RunNullModel.Checked = True Then
                            residual = FindResidual(False, pT, StartValue(0) + StartValue(1), StartValue(2), ESValue, ICCYValue)
                            Var = FindVar(StartValue(0), StartValue(1), StartValue(2), residual(0), residual(1))
                            ObtainedPower.Text = FindPower(ESValue, Var).ToString("f3")
                            ObtainedWidth95.Text = PrintWidth(ESValue, FindWidthCIES(Var)) 'FindWidthCIES(Var).ToString("f3")
                            ObtainedWidth99.Text = PrintWidth(ESValue, FindWidthCIES(Var, 0.99)) 'FindWidthCIES(Var, 0.99).ToString("f3")
                        End If
                    Else
                        Dim residual() As Double = FindResidual(False, pT, StartValue(0) + StartValue(1), StartValue(2), ESValue, ICCYValue)
                        Dim Var As Double = FindVar(StartValue(0), StartValue(1), StartValue(2), residual(0), residual(1))
                        ObtainedPower.Text = FindPower(ESValue, Var).ToString("f3")
                        ObtainedWidth95.Text = PrintWidth(ESValue, FindWidthCIES(Var)) 'FindWidthCIES(Var).ToString("f3")
                        ObtainedWidth99.Text = PrintWidth(ESValue, FindWidthCIES(Var, 0.99)) 'FindWidthCIES(Var, 0.99).ToString("f3")
                    End If
                End If
            End If
            ProgressBar1.Value = 1000
            StatusNum.Text = 100
            status.Text = "Done"
            ActivateDegreeofCertainty()
            WriteOutput(My.Application.Info.DirectoryPath & "\output.txt")
        End If
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        Options.ShowDialog()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DefineCriteria.Click
        Criteria.ShowDialog()
    End Sub

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Criteria.CILevel.Items.Add("95%") 'Add the level in CI Level drop-down list
        Criteria.CILevel.Items.Add("99%")
        Criteria.IsDegreeOfCertainty.Items.Add("No")
        Criteria.IsDegreeOfCertainty.Items.Add("Yes (.50 to .99)")
        IsMplusInstalled = checkMplus() 'Check whether the Mplus is installed in the program.
        If (My.Settings.ShowHelper.Equals(False)) Then StartUp.ShowDialog() 'Does not show the startup window again if users do not prefer
        Try
            Storage.NRep = CInt(My.Settings.StorageNRep)
        Catch ex As Exception
            Storage.NRep = Storage.DefaultReplication
        End Try
        If IsMplusInstalled = True Then 'Disable the Monte Carlo options if the Mplus is not installed
            Dim MplusExist As Boolean = System.IO.File.Exists(My.Settings.MplusDirectory)
            If MplusExist = False Then
                If System.IO.File.Exists("C:\Program Files\Mplus\mplus.exe") Then
                    My.Settings.MplusDirectory = "C:\Program Files\Mplus\mplus.exe"
                    MplusExist = True
                End If
                If System.IO.File.Exists("C:\Program Files (x86)\Mplus\mplus.exe") Then
                    My.Settings.MplusDirectory = "C:\Program Files (x86)\Mplus\mplus.exe"
                    MplusExist = True
                End If
            End If
            While Not MplusExist
                Dim open As OpenFileDialog = New OpenFileDialog
                open.Filter = "Application (*.exe)|*.exe"
                open.Title = "Please browse the 'mplus.exe' in your file directory"
                If open.ShowDialog() = DialogResult.OK Then
                    My.Settings.MplusDirectory = open.FileName
                    If My.Settings.MplusDirectory.Contains("Mplus.exe") Then MplusExist = True
                    If My.Settings.MplusDirectory.Contains("MplusWin.exe") Then MsgBox("Please browse the mplus file, NOT the mpluswin file.")
                Else
                    MsgBox("Since the Mplus directory is not selected, the a priori Monte Carlo Simulation cannot be run.")
                    MonteCarlo.Enabled = False
                    MonteCarlo.Checked = False
                    MplusExist = True
                End If
            End While
        Else
            MonteCarlo.Enabled = False
            MonteCarlo.Checked = False
        End If
        Dim di As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(Path) 'Create path to run the program
        ' Determine whether the directory exists.
        If Not di.Exists Then
            di.Create()
        End If
        ProgressBar1.Minimum = 0 'Adjust the progress bar
        ProgressBar1.Maximum = 1000
        While DetectProcess("Mplus")
            MsgBox("Please close the Mplus before running the program")
        End While
    End Sub

    Private Sub RunSimulation(ByVal NT As Integer, ByVal NC As Integer, ByVal NIndiv As Integer)
        'Process.Start("c:\Program Files\Mplus\Mplus.exe c:\CRD\simulation.inp c:\CRD\simulation.out")
        If (checkCovariate.Checked = True) Then
            WriteMCCovariateCode(NT, NC, NIndiv, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, True, ClusterOnlyZ.Checked, ReverseZ.Checked)
        Else
            WriteMCNullCode(NT, NC, NIndiv, ICCYValue, ESValue, NRep)
        End If
        Dim CorrectedResult(,) As Double
        Dim sim As String = "simulation"
        RunAnalysis(sim)
        System.Threading.Thread.Sleep(5000)
        If (TabControl1.SelectedIndex = 1) Then
            ProgressBar1.Value = 900
            StatusNum.Text = 90
        Else
            ProgressBar1.Value = 500
            StatusNum.Text = 50
        End If
        If checkCovariate.Checked = False Then
            Dim result(,) As Double = RunEachSimulatedData(False, NRep, False, ClusterOnlyZ.Checked, NT + NC, NIndiv)
            Dim sumWidth95 As Double = 0, sumWidth99 As Double = 0, sumEst As Double = 0, sumPower As Double = 0
            Dim Savedata(0 To NRep - 1, 0 To 3) As Double
            Dim ArrayWidth95(0 To NRep - 1) As Double
            Dim ArrayWidth99(0 To NRep - 1) As Double
            Dim width95, width99 As Double
            For i = 0 To NRep - 1
                While result(i, 5) = -1
                    WriteMCNullCode(NT, NC, NIndiv, ICCYValue, ESValue, 1)
                    RunAnalysis(sim)
                    CorrectedResult = RunEachSimulatedData(False, 1, False, ClusterOnlyZ.Checked, NT + NC, NIndiv)
                    For k = 0 To 5
                        result(i, k) = CorrectedResult(0, k)
                    Next k
                End While
                width95 = result(i, 3) - result(i, 1)
                width99 = result(i, 4) - result(i, 0)
                ArrayWidth95(i) = width95
                ArrayWidth99(i) = width99
                sumEst += result(i, 2)
                sumPower += result(i, 5)
                sumWidth95 += width95
                sumWidth99 += width99
            Next i
            Array.Sort(ArrayWidth95)
            Array.Sort(ArrayWidth99)
            For i = 0 To NRep - 1
                Savedata(i, 0) = ArrayWidth95(i)
                Savedata(i, 1) = ArrayWidth99(i)
            Next i
            Storage.Width = Savedata
            If TabControl1.SelectedIndex = 0 Then
                WidthCI95.Text = PrintWidth(ESValue, (sumWidth95 / NRep)) '(sumWidth95 / NRep).ToString("f3")
                WidthCI99.Text = PrintWidth(ESValue, (sumWidth99 / NRep)) '(sumWidth99 / NRep).ToString("f3")
                Power.Text = (sumPower / NRep).ToString("f3")
            Else
                ObtainedWidth95.Text = PrintWidth(ESValue, (sumWidth95 / NRep)) '(sumWidth95 / NRep).ToString("f3")
                ObtainedWidth99.Text = PrintWidth(ESValue, (sumWidth99 / NRep)) '(sumWidth99 / NRep).ToString("f3")
                ObtainedPower.Text = (sumPower / NRep).ToString("f3")
            End If
        End If

        If checkCovariate.Checked = True Then
            Dim result(,) As Double = RunEachSimulatedData(True, NRep, False, ClusterOnlyZ.Checked, NT + NC, NIndiv)
            Dim sumWidth95 As Double = 0, sumWidth99 As Double = 0, sumEst As Double = 0, sumPower As Double = 0
            Dim width95, width99 As Double
            Dim Savedata(0 To NRep - 1, 0 To 3) As Double
            Dim ArrayWidth95(0 To NRep - 1) As Double
            Dim ArrayWidth99(0 To NRep - 1) As Double
            For i = 0 To NRep - 1
                While result(i, 5) = -1
                    WriteMCCovariateCode(NT, NC, NIndiv, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, 1, True, ClusterOnlyZ.Checked, ReverseZ.Checked)
                    RunAnalysis(sim)
                    CorrectedResult = RunEachSimulatedData(True, 1, False, ClusterOnlyZ.Checked, NT + NC, NIndiv)
                    For k = 0 To 5
                        result(i, k) = CorrectedResult(0, k)
                    Next k
                End While
                width95 = result(i, 3) - result(i, 1)
                width99 = result(i, 4) - result(i, 0)
                ArrayWidth95(i) = width95
                ArrayWidth99(i) = width99
                sumEst += result(i, 2)
                sumPower += result(i, 5)
                sumWidth95 += width95
                sumWidth99 += width99
            Next i
            Array.Sort(ArrayWidth95)
            Array.Sort(ArrayWidth99)
            For i = 0 To NRep - 1
                Savedata(i, 2) = ArrayWidth95(i)
                Savedata(i, 3) = ArrayWidth99(i)
            Next i
            Storage.Width = Savedata
            If TabControl1.SelectedIndex = 0 Then
                WidthCI95C.Text = PrintWidth(ESValue, (sumWidth95 / NRep)) '(sumWidth95 / NRep).ToString("f3")
                WidthCI99C.Text = PrintWidth(ESValue, (sumWidth99 / NRep)) '(sumWidth99 / NRep).ToString("f3")
                PowerC.Text = (sumPower / NRep).ToString("f3")
            Else
                ObtainedWidth95C.Text = PrintWidth(ESValue, (sumWidth95 / NRep)) '(sumWidth95 / NRep).ToString("f3")
                ObtainedWidth99C.Text = PrintWidth(ESValue, (sumWidth99 / NRep)) '(sumWidth99 / NRep).ToString("f3")
                ObtainedPowerC.Text = (sumPower / NRep).ToString("f3")
            End If
        End If

        If (checkCovariate.Checked = True And RunNullModel.Checked = True) Then
            Dim result(,) As Double = RunEachSimulatedData(False, NRep, True, ClusterOnlyZ.Checked, NT + NC, NIndiv)
            Dim sumWidth95 As Double = 0, sumWidth99 As Double = 0, sumEst As Double = 0, sumPower As Double = 0
            Dim width95, width99 As Double
            Dim ArrayWidth95(0 To NRep - 1) As Double
            Dim ArrayWidth99(0 To NRep - 1) As Double
            For i = 0 To NRep - 1
                While result(i, 5) = -1
                    WriteMCCovariateCode(NT, NC, NIndiv, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, 1, True, ClusterOnlyZ.Checked, ReverseZ.Checked)
                    RunAnalysis(sim)
                    CorrectedResult = RunEachSimulatedData(False, 1, True, ClusterOnlyZ.Checked, NT + NC, NIndiv)
                    For k = 0 To 5
                        result(i, k) = CorrectedResult(0, k)
                    Next k
                End While
                width95 = result(i, 3) - result(i, 1)
                width99 = result(i, 4) - result(i, 0)
                ArrayWidth95(i) = width95
                ArrayWidth99(i) = width99
                sumEst += result(i, 2)
                sumPower += result(i, 5)
                sumWidth95 += width95
                sumWidth99 += width99
            Next i
            Array.Sort(ArrayWidth95)
            Array.Sort(ArrayWidth99)
            For i = 0 To NRep - 1
                Storage.Width(i, 0) = ArrayWidth95(i)
                Storage.Width(i, 1) = ArrayWidth99(i)
            Next i
            If TabControl1.SelectedIndex = 0 Then
                WidthCI95.Text = PrintWidth(ESValue, (sumWidth95 / NRep)) '(sumWidth95 / NRep).ToString("f3")
                WidthCI99.Text = PrintWidth(ESValue, (sumWidth99 / NRep)) '(sumWidth99 / NRep).ToString("f3")
                Power.Text = (sumPower / NRep).ToString("f3")
            Else
                ObtainedWidth95.Text = PrintWidth(ESValue, (sumWidth95 / NRep)) '(sumWidth95 / NRep).ToString("f3")
                ObtainedWidth99.Text = PrintWidth(ESValue, (sumWidth99 / NRep)) '(sumWidth99 / NRep).ToString("f3")
                ObtainedPower.Text = (sumPower / NRep).ToString("f3")
            End If
        End If

        My.Computer.FileSystem.DeleteFile(Path & "\simulation.inp")
        My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\simulation.out")
        My.Computer.FileSystem.DeleteFile(Path & "\rawlist.dat")
        ClearAllData(NRep)

    End Sub

    Private Function FindingSampleSizeBySimulation(ByVal StartNT As Integer, ByVal StartNC As Integer, ByVal StartingValueNIndiv As Integer, ByVal StartingValuePT As Double)
        'Process.Start("c:\Program Files\Mplus\Mplus.exe c:\CRD\simulation.inp c:\CRD\simulation.out")
        Dim StartNJ As Integer = StartNT + StartNC
        Dim StartPT As Double
        Dim result(0 To 2) As Integer
        If Criteria.PCheckBox.Checked = True Then
            StartPT = PTValue
        Else
            StartPT = StartingValuePT
        End If

        Dim DesiredVar As Double
        If OptionValue = 1 Or OptionValue = 2 Then
            If PowerOrWidthValue = 1 Then
                DesiredVar = FindVarForPower(ESValue, PowerValue)
            Else
                DesiredVar = FindVarForWidth(WidthValue, CILevelValue)
            End If
        End If
        If OptionValue = 1 Or OptionValue = 2 Then
            '    Dim RunNJ(0 To 2) As Integer
            '    If StartNJ > 4 Then
            '        RunNJ(0) = StartNJ - 1
            '        RunNJ(1) = StartNJ
            '        RunNJ(2) = StartNJ + 1
            '    Else
            '        RunNJ(0) = 4
            '        RunNJ(1) = 5
            '        RunNJ(2) = 6
            '    End If
            '    Dim StartNIndiv(0 To 2) As Integer
            '    Dim RunNIndiv(0 To 2) As Integer
            '    Dim RunTNIndiv(0 To 2) As Double
            '    For i = 0 To 2
            '        StartNIndiv(i) = FindNIndiv_PJ_Var(checkCovariate.Checked, StartPT, RunNJ(i), DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
            '        RunNIndiv(i) = FindNIndiv_pJ_Var_MonteCarlo(PowerOrWidthValue, checkCovariate.Checked, RunNJ(i), StartPT, StartNIndiv(i), ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, PowerValue, CILevelValue, WidthValue, ClusterOnlyZ.Checked, ReverseZ.Checked, IsDegreeOfCertaintyValue, DegreeOfCertaintyValue)
            '        RunTNIndiv(i) = CDbl(RunNIndiv(i)) * RunNJ(i)
            '        ProgressBar1.Value = 50 + (250 * (i + 1))
            '        StatusNum.Text = CInt((50 + (250 * (i + 1))) / 10)
            '    Next i
            '    Dim MinPosition As Integer = FindMinimum(RunTNIndiv)
            '    Do While (MinPosition <> 1)
            '        status.Text = "Be patient. We need more algorithm to give the better estimation."
            '        If (MinPosition = 0) Then
            '            If (RunNJ(0) = 4) Then Exit Do
            '            RunNJ(2) = RunNJ(1)
            '            RunNJ(1) = RunNJ(0)
            '            RunNJ(0) = RunNJ(0) - 1
            '            StartNIndiv(2) = StartNIndiv(1)
            '            StartNIndiv(1) = StartNIndiv(0)
            '            StartNIndiv(0) = FindNIndiv_PJ_Var(checkCovariate.Checked, StartPT, RunNJ(0), DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
            '            RunNIndiv(2) = RunNIndiv(1)
            '            RunNIndiv(1) = RunNIndiv(0)
            '            RunNIndiv(0) = FindNIndiv_pJ_Var_MonteCarlo(PowerOrWidthValue, checkCovariate.Checked, RunNJ(0), StartPT, StartNIndiv(0), ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, PowerValue, CILevelValue, WidthValue, ClusterOnlyZ.Checked, ReverseZ.Checked, IsDegreeOfCertaintyValue, DegreeOfCertaintyValue)
            '            RunTNIndiv(2) = RunTNIndiv(1)
            '            RunTNIndiv(1) = RunTNIndiv(0)
            '            RunTNIndiv(0) = CDbl(RunNJ(0)) * RunNIndiv(0)
            '        Else
            '            RunNJ(0) = RunNJ(1)
            '            RunNJ(1) = RunNJ(2)
            '            RunNJ(2) = RunNJ(2) + 1
            '            StartNIndiv(0) = StartNIndiv(1)
            '            StartNIndiv(1) = StartNIndiv(2)
            '            StartNIndiv(2) = FindNIndiv_PJ_Var(checkCovariate.Checked, StartPT, RunNJ(2), DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
            '            RunNIndiv(0) = RunNIndiv(1)
            '            RunNIndiv(1) = RunNIndiv(2)
            '            RunNIndiv(2) = FindNIndiv_pJ_Var_MonteCarlo(PowerOrWidthValue, checkCovariate.Checked, RunNJ(2), StartPT, StartNIndiv(2), ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, PowerValue, CILevelValue, WidthValue, ClusterOnlyZ.Checked, ReverseZ.Checked, IsDegreeOfCertaintyValue, DegreeOfCertaintyValue)
            '            RunTNIndiv(0) = RunTNIndiv(1)
            '            RunTNIndiv(1) = RunTNIndiv(2)
            '            RunTNIndiv(2) = CDbl(RunNJ(2)) * RunNIndiv(2)
            '        End If
            '        MinPosition = FindMinimum(RunTNIndiv)
            '    Loop
            '    result(0) = Math.Round(RunNJ(MinPosition) * StartPT)
            '    result(1) = Math.Round(RunNJ(MinPosition) * (1 - StartPT))
            '    result(2) = RunNIndiv(MinPosition)
            '    If (RunNJ(MinPosition) > (result(0) + result(1))) Then
            '        result(0) += 1
            '    ElseIf (RunNJ(MinPosition) < (result(0) + result(1))) Then
            '        result(1) -= 1
            '    End If
            'ElseIf OptionValue = 2 Then
            If (Criteria.NIndivCheckBox.Checked = True) And (Criteria.NGCheckBox.Checked = True) Then
                ' Change P Only
                result(0) = StartNT
                result(1) = StartNC
                result(2) = StartingValueNIndiv
                'Dim pT() As Double = {StartPT - 0.01, StartPT, StartPT + 1}
                'Dim UsePT As Double
                'Dim RunTNIndiv(0 To UBound(pT)) As Double
                'Dim TCostLess As Boolean = (TGroupCostValue + (NIndivValue * TIndivCostValue)) <= (CGroupCostValue + (NIndivValue * CIndivCostValue))
                'Dim IsPTGreaterThanHalf As Integer
                'If StartPT > 0.5 Then
                '    IsPTGreaterThanHalf = 1
                'ElseIf StartPT < 0.5 Then
                '    IsPTGreaterThanHalf = -1
                'Else
                '    IsPTGreaterThanHalf = 0
                'End If
                'If PowerOrWidthValue = 1 Then
                '    For i = 0 To UBound(pT)
                '        RunTNIndiv(i) = RunSimulationAndReadPower(checkCovariate.Checked, NGValue, pT(i), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep)
                '    Next i
                '    Dim Index As Integer = CompareAllElementWithScalar(RunTNIndiv, PowerValue)
                '    ' PT may > 0.5, <0.5 or = 0.5
                '    If IsPTGreaterThanHalf = 0 Then
                '        If Index = -1 Then
                '            MsgBox("The power specified is infeasible because the number of groups and number of individuals are not enough to make that power.")
                '        ElseIf Index = 1 Then
                '            If TCostLess = True Then
                '                Do
                '                    pT(2) = pT(1)
                '                    pT(1) = pT(0)
                '                    pT(0) = pT(0) - 0.01
                '                    RunTNIndiv(2) = RunTNIndiv(1)
                '                    RunTNIndiv(1) = RunTNIndiv(0)
                '                    RunTNIndiv(0) = RunSimulationAndReadPower(checkCovariate.Checked, NGValue, pT(0), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep)
                '                    Index = CompareAllElementWithScalar(RunTNIndiv, PowerValue)
                '                Loop While Index = 1
                '            Else
                '                Do
                '                    pT(0) = pT(1)
                '                    pT(1) = pT(2)
                '                    pT(2) = pT(2) + 0.01
                '                    RunTNIndiv(0) = RunTNIndiv(1)
                '                    RunTNIndiv(1) = RunTNIndiv(2)
                '                    RunTNIndiv(2) = RunSimulationAndReadPower(checkCovariate.Checked, NGValue, pT(2), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep)
                '                    Index = CompareAllElementWithScalar(RunTNIndiv, PowerValue)
                '                Loop While Index = 1
                '            End If
                '        End If
                '        UsePT = pT(FindCloset(RunTNIndiv, PowerValue, True))
                '    ElseIf IsPTGreaterThanHalf = 1 Then
                '        If Index = 1 Then
                '            Do
                '                pT(0) = pT(1)
                '                pT(1) = pT(2)
                '                pT(2) = pT(2) + 0.01
                '                RunTNIndiv(0) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(2)
                '                RunTNIndiv(2) = RunSimulationAndReadPower(checkCovariate.Checked, NGValue, pT(2), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, PowerValue)
                '            Loop While Index = 1
                '        ElseIf Index = -1 Then
                '            Do
                '                pT(2) = pT(1)
                '                pT(1) = pT(0)
                '                pT(0) = pT(0) - 0.01
                '                RunTNIndiv(2) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(0)
                '                RunTNIndiv(0) = RunSimulationAndReadPower(checkCovariate.Checked, NGValue, pT(0), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, PowerValue)
                '            Loop While Index = 1
                '        End If
                '        UsePT = pT(FindCloset(RunTNIndiv, PowerValue, True))
                '    Else
                '        If Index = -1 Then
                '            Do
                '                pT(0) = pT(1)
                '                pT(1) = pT(2)
                '                pT(2) = pT(2) + 0.01
                '                RunTNIndiv(0) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(2)
                '                RunTNIndiv(2) = RunSimulationAndReadPower(checkCovariate.Checked, NGValue, pT(2), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, PowerValue)
                '            Loop While Index = 1
                '        ElseIf Index = 1 Then
                '            Do
                '                pT(2) = pT(1)
                '                pT(1) = pT(0)
                '                pT(0) = pT(0) - 0.01
                '                RunTNIndiv(2) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(0)
                '                RunTNIndiv(0) = RunSimulationAndReadPower(checkCovariate.Checked, NGValue, pT(0), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, PowerValue)
                '            Loop While Index = 1
                '        End If
                '        UsePT = pT(FindCloset(RunTNIndiv, PowerValue, True))
                '    End If
                'Else
                '    For i = 0 To UBound(pT)
                '        RunTNIndiv(i) = RunSimulationAndReadCIES(checkCovariate.Checked, NGValue, pT(i), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep, CILevelValue)
                '    Next i
                '    Dim Index As Integer = CompareAllElementWithScalar(RunTNIndiv, WidthValue)
                '    ' PT may > 0.5, <0.5 or = 0.5
                '    If IsPTGreaterThanHalf = 0 Then
                '        If Index = 1 Then
                '            MsgBox("The width specified is infeasible because the number of groups and number of individuals are not enough to make that width.")
                '        ElseIf Index = -1 Then
                '            If TCostLess = True Then
                '                Do
                '                    pT(2) = pT(1)
                '                    pT(1) = pT(0)
                '                    pT(0) = pT(0) - 0.01
                '                    RunTNIndiv(2) = RunTNIndiv(1)
                '                    RunTNIndiv(1) = RunTNIndiv(0)
                '                    RunTNIndiv(0) = RunSimulationAndReadCIES(checkCovariate.Checked, NGValue, pT(0), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep, CILevelValue)
                '                    Index = CompareAllElementWithScalar(RunTNIndiv, WidthValue)
                '                Loop While Index = -1
                '            Else
                '                Do
                '                    pT(0) = pT(1)
                '                    pT(1) = pT(2)
                '                    pT(2) = pT(2) + 0.01
                '                    RunTNIndiv(0) = RunTNIndiv(1)
                '                    RunTNIndiv(1) = RunTNIndiv(2)
                '                    RunTNIndiv(2) = RunSimulationAndReadCIES(checkCovariate.Checked, NGValue, pT(2), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep, CILevelValue)
                '                    Index = CompareAllElementWithScalar(RunTNIndiv, WidthValue)
                '                Loop While Index = -1
                '            End If
                '        End If
                '        UsePT = pT(FindCloset(RunTNIndiv, WidthValue, False))
                '    ElseIf IsPTGreaterThanHalf = 1 Then
                '        If Index = 1 Then 'Width So Wide
                '            Do
                '                pT(2) = pT(1)
                '                pT(1) = pT(0)
                '                pT(0) = pT(0) - 0.01
                '                RunTNIndiv(2) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(0)
                '                RunTNIndiv(0) = RunSimulationAndReadCIES(checkCovariate.Checked, NGValue, pT(0), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep, CILevelValue)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, WidthValue)
                '            Loop While Index = 1
                '        ElseIf Index = -1 Then 'Width So Narrow
                '            Do
                '                pT(0) = pT(1)
                '                pT(1) = pT(2)
                '                pT(2) = pT(2) + 0.01
                '                RunTNIndiv(0) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(2)
                '                RunTNIndiv(2) = RunSimulationAndReadCIES(checkCovariate.Checked, NGValue, pT(2), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep, CILevelValue)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, WidthValue)
                '            Loop While Index = 1
                '        End If
                '        UsePT = pT(FindCloset(RunTNIndiv, WidthValue, False))
                '    Else
                '        If Index = 1 Then 'Width So Wide'
                '            Do 'Move to the right'
                '                pT(0) = pT(1)
                '                pT(1) = pT(2)
                '                pT(2) = pT(2) + 0.01
                '                RunTNIndiv(0) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(2)
                '                RunTNIndiv(2) = RunSimulationAndReadCIES(checkCovariate.Checked, NGValue, pT(2), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep, CILevelValue)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, WidthValue)
                '            Loop While Index = 1
                '        ElseIf Index = -1 Then 'Width So Wide'
                '            Do 'Move to the left'
                '                pT(2) = pT(1)
                '                pT(1) = pT(0)
                '                pT(0) = pT(0) - 0.01
                '                RunTNIndiv(2) = RunTNIndiv(1)
                '                RunTNIndiv(1) = RunTNIndiv(0)
                '                RunTNIndiv(0) = RunSimulationAndReadCIES(checkCovariate.Checked, NGValue, pT(0), NIndivValue, ICCYValue, ESValue, ICCZValue, TEZValue, CEZValue, NRep, CILevelValue)
                '                Index = CompareAllElementWithScalar(RunTNIndiv, WidthValue)
                '            Loop While Index = 1
                '        End If
                '        UsePT = pT(FindCloset(RunTNIndiv, WidthValue, False))
                '    End If
                'End If
                'result(0) = Math.Round(NGValue * UsePT)
                'result(1) = Math.Round(NGValue * (1 - UsePT))
                'result(2) = NIndivValue
                'If (NGValue > (result(0) + result(1))) Then
                '    If (TCostLess = True) Then
                '        result(0) += 1
                '    Else
                '        result(1) += 1
                '    End If
                'ElseIf (NGValue < (result(0) + result(1))) Then
                '    If (TCostLess = True) Then
                '        result(1) -= 1
                '    Else
                '        result(0) -= 1
                '    End If
                'End If
            ElseIf (Criteria.NGCheckBox.Checked = True) Then
                'Change NIndiv Only
                Dim StartNIndiv As Integer = FindNIndiv_PJ_Var(checkCovariate.Checked, StartPT, NGValue, DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                Dim RunNIndiv As Integer = FindNIndiv_pJ_Var_MonteCarlo(PowerOrWidthValue, checkCovariate.Checked, NGValue, StartPT, StartNIndiv, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, PowerValue, CILevelValue, WidthValue, ClusterOnlyZ.Checked, ReverseZ.Checked, IsDegreeOfCertaintyValue, DegreeOfCertaintyValue)
                result(0) = Math.Round(NGValue * StartPT)
                result(1) = Math.Round(NGValue * (1 - StartPT))
                result(2) = RunNIndiv
                Dim TCostLess As Boolean = (TGroupCostValue + (RunNIndiv * TIndivCostValue)) <= (CGroupCostValue + (RunNIndiv * CIndivCostValue))
                If (NGValue > (result(0) + result(1))) Then
                    If (TCostLess = True) Then
                        result(0) += 1
                    Else
                        result(1) += 1
                    End If
                ElseIf (NGValue < (result(0) + result(1))) Then
                    If (TCostLess = True) Then
                        result(1) -= 1
                    Else
                        result(0) -= 1
                    End If
                End If
            ElseIf (Criteria.NIndivCheckBox.Checked = True) Then
                'Change NJ Only
                Dim RunNJ(0 To 2) As Integer
                If StartNJ > 4 Then
                    RunNJ(0) = StartNJ - 1
                    RunNJ(1) = StartNJ
                    RunNJ(2) = StartNJ + 1
                Else
                    RunNJ(0) = 4
                    RunNJ(1) = 5
                    RunNJ(2) = 6
                End If
                Dim UseNJ As Integer
                Dim RunTNIndiv(0 To 2) As Double
                Dim TCostLess As Boolean = (TGroupCostValue + (NIndivValue * TIndivCostValue)) <= (CGroupCostValue + (NIndivValue * CIndivCostValue))
                If PowerOrWidthValue = 1 Then
                    For i = 0 To 2
                        RunTNIndiv(i) = RunSimulationAndReadPower(checkCovariate.Checked, RunNJ(i), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, ClusterOnlyZ.Checked, ReverseZ.Checked)
                        ProgressBar1.Value = 50 + (250 * (i + 1))
                        StatusNum.Text = CInt((50 + (250 * (i + 1))) / 10)
                    Next
                    Dim Index As Integer = FindIndexPass(RunTNIndiv, PowerValue, True)
                    If Index = -1 Then
                        Do
                            status.Text = "Be patient. We need more algorithm to give the better estimation."
                            RunNJ(0) = RunNJ(1)
                            RunNJ(1) = RunNJ(2)
                            RunNJ(2) = RunNJ(2) + 1
                            RunTNIndiv(0) = RunTNIndiv(1)
                            RunTNIndiv(1) = RunTNIndiv(2)
                            RunTNIndiv(2) = RunSimulationAndReadPower(checkCovariate.Checked, RunNJ(2), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, ClusterOnlyZ.Checked, ReverseZ.Checked)
                            Index = FindIndexPass(RunTNIndiv, PowerValue, True)
                        Loop While Index = -1
                    ElseIf Index = 0 Then
                        Do
                            If (RunNJ(0) = 4) Then Exit Do
                            status.Text = "Be patient. We need more algorithm to give the better estimation."
                            RunNJ(2) = RunNJ(1)
                            RunNJ(1) = RunNJ(0)
                            RunNJ(0) = RunNJ(0) - 1
                            RunTNIndiv(2) = RunTNIndiv(1)
                            RunTNIndiv(1) = RunTNIndiv(0)
                            RunTNIndiv(0) = RunSimulationAndReadPower(checkCovariate.Checked, RunNJ(0), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, ClusterOnlyZ.Checked, ReverseZ.Checked)
                            Index = FindIndexPass(RunTNIndiv, PowerValue, True)
                        Loop While Index = 0
                    End If
                    UseNJ = RunNJ(Index)
                Else
                    If IsDegreeOfCertaintyValue = True Then
                        For i = 0 To 2
                            '(checkCovariate.Checked, RunNJ(i), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, ClusterOnlyZ.Checked, ReverseZ.Checked)
                            RunTNIndiv(i) = RunSimulationAndReadDegreeOfCertainty(checkCovariate.Checked, RunNJ(i), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, CILevelValue, ClusterOnlyZ.Checked, ReverseZ.Checked, WidthValue)
                            ProgressBar1.Value = 50 + (250 * (i + 1))
                            StatusNum.Text = CInt((50 + (250 * (i + 1))) / 10)
                        Next i
                        Dim Index As Integer = FindIndexPass(RunTNIndiv, DegreeOfCertaintyValue, True)
                        If Index = -1 Then
                            Do
                                status.Text = "Be patient. We need more algorithm to give the better estimation."
                                RunNJ(0) = RunNJ(1)
                                RunNJ(1) = RunNJ(2)
                                RunNJ(2) = RunNJ(2) + 1
                                RunTNIndiv(0) = RunTNIndiv(1)
                                RunTNIndiv(1) = RunTNIndiv(2)
                                '(checkCovariate.Checked, RunNJ(2), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, ClusterOnlyZ.Checked, ReverseZ.Checked)
                                RunTNIndiv(2) = RunSimulationAndReadDegreeOfCertainty(checkCovariate.Checked, RunNJ(2), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, CILevelValue, ClusterOnlyZ.Checked, ReverseZ.Checked, WidthValue)
                                Index = FindIndexPass(RunTNIndiv, DegreeOfCertaintyValue, True)
                            Loop While Index = -1
                        ElseIf Index = 0 Then
                            Do
                                If (RunNJ(0) = 4) Then Exit Do
                                status.Text = "Be patient. We need more algorithm to give the better estimation."
                                RunNJ(2) = RunNJ(1)
                                RunNJ(1) = RunNJ(0)
                                RunNJ(0) = RunNJ(0) - 1
                                RunTNIndiv(2) = RunTNIndiv(1)
                                RunTNIndiv(1) = RunTNIndiv(0)
                                RunTNIndiv(0) = RunSimulationAndReadDegreeOfCertainty(checkCovariate.Checked, RunNJ(0), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, CILevelValue, ClusterOnlyZ.Checked, ReverseZ.Checked, WidthValue)
                                Index = FindIndexPass(RunTNIndiv, DegreeOfCertaintyValue, True)
                            Loop While Index = 0
                        End If
                        UseNJ = RunNJ(Index)
                    Else
                        For i = 0 To 2
                            RunTNIndiv(i) = RunSimulationAndReadCIES(checkCovariate.Checked, RunNJ(i), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, CILevelValue, ClusterOnlyZ.Checked, ReverseZ.Checked)
                            ProgressBar1.Value = 50 + (250 * (i + 1))
                            StatusNum.Text = CInt((50 + (250 * (i + 1))) / 10)
                        Next
                        Dim Index As Integer = FindIndexPass(RunTNIndiv, WidthValue, False)
                        If Index = -1 Then
                            Do
                                status.Text = "Be patient. We need more algorithm to give the better estimation."
                                RunNJ(0) = RunNJ(1)
                                RunNJ(1) = RunNJ(2)
                                RunNJ(2) = RunNJ(2) + 1
                                RunTNIndiv(0) = RunTNIndiv(1)
                                RunTNIndiv(1) = RunTNIndiv(2)
                                RunTNIndiv(2) = RunSimulationAndReadCIES(checkCovariate.Checked, RunNJ(2), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, CILevelValue, ClusterOnlyZ.Checked, ReverseZ.Checked)
                                Index = FindIndexPass(RunTNIndiv, WidthValue, False)
                            Loop While Index = -1
                        ElseIf Index = 0 Then
                            Do
                                If (RunNJ(0) = 4) Then Exit Do
                                status.Text = "Be patient. We need more algorithm to give the better estimation."
                                RunNJ(2) = RunNJ(1)
                                RunNJ(1) = RunNJ(0)
                                RunNJ(0) = RunNJ(0) - 1
                                RunTNIndiv(2) = RunTNIndiv(1)
                                RunTNIndiv(1) = RunTNIndiv(0)
                                RunTNIndiv(0) = RunSimulationAndReadCIES(checkCovariate.Checked, RunNJ(0), StartPT, NIndivValue, ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, CILevelValue, ClusterOnlyZ.Checked, ReverseZ.Checked)
                                Index = FindIndexPass(RunTNIndiv, WidthValue, False)
                            Loop While Index = 0
                        End If
                        UseNJ = RunNJ(Index)
                    End If
                End If
                result(0) = Math.Round(UseNJ * StartPT)
                result(1) = Math.Round(UseNJ * (1 - StartPT))
                result(2) = NIndivValue
                If (UseNJ > (result(0) + result(1))) Then
                    If (TCostLess = True) Then
                        result(0) += 1
                    Else
                        result(1) += 1
                    End If
                ElseIf (UseNJ < (result(0) + result(1))) Then
                    If (TCostLess = True) Then
                        result(1) -= 1
                    Else
                        result(0) -= 1
                    End If
                End If
            Else
                'Change Both NIndiv and NJ
                Dim RunNJ(0 To 2) As Integer
                If StartNJ > 4 Then
                    RunNJ(0) = StartNJ - 1
                    RunNJ(1) = StartNJ
                    RunNJ(2) = StartNJ + 1
                Else
                    RunNJ(0) = 4
                    RunNJ(1) = 5
                    RunNJ(2) = 6
                End If
                Dim StartNIndiv(0 To 2) As Integer
                Dim RunNIndiv(0 To 2) As Integer
                Dim RunTNIndiv(0 To 2) As Double
                For i = 0 To 2
                    StartNIndiv(i) = FindNIndiv_PJ_Var(checkCovariate.Checked, StartPT, RunNJ(i), DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                    RunNIndiv(i) = FindNIndiv_pJ_Var_MonteCarlo(PowerOrWidthValue, checkCovariate.Checked, RunNJ(i), StartPT, StartNIndiv(i), ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, PowerValue, CILevelValue, WidthValue, ClusterOnlyZ.Checked, ReverseZ.Checked, IsDegreeOfCertaintyValue, DegreeOfCertaintyValue)
                    RunTNIndiv(i) = FindTotalCostExact(RunNJ(i), StartPT, RunNIndiv(i), TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue)
                    ProgressBar1.Value = 50 + (250 * (i + 1))
                    StatusNum.Text = CInt((50 + (250 * (i + 1))) / 10)
                Next
                Dim MinPosition As Integer = FindMinimum(RunTNIndiv)
                Do While (MinPosition <> 1)
                    If (MinPosition = 0) Then
                        If (RunNJ(0) = 4) Then Exit Do
                        status.Text = "Be patient. We need more algorithm to give the better estimation."
                        RunNJ(2) = RunNJ(1)
                        RunNJ(1) = RunNJ(0)
                        RunNJ(0) = RunNJ(0) - 1
                        StartNIndiv(2) = StartNIndiv(1)
                        StartNIndiv(1) = StartNIndiv(0)
                        StartNIndiv(0) = FindNIndiv_PJ_Var(checkCovariate.Checked, StartPT, RunNJ(0), DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                        RunNIndiv(2) = RunNIndiv(1)
                        RunNIndiv(1) = RunNIndiv(0)
                        RunNIndiv(0) = FindNIndiv_pJ_Var_MonteCarlo(PowerOrWidthValue, checkCovariate.Checked, RunNJ(0), StartPT, StartNIndiv(0), ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, PowerValue, CILevelValue, WidthValue, ClusterOnlyZ.Checked, ReverseZ.Checked, IsDegreeOfCertaintyValue, DegreeOfCertaintyValue)
                        RunTNIndiv(2) = RunTNIndiv(1)
                        RunTNIndiv(1) = RunTNIndiv(0)
                        RunTNIndiv(0) = FindTotalCostExact(RunNJ(0), StartPT, RunNIndiv(0), TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue)
                    Else
                        status.Text = "Be patient. We need more algorithm to give the better estimation."
                        RunNJ(0) = RunNJ(1)
                        RunNJ(1) = RunNJ(2)
                        RunNJ(2) = RunNJ(2) + 1
                        StartNIndiv(0) = StartNIndiv(1)
                        StartNIndiv(1) = StartNIndiv(2)
                        StartNIndiv(2) = FindNIndiv_PJ_Var(checkCovariate.Checked, StartPT, RunNJ(2), DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                        RunNIndiv(0) = RunNIndiv(1)
                        RunNIndiv(1) = RunNIndiv(2)
                        RunNIndiv(2) = FindNIndiv_pJ_Var_MonteCarlo(PowerOrWidthValue, checkCovariate.Checked, RunNJ(2), StartPT, StartNIndiv(2), ICCYValue, ESValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NRep, PowerValue, CILevelValue, WidthValue, ClusterOnlyZ.Checked, ReverseZ.Checked, IsDegreeOfCertaintyValue, DegreeOfCertaintyValue)
                        RunTNIndiv(0) = RunTNIndiv(1)
                        RunTNIndiv(1) = RunTNIndiv(2)
                        RunTNIndiv(2) = FindTotalCostExact(RunNJ(0), StartPT, RunNIndiv(0), TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue)
                    End If
                    MinPosition = FindMinimum(RunTNIndiv)
                Loop
                result(0) = Math.Round(RunNJ(MinPosition) * StartPT)
                result(1) = Math.Round(RunNJ(MinPosition) * (1 - StartPT))
                result(2) = RunNIndiv(MinPosition)
                If (RunNJ(MinPosition) > (result(0) + result(1))) Then
                    result(0) += 1
                ElseIf (RunNJ(MinPosition) < (result(0) + result(1))) Then
                    result(1) -= 1
                End If
            End If
        Else
            result(0) = StartNT
            result(1) = StartNC
            result(2) = StartingValueNIndiv
        End If
        Return result
    End Function

    Private Sub checkCovariate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles checkCovariate.CheckedChanged
        CheckCovariateChanged()
    End Sub

    Public Sub CheckCovariateChanged()
        If (checkCovariate.Checked = False) Then
            covariateChecked(False)
            nullOutputChecked(True)
        Else
            covariateChecked(True)
            If (RunNullModel.Checked = False) Then nullOutputChecked(False)
        End If
    End Sub

    Private Sub covariateChecked(ByVal activate As Boolean)
        ICCZ.Enabled = activate
        ICCZ2.Enabled = activate
        ICCZLabel.Enabled = activate
        ICCZLabel2.Enabled = activate
        RunNullModel.Enabled = activate
        CovariateGroupBox1.Enabled = activate
        CovariateGroupBox2.Enabled = activate
        PowerOutputLabelC.Enabled = activate
        PowerOutputLabel2C.Enabled = activate
        Width95OutputLabelC.Enabled = activate
        Width99OutputLabelC.Enabled = activate
        Width95OutputLabel2C.Enabled = activate
        OutputC.Enabled = activate
        OutputC2.Enabled = activate
        ClusterOnlyZ.Enabled = activate
        IndivOnlyZ.Enabled = activate
        If activate And (ClusterOnlyZ.Checked) Then
            R2IndivZ.Enabled = False
            R2IndivZ2.Enabled = False
            R2IndivdualZLabel.Enabled = False
            R2IndivZLabel2.Enabled = False
            ReverseZ.Enabled = False
            ICCZ.Enabled = False
            ICCZ2.Enabled = False
            ICCZLabel.Enabled = False
            ICCZLabel2.Enabled = False
            R2ClusterZ.Enabled = True
            R2ClusterZ2.Enabled = True
            R2ClusterZLabel.Enabled = True
            R2ClusterZLabel2.Enabled = True
            R2IndivZ.Text = "0.0"
            R2IndivZ2.Text = "0.0"
            ICCZ.Text = "0.0"
            ICCZ2.Text = "0.0"
        ElseIf activate And (IndivOnlyZ.Checked) Then
            R2ClusterZ.Enabled = False
            R2ClusterZ2.Enabled = False
            R2ClusterZLabel.Enabled = False
            R2ClusterZLabel2.Enabled = False
            ReverseZ.Enabled = False
            ICCZ.Enabled = False
            ICCZ2.Enabled = False
            ICCZLabel.Enabled = False
            ICCZLabel2.Enabled = False
            R2IndivZ.Enabled = True
            R2IndivZ2.Enabled = True
            R2IndivdualZLabel.Enabled = True
            R2IndivZLabel2.Enabled = True
            R2ClusterZ.Text = "0.0"
            R2ClusterZ2.Text = "0.0"
            ICCZ.Text = "0.0"
            ICCZ2.Text = "0.0"
        Else
            R2IndivZ.Enabled = activate
            R2IndivZ2.Enabled = activate
            R2IndivdualZLabel.Enabled = activate
            R2IndivZLabel2.Enabled = activate
            ReverseZ.Enabled = activate
            R2ClusterZLabel.Enabled = activate
            R2ClusterZLabel2.Enabled = activate
            R2ClusterZ.Enabled = activate
            R2ClusterZ2.Enabled = activate
            ICCZ.Enabled = activate
            ICCZ2.Enabled = activate
            ICCZLabel.Enabled = activate
            ICCZLabel2.Enabled = activate
        End If
    End Sub

    Private Sub nullOutputChecked(ByVal activate As Boolean)
        PowerOutputLabel.Enabled = activate
        PowerOutputLabel2.Enabled = activate
        Width95OutputLabel.Enabled = activate
        Width99OutputLabel.Enabled = activate
        Width95OutputLabel2.Enabled = activate
        OutputWC.Enabled = activate
        OutputWC2.Enabled = activate
    End Sub

    Private Function ValidateInputs()
        Dim Result As Boolean = True
        If (TabControl1.SelectedIndex = 0) Then
            If (Not IsIntegerLBound(NTGroups, 2, "Number of treatment groups")) Then
                Result = False
            End If
            If (Not IsIntegerLBound(NCGroups, 2, "Number of control groups")) Then
                Result = False
            End If

            If (Not IsIntegerLBound(NIndiv, 2, "Number of individuals in each group")) Then
                Result = False
            End If

            If (Not IsDoubleInRange(ICCY, 0.0, 1.0, "Intraclass correlation of Y")) Then
                Result = False
            End If

            If (Not IsDoubleInRange(ES, -3.0, 3.0, "Effect size")) Then
                Result = False
            End If

            If (checkCovariate.Checked = True) Then
                If (Not IsDoubleInRange(ICCZ, 0.0, 1.0, "Intraclass correlation of Z")) Then
                    Result = False
                End If

                If (Not IsDoubleInRange(R2ClusterZ, 0.0, 1.0, "Cluster-Level Error Variance Explained")) Then
                    Result = False
                End If

                If (Not IsDoubleInRange(R2IndivZ, 0.0, 1.0, "Individual-Level Error Variance Explained")) Then
                    Result = False
                Else
                    If (CDbl(R2IndivZ.Text) >= 0 And CDbl(R2IndivZ.Text) < 0.01) And Not ClusterOnlyZ.Checked Then
                        MsgBox("The individual-level error variance explained is too close to 1. It should be defined as cluster-level covariate by checking the checkbox on the bottom left.")
                        Result = False
                    End If
                End If

            End If
        ElseIf (TabControl1.SelectedIndex = 1) Then
            If (Not IsDoubleInRange(ICCY2, 0.0, 1.0, "Intraclass correlation of Y")) Then
                Result = False
            End If

            If (Not IsDoubleInRange(ES2, -3.0, 3.0, "Effect size")) Then
                Result = False
            End If

            If (checkCovariate.Checked = True) Then
                If (Not IsDoubleInRange(ICCZ2, 0.0, 1.0, "Intraclass correlation of Z")) Then
                    Result = False
                End If

                If (Not IsDoubleInRange(R2ClusterZ2, 0.0, 1.0, "Cluster-level Error Variance Explained")) Then
                    Result = False
                End If

                If (Not IsDoubleInRange(R2IndivZ2, 0.0, 1.0, "Individual-level Error Variance Explained")) Then
                    Result = False
                Else
                    If (CDbl(R2IndivZ2.Text) >= 0 And CDbl(R2IndivZ2.Text) < 0.01) And Not ClusterOnlyZ.Checked Then
                        MsgBox("The individual-level error variance explained is too close to 1. It should be defined as cluster-level covariate by checking the checkbox on the bottom left.")
                        Result = False
                    End If
                End If
            End If

            If IsCriteriaDefined = False Then
                Result = False
                MsgBox("Please specify input criteria")
            End If
        End If
        Return Result
    End Function

    Private Function CheckPossibleParameter()
        Dim result As Boolean = True
        'If (TabControl1.SelectedIndex = 0) Then
        '    Dim tempICCY As Double = CDbl(ICCY.Text)
        '    Dim ResidualYBetween, ResidualYWithin As Double
        '    If (checkCovariate.Checked = True) Then
        '        Dim totalEffect, contextualEffect, tempICCZ As Double
        '        Dim dfB As Integer = CInt(NTGroups.Text) + CInt(NCGroups.Text) - 1
        '        Dim dfW As Integer = (CInt(NTGroups.Text) + CInt(NCGroups.Text)) * (CInt(NIndiv.Text) - 1)
        '        If TabControl1.SelectedIndex = 0 Then
        '            totalEffect = CDbl(R2ClusterZ.Text)
        '            contextualEffect = CDbl(R2IndivZ.Text)
        '            tempICCZ = CDbl(ICCZ.Text)
        '        Else
        '            totalEffect = CDbl(R2ClusterZ2.Text)
        '            contextualEffect = CDbl(R2IndivZ2.Text)
        '            tempICCZ = CDbl(ICCZ2.Text)
        '        End If
        '        Dim tauZ As Double = (dfB + dfW) / (dfB + (dfW * (1 - tempICCZ) / tempICCZ))
        '        Dim sigmaZ As Double = tauZ * (1 - tempICCZ) / tempICCZ
        '        Dim SSBZ As Double = dfB * tauZ
        '        Dim SSWZ As Double = dfW * sigmaZ
        '        Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
        '        Dim BetaB As Double = totalEffect + ((1 - etaZ) * contextualEffect)
        '        Dim BetaW As Double = BetaB - contextualEffect
        '        If (BetaB > 1.0) Or (BetaB < -1.0) Then
        '            result = False
        '        End If
        '        If (BetaW > 1.0) Or (BetaW < -1.0) Then
        '            result = False
        '        End If
        '        If (result = False) Then
        '            MsgBox("Please enter new combination of total effect of Y on Z and contextual effect of Y on Z" & vbCrLf & _
        '                   "Since the effect between group = " & BetaB.ToString("f3") & " and the effect within group = " & BetaW.ToString("f3"))
        '        End If
        '        ResidualYBetween = (tempICCY / (1 - tempICCY)) - (BetaB * BetaB * tauZ)
        '        ResidualYWithin = 1 - (BetaW * BetaW * sigmaZ)
        '        If (ResidualYBetween < 0) Then
        '            result = False
        '            MsgBox("Please decrease effect size of Y on X, increase intraclass correlation of Y, or decrease the effect of Y on covariate because the residual of Y between group is less than 0.")
        '        End If
        '        If (ResidualYWithin < 0) Then
        '            result = False
        '            MsgBox("Please decrease the effect of Y on covariate because the residual of Y within group is less than 0.")
        '        End If
        '    End If
        'End If ' We cannot check whether the cez and tez is out of bound in tabControl = 1, since no NT, NC, NINDIV Values

        'If (checkCovariate.Checked = True) Then
        '    Dim TempICCY, TempR2ClusterZ, ResidualCluster As Double
        '    If (TabControl1.SelectedIndex = 0) Then
        '        TempICCY = CDbl(ICCY.Text)
        '        TempR2ClusterZ = CDbl(R2ClusterZ.Text)
        '    Else
        '        TempICCY = CDbl(ICCY2.Text)
        '        TempR2ClusterZ = CDbl(R2ClusterZ2.Text)
        '    End If
        '    ResidualCluster = (TempICCY / (1 - TempICCY)) - TempR2ClusterZ
        '    If (ResidualCluster < 0) Then
        '        result = False
        '        MsgBox("Please increase intraclass correlation of Y, or decrease the proportion of cluster level explained by the covariate because the residual of Y between clusters is less than 0.")
        '    End If
        'End If
        Return result
    End Function
    Private Sub WriteInput()
        If (TabControl1.SelectedIndex = 0) Then
            NTValue = CInt(NTGroups.Text)
            NCValue = CInt(NCGroups.Text)
            NIndivValue = CInt(NIndiv.Text)
            ICCYValue = CDbl(ICCY.Text)
            ESValue = CDbl(ES.Text)
            If (checkCovariate.Checked = True) Then
                ICCZValue = CDbl(ICCZ.Text)
                R2ClusterZValue = CDbl(R2ClusterZ.Text)
                R2IndivZValue = CDbl(R2IndivZ.Text)
            End If
        ElseIf (TabControl1.SelectedIndex = 1) Then
            ICCYValue = CDbl(ICCY2.Text)
            ESValue = CDbl(ES2.Text)
            If (checkCovariate.Checked = True) Then
                ICCZValue = CDbl(ICCZ2.Text)
                R2ClusterZValue = CDbl(R2ClusterZ2.Text)
                R2IndivZValue = CDbl(R2IndivZ2.Text)
                Storage.MinIndiv = 3
            Else
                Storage.MinIndiv = 2
            End If
        End If
    End Sub

    Private Function FindStartingValue()
        Dim result(0 To 3) As Double
        If OptionValue = 1 Or OptionValue = 2 Then
            '    Dim DesiredVar As Double

            '    If (PowerOrWidthValue = 1) Then
            '        DesiredVar = FindVarForPower(ESValue, PowerValue)
            '    ElseIf (PowerOrWidthValue = 2) Then
            '        DesiredVar = FindVarForWidth(WidthValue, CILevelValue)
            '    End If
            '    If checkCovariate.Checked = True Then
            '        If (Criteria.PCheckBox.Checked = True) Then
            '            result = FindOption1(True, DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue, PTValue)
            '        Else
            '            result = FindOption1(True, DesiredVar, ESValue, ICCYValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
            '        End If
            '    Else
            '        If (Criteria.PCheckBox.Checked = True) Then
            '            result = FindOption1(False, DesiredVar, ESValue, ICCYValue, pT2:=PTValue)
            '        Else
            '            result = FindOption1(False, DesiredVar, ESValue, ICCYValue)
            '        End If
            '    End If
            '    'Start n and find J until reaching the specified power or width (pT = .5 unless otherwise specified)
            '    'Find whether decrease or increase n will increase n*J when achieve power or CIES
            '    'Finalize
            'ElseIf OptionValue = 2 Then
            Dim DesiredVar As Double
            If (PowerOrWidthValue = 1) Then
                DesiredVar = FindVarForPower(ESValue, PowerValue)
            ElseIf (PowerOrWidthValue = 2) Then
                DesiredVar = FindVarForWidth(WidthValue, CILevelValue)
            End If
            If checkCovariate.Checked = True Then
                If (Criteria.PCheckBox.Checked = True) Then
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        'P & NINDIV
                        result = FindOption2(True, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, pT2:=PTValue, NIndiv2:=NIndivValue)
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'P & NG
                            result = FindOption2(True, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, pT2:=PTValue, NJ2:=NGValue)
                        Else
                            'P only
                            result = FindOption2(True, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, pT2:=PTValue)
                        End If
                    End If
                Else
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NIndiv & NG
                            result = FindOption2(True, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NIndiv2:=NIndivValue, NJ2:=NGValue)
                        Else
                            'NIndiv Only
                            result = FindOption2(True, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NIndiv2:=NIndivValue)
                        End If
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NG Only
                            result = FindOption2(True, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NJ2:=NGValue)
                        Else
                            'None specified
                            result = FindOption2(True, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                        End If
                    End If
                End If
            Else 'NO COVARIATE
                If (Criteria.PCheckBox.Checked = True) Then
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        'P & NINDIV
                        result = FindOption2(False, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, pT2:=PTValue, NIndiv2:=NIndivValue)
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'P & NG
                            result = FindOption2(False, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, pT2:=PTValue, NJ2:=NGValue)
                        Else
                            'P only
                            result = FindOption2(False, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, pT2:=PTValue)
                        End If
                    End If
                Else
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NIndiv & NG
                            result = FindOption2(False, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, NJ2:=NGValue, NIndiv2:=NIndivValue)
                        Else
                            'NIndiv Only
                            result = FindOption2(False, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, NIndiv2:=NIndivValue)
                        End If
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NG Only
                            result = FindOption2(False, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, NJ2:=NGValue)
                        Else
                            'None specified
                            result = FindOption2(False, DesiredVar, ESValue, ICCYValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue)
                        End If
                    End If
                End If
            End If
        ElseIf OptionValue = 3 Then
            If checkCovariate.Checked = True Then
                If (Criteria.PCheckBox.Checked = True) Then
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        'P & NINDIV
                        result = FindOption3(True, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, pT2:=PTValue, NIndiv2:=NIndivValue)
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'P & NG
                            result = FindOption3(True, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, pT2:=PTValue, NJ2:=NGValue)
                        Else
                            'P only
                            result = FindOption3(True, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, pT2:=PTValue)
                        End If
                    End If
                Else
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NIndiv & NG
                            result = FindOption3(True, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NIndiv2:=NIndivValue, NJ2:=NGValue)
                        Else
                            'NIndiv Only
                            result = FindOption3(True, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NIndiv2:=NIndivValue)
                        End If
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NG Only
                            result = FindOption3(True, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue, NJ2:=NGValue)
                        Else
                            'None specified
                            result = FindOption3(True, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, ICCZValue, R2ClusterZValue, R2IndivZValue)
                        End If
                    End If
                End If
            Else 'NO COVARIATE
                If (Criteria.PCheckBox.Checked = True) Then
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        'P & NINDIV
                        result = FindOption3(False, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, pT2:=PTValue, NIndiv2:=NIndivValue)
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'P & NG
                            result = FindOption3(False, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, pT2:=PTValue, NJ2:=NGValue)
                        Else
                            'P only
                            result = FindOption3(False, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, pT2:=PTValue)
                        End If
                    End If
                Else
                    If (Criteria.NIndivCheckBox.Checked = True) Then
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NIndiv & NG
                            result = FindOption3(False, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, NJ2:=NGValue, NIndiv2:=NIndivValue)
                        Else
                            'NIndiv Only
                            result = FindOption3(False, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, NIndiv2:=NIndivValue)
                        End If
                    Else
                        If (Criteria.NGCheckBox.Checked = True) Then
                            'NG Only
                            result = FindOption3(False, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue, NJ2:=NGValue)
                        Else
                            'None specified
                            result = FindOption3(False, ESValue, ICCYValue, TotalCostValue, TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue)
                        End If
                    End If
                End If
            End If
        End If
        Return result
    End Function

    Private Sub RunNullModel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunNullModel.CheckedChanged
        RunNullModelChanged()
    End Sub

    Public Sub RunNullModelChanged()
        nullOutputChecked(RunNullModel.Checked)
    End Sub

    Private Sub ExitProgramToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitProgramToolStripMenuItem.Click
        End
    End Sub

    Private Sub ShowSummaryMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowSummaryMenuItem.Click
        'LoadFile()
        'Dim dTaskID As Double, path As String, file As String
        'path = "C:\WINDOWS\notepad.exe"
        'file = My.Application.Info.DirectoryPath & "output.txt"
        'dTaskID = Shell(path + " " + file, vbNormalFocus)

        'Dim StartInfo As New ProcessStartInfo("notepad.exe")
        'StartInfo.Arguments = My.Application.Info.DirectoryPath & "\output.txt"
        'Process.Start(StartInfo)

        Dim process As Process
        process = New Process()
        process.StartInfo.FileName = "notepad"
        process.StartInfo.Arguments = My.Application.Info.DirectoryPath & "\output.txt"
        process.Start()

    End Sub

    Private Sub Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Reset.Click
        Storage.Reset()
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        Storage.Reset()
    End Sub

    Private Sub ClusterOnlyZ_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClusterOnlyZ.CheckedChanged
        If ClusterOnlyZ.Checked = True Then
            If (TabControl1.SelectedIndex = 0) Then
                R2IndivZ.Text = "0.0"
                ICCZ.Text = "1.0"
            Else
                R2IndivZ2.Text = "0.0"
                ICCZ2.Text = "1.0"
            End If
            IndivOnlyZ.Checked = False
            ICCZ.Enabled = False
            ICCZ2.Enabled = False
            ICCZLabel.Enabled = False
            ICCZLabel2.Enabled = False
            R2ClusterZ.Enabled = True
            R2ClusterZ2.Enabled = True
            R2ClusterZLabel.Enabled = True
            R2ClusterZLabel2.Enabled = True
        Else
            If (TabControl1.SelectedIndex = 0) Then
                R2IndivZ.Text = ""
            Else
                R2IndivZ2.Text = ""
            End If
        End If
        Dim activate As Boolean = Not ClusterOnlyZ.Checked
        R2IndivZ.Enabled = activate
        R2IndivZ2.Enabled = activate
        R2IndivdualZLabel.Enabled = activate
        R2IndivZLabel2.Enabled = activate


        If ClusterOnlyZ.Checked And (R2ClusterZ.Text = "0.0") Then R2ClusterZ.Text = ""
        If ClusterOnlyZ.Checked And (R2ClusterZ2.Text = "0.0") Then R2ClusterZ2.Text = ""
        If ClusterOnlyZ.Checked = True Or IndivOnlyZ.Checked = True Then
            ICCZ.Enabled = False
            ICCZ2.Enabled = False
            ICCZLabel.Enabled = False
            ICCZLabel2.Enabled = False
            ReverseZ.Enabled = False
        Else
            ICCZ.Enabled = True
            ICCZ2.Enabled = True
            ICCZLabel.Enabled = True
            ICCZLabel2.Enabled = True
            ReverseZ.Enabled = True
            If (TabControl1.SelectedIndex = 0) Then
                ICCZ.Text = ""
            Else
                ICCZ2.Text = ""
            End If
        End If
    End Sub

    Private Sub CleanOutput()
        ProgressBar1.Value = 0
        StatusNum.Text = 0
        WidthCI95.Text = ""
        WidthCI95C.Text = ""
        WidthCI99.Text = ""
        WidthCI99C.Text = ""
        Power.Text = ""
        PowerC.Text = ""
        NIndiv2.Text = ""
        NTGroups2.Text = ""
        NCGroups2.Text = ""
        TotalCost.Text = ""
        ObtainedPower.Text = ""
        ObtainedWidth95.Text = ""
        ObtainedWidth99.Text = ""
        ObtainedPowerC.Text = ""
        ObtainedWidth95C.Text = ""
        ObtainedWidth99C.Text = ""
        IsOutputDegree.Enabled = False
        IsOutputDegree2.Enabled = False
        OutputCalcDegree.Enabled = False
        OutputCalcDegree2.Enabled = False
        OutputDegreeLevel.Enabled = False
        OutputDegreeLevel2.Enabled = False
        OutputDegreeLevel.Text = ""
        OutputDegreeLevel2.Text = ""
        IsOutputDegree.ForeColor = Color.Black
        IsOutputDegree.ForeColor = Color.Black
        Width95OutputLabel.ForeColor = Color.Black
        Width99OutputLabel.ForeColor = Color.Black
        Width95OutputLabelC.ForeColor = Color.Black
        Width99OutputLabelC.ForeColor = Color.Black
        Width95OutputLabel2.ForeColor = Color.Black
        Width99OutputLabel2.ForeColor = Color.Black
        Width95OutputLabel2C.ForeColor = Color.Black
        Width99OutputLabel2C.ForeColor = Color.Black
        IsOutputDegree.Checked = False
        IsOutputDegree2.Checked = False
    End Sub

    Private Sub OutputCalcDegree2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputCalcDegree2.Click
        If IsOutputDegree2.Checked = True Then
            If IsDoubleInRange(OutputDegreeLevel2, 0.001, 0.999, "Degree of Certainty") Then
                Dim Position As Integer = Math.Ceiling(CDbl(OutputDegreeLevel2.Text) * NRep) - 1
                IsOutputDegree2.ForeColor = Color.Blue
                Width95OutputLabel2.ForeColor = Color.Blue
                Width99OutputLabel2.ForeColor = Color.Blue
                Width95OutputLabel2C.ForeColor = Color.Blue
                Width99OutputLabel2C.ForeColor = Color.Blue
                If checkCovariateInstanceVariable Then
                    ObtainedWidth95C.Text = PrintWidth(ESValue, Storage.Width(Position, 2)) 'Storage.Width(Position, 2).ToString("f3") & " (" & (ESValue - Storage.Width(Position, 2) / 2).ToString("f3") & ", " & (ESValue + Storage.Width(Position, 2) / 2).ToString("f3") & ")"
                    ObtainedWidth99C.Text = PrintWidth(ESValue, Storage.Width(Position, 3)) 'Storage.Width(Position, 3).ToString("f3") & " (" & (ESValue - Storage.Width(Position, 3) / 2).ToString("f3") & ", " & (ESValue + Storage.Width(Position, 3) / 2).ToString("f3") & ")"
                    If RunNullModelInstanceVariable Then
                        ObtainedWidth95.Text = PrintWidth(ESValue, Storage.Width(Position, 0)) 'Storage.Width(Position, 0).ToString("f3") & " (" & (ESValue - Storage.Width(Position, 0) / 2).ToString("f3") & ", " & (ESValue + Storage.Width(Position, 0) / 2).ToString("f3") & ")"
                        ObtainedWidth99.Text = PrintWidth(ESValue, Storage.Width(Position, 1)) 'Storage.Width(Position, 1).ToString("f3") & " (" & (ESValue - Storage.Width(Position, 1) / 2).ToString("f3") & ", " & (ESValue + Storage.Width(Position, 1) / 2).ToString("f3") & ")"
                    End If
                Else
                    ObtainedWidth95.Text = PrintWidth(ESValue, Storage.Width(Position, 0)) 'Storage.Width(Position, 0).ToString("f3") & " (" & (ESValue - Storage.Width(Position, 0) / 2).ToString("f3") & ", " & (ESValue + Storage.Width(Position, 0) / 2).ToString("f3") & ")"
                    ObtainedWidth99.Text = PrintWidth(ESValue, Storage.Width(Position, 1)) 'Storage.Width(Position, 1).ToString("f3")
                End If
            End If
        Else
            IsOutputDegree2.ForeColor = Color.Black
            Width95OutputLabel2.ForeColor = Color.Black
            Width99OutputLabel2.ForeColor = Color.Black
            Width95OutputLabel2C.ForeColor = Color.Black
            Width99OutputLabel2C.ForeColor = Color.Black
            If checkCovariateInstanceVariable = True Then
                Dim sumwidth95C As Double
                Dim sumwidth99C As Double
                For i = 0 To Storage.Width.GetLength(0) - 1
                    sumwidth95C += Storage.Width(i, 2)
                    sumwidth99C += Storage.Width(i, 3)
                Next i
                Dim Average95C As Double = sumwidth95C / Storage.Width.GetLength(0)
                Dim Average99C As Double = sumwidth99C / Storage.Width.GetLength(0)
                ObtainedWidth95C.Text = PrintWidth(ESValue, Average95C) 'Average95C.ToString("f3")
                ObtainedWidth99C.Text = PrintWidth(ESValue, Average99C) 'Average99C.ToString("f3")
                If RunNullModelInstanceVariable = True Then
                    Dim sumwidth95 As Double
                    Dim sumwidth99 As Double
                    For i = 0 To Storage.Width.GetLength(0) - 1
                        sumwidth95 += Storage.Width(i, 0)
                        sumwidth99 += Storage.Width(i, 1)
                    Next i
                    Dim Average95 As Double = sumwidth95 / Storage.Width.GetLength(0)
                    Dim Average99 As Double = sumwidth99 / Storage.Width.GetLength(0)
                    ObtainedWidth95.Text = PrintWidth(ESValue, Average95) 'Average95.ToString("f3")
                    ObtainedWidth99.Text = PrintWidth(ESValue, Average99) ' Average99.ToString("f3")
                End If
            Else
                Dim sumwidth95 As Double
                Dim sumwidth99 As Double
                For i = 0 To Storage.Width.GetLength(0) - 1
                    sumwidth95 += Storage.Width(i, 0)
                    sumwidth99 += Storage.Width(i, 1)
                Next i
                Dim Average95 As Double = sumwidth95 / Storage.Width.GetLength(0)
                Dim Average99 As Double = sumwidth99 / Storage.Width.GetLength(0)
                ObtainedWidth95.Text = PrintWidth(ESValue, Average95) 'Average95.ToString("f3")
                ObtainedWidth99.Text = PrintWidth(ESValue, Average99) 'Average99.ToString("f3")
            End If
        End If
    End Sub

    Private Sub ActivateDegreeofCertainty()
        checkCovariateInstanceVariable = checkCovariate.Checked
        RunNullModelInstanceVariable = RunNullModel.Checked
        If MonteCarlo.Checked = True Then
            If TabControl1.SelectedIndex = 0 Then
                IsOutputDegree.Enabled = True
                OutputCalcDegree.Enabled = True
            Else
                IsOutputDegree2.Enabled = True
                OutputCalcDegree2.Enabled = True
            End If
        End If
    End Sub

    Private Sub IsOutputDegree2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IsOutputDegree2.CheckedChanged
        If IsOutputDegree2.Checked Then
            OutputDegreeLevel2.Enabled = True
        Else
            OutputDegreeLevel2.Enabled = False
        End If
    End Sub

    Private Sub IsOutputDegree_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IsOutputDegree.CheckedChanged
        If IsOutputDegree.Checked Then
            OutputDegreeLevel.Enabled = True
        Else
            OutputDegreeLevel.Enabled = False
        End If
    End Sub

    Private Sub OutputCalcDegree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputCalcDegree.Click
        If IsOutputDegree.Checked = True Then
            If IsDoubleInRange(OutputDegreeLevel, 0.001, 0.999, "Degree of Certainty") Then
                Dim Position As Integer = Math.Ceiling(CDbl(OutputDegreeLevel.Text) * NRep) - 1
                IsOutputDegree.ForeColor = Color.Blue
                Width95OutputLabel.ForeColor = Color.Blue
                Width99OutputLabel.ForeColor = Color.Blue
                Width95OutputLabelC.ForeColor = Color.Blue
                Width99OutputLabelC.ForeColor = Color.Blue
                If checkCovariateInstanceVariable Then
                    WidthCI95C.Text = PrintWidth(ESValue, Storage.Width(Position, 2)) 'Storage.Width(Position, 2).ToString("f3")
                    WidthCI99C.Text = PrintWidth(ESValue, Storage.Width(Position, 3)) 'Storage.Width(Position, 3).ToString("f3")
                    If RunNullModelInstanceVariable Then
                        WidthCI95.Text = PrintWidth(ESValue, Storage.Width(Position, 0)) 'Storage.Width(Position, 0).ToString("f3")
                        WidthCI99.Text = PrintWidth(ESValue, Storage.Width(Position, 1)) 'Storage.Width(Position, 1).ToString("f3")
                    End If
                Else
                    WidthCI95.Text = PrintWidth(ESValue, Storage.Width(Position, 0)) 'Storage.Width(Position, 0).ToString("f3")
                    WidthCI99.Text = PrintWidth(ESValue, Storage.Width(Position, 1)) 'Storage.Width(Position, 1).ToString("f3")
                End If
            End If
        Else
            IsOutputDegree.ForeColor = Color.Black
            Width95OutputLabel.ForeColor = Color.Black
            Width99OutputLabel.ForeColor = Color.Black
            Width95OutputLabelC.ForeColor = Color.Black
            Width99OutputLabelC.ForeColor = Color.Black
            If checkCovariateInstanceVariable = True Then
                Dim sumwidth95C As Double
                Dim sumwidth99C As Double
                For i = 0 To Storage.Width.GetLength(0) - 1
                    sumwidth95C += Storage.Width(i, 2)
                    sumwidth99C += Storage.Width(i, 3)
                Next i
                Dim Average95C As Double = sumwidth95C / Storage.Width.GetLength(0)
                Dim Average99C As Double = sumwidth99C / Storage.Width.GetLength(0)
                WidthCI95C.Text = PrintWidth(ESValue, Average95C) 'Average95C.ToString("f3")
                WidthCI99C.Text = PrintWidth(ESValue, Average99C) 'Average99C.ToString("f3")
                If RunNullModelInstanceVariable = True Then
                    Dim sumwidth95 As Double
                    Dim sumwidth99 As Double
                    For i = 0 To Storage.Width.GetLength(0) - 1
                        sumwidth95 += Storage.Width(i, 0)
                        sumwidth99 += Storage.Width(i, 1)
                    Next i
                    Dim Average95 As Double = sumwidth95 / Storage.Width.GetLength(0)
                    Dim Average99 As Double = sumwidth99 / Storage.Width.GetLength(0)
                    WidthCI95.Text = PrintWidth(ESValue, Average95) 'Average95.ToString("f3")
                    WidthCI99.Text = PrintWidth(ESValue, Average99) 'Average99.ToString("f3")
                End If
            Else
                Dim sumwidth95 As Double
                Dim sumwidth99 As Double
                For i = 0 To Storage.Width.GetLength(0) - 1
                    sumwidth95 += Storage.Width(i, 0)
                    sumwidth99 += Storage.Width(i, 1)
                Next i
                Dim Average95 As Double = sumwidth95 / Storage.Width.GetLength(0)
                Dim Average99 As Double = sumwidth99 / Storage.Width.GetLength(0)
                WidthCI95.Text = PrintWidth(ESValue, Average95) 'Average95.ToString("f3")
                WidthCI99.Text = PrintWidth(ESValue, Average99) 'Average99.ToString("f3")
            End If
        End If

    End Sub

    Private Sub MonteCarlo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MonteCarlo.CheckedChanged
        If MonteCarlo.Checked = False Then
            If Criteria.IsDegreeOfCertainty.SelectedIndex = 1 Then MsgBox("The degree of certainty feature depends on the simulation results. Because you checked on the degree of certainty, it is changed to be unchecked.")
            Criteria.IsDegreeOfCertainty.SelectedIndex = 0
            Criteria.DegreeOfCertainty.Text = ""
            Storage.IsDegreeOfCertaintyValue = False
            Storage.DegreeOfCertaintyValue = -1
            Criteria.DegreeOfCertaintyLabel.Enabled = False
            Criteria.IsDegreeOfCertainty.Enabled = False
            Criteria.DegreeOfCertaintyLabel2.Enabled = False
            Criteria.DegreeOfCertainty.Enabled = False
        End If
    End Sub


    Private Sub IndivOnlyZ_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IndivOnlyZ.CheckedChanged
        If IndivOnlyZ.Checked = True Then
            ICCZLabel.Enabled = False
            ICCZLabel2.Enabled = False
            ICCZ.Enabled = False
            ICCZ2.Enabled = False
            R2ClusterZ.Enabled = False
            R2ClusterZ2.Enabled = False
            R2ClusterZLabel.Enabled = False
            R2ClusterZLabel2.Enabled = False
            ClusterOnlyZ.Checked = False
            R2IndivZ.Enabled = True
            R2IndivZ2.Enabled = True
            R2IndivdualZLabel.Enabled = True
            R2IndivZLabel2.Enabled = True
            If TabControl1.SelectedIndex = 0 Then
                ICCZ.Text = "0.0"
                R2ClusterZ.Text = "0.0"
            ElseIf TabControl1.SelectedIndex = 1 Then
                ICCZ2.Text = "0.0"
                R2ClusterZ2.Text = "0.0"
            End If
            If R2IndivZ.Text = "0.0" Then R2IndivZ.Text = ""
            If R2IndivZ2.Text = "0.0" Then R2IndivZ2.Text = ""
            ReverseZ.Checked = False
        Else
            R2ClusterZ.Enabled = True
            R2ClusterZ2.Enabled = True
            R2ClusterZ.Text = ""
            R2ClusterZ2.Text = ""
            R2ClusterZLabel.Enabled = True
            R2ClusterZLabel2.Enabled = True
        End If
        If ClusterOnlyZ.Checked = True Or IndivOnlyZ.Checked = True Then
            ICCZ.Enabled = False
            ICCZ2.Enabled = False
            ICCZLabel.Enabled = False
            ICCZLabel2.Enabled = False
            ReverseZ.Enabled = False
        Else
            ICCZ.Enabled = True
            ICCZ2.Enabled = True
            ICCZLabel.Enabled = True
            ICCZLabel2.Enabled = True
            ReverseZ.Enabled = True
            If (TabControl1.SelectedIndex = 0) Then
                ICCZ.Text = ""
            Else
                ICCZ2.Text = ""
            End If
        End If
    End Sub

    Private Sub UsersManualToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsersManualToolStripMenuItem.Click
        Dim process As Process
        process = New Process()
        process.StartInfo.FileName = "IExplore"
        process.StartInfo.Arguments = "www.google.com"
        'process.Start("PAWS User Manual.pdf")

    End Sub
End Class
