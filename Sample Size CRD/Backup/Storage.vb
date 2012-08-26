Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.ComponentModel
Imports System.Threading

Module Storage

    Public Const DefaultReplication As Int16 = 1000
    Public Const Banner As String = "CRD Helper" & vbCrLf & "version 1.0 (February 2010)" & vbCrLf & "Developer/Copyright: Sunthud Pornprasertmanit & W. Joel Schneider" & vbCrLf

    'Options instance variables
    Public NRep As Integer
    Public Path As String = My.Application.Info.DirectoryPath
    'Model instance variables
    Public ICCYValue As Double
    Public ESValue As Double
    Public ICCZValue As Double
    Public R2ClusterZValue As Double
    Public R2IndivZValue As Double
    'Sample Size Instance Variables
    Public NTValue As Integer = -1
    Public NCValue As Integer = -1
    Public NGValue As Integer = -1
    Public PTValue As Double = -1.0
    Public NIndivValue As Integer = -1
    'Option variable
    Public OptionValue As Integer = -1
    'Power or width instance variables
    Public CILevelValue As Double = -1
    Public PowerOrWidthValue As Integer = -1
    Public PowerValue As Double = -1 ' 1 = Power, 2 = Width
    Public WidthValue As Double = -1
    Public IsDegreeOfCertaintyValue As Boolean = False
    Public DegreeOfCertaintyValue As Double = -1
    'Cost instance variable
    Public TGroupCostValue As Double = -1
    Public CGroupCostValue As Double = -1
    Public TIndivCostValue As Double = -1
    Public CIndivCostValue As Double = -1
    Public TotalCostValue As Double = -1
    'Storage Widths
    Public Width(,) As Double
    Public StartingValue(0 To 2) As Integer
    Public MonteCarloArchive(0 To 999, 0 To 3) As Double
    Public MonteCarloLine As Integer = 0
    Public MonteCarloMaxLine As Integer = 1000
    'Minimumum Individual
    Public MinIndiv As Integer

    'Thread
    'Public MplusDirectory As String = "C:\Program Files\Mplus\mplus.exe"

    Public pJ11(,) As Integer = {{2, 2}, _
                               {2, 3}, {3, 2}, _
                               {2, 4}, {3, 3}, {4, 2}, _
                               {2, 5}, {3, 4}, {4, 3}, {5, 2}, _
                               {2, 6}, {3, 5}, {4, 4}, {5, 3}, {6, 2}, _
                               {2, 7}, {3, 6}, {4, 5}, {5, 4}, {6, 3}, {7, 2}, _
                               {2, 8}, {3, 7}, {4, 6}, {5, 5}, {6, 4}, {7, 3}, {8, 2}, _
                               {2, 9}, {3, 8}, {4, 7}, {5, 6}, {6, 5}, {7, 4}, {8, 3}, {9, 2}}



    Sub WriteMCNullCode(ByVal nT As Integer, ByVal NC As Integer, ByVal ni As Integer, ByVal iccy As Double, ByVal es As Double, ByVal NRep As Integer, Optional ByVal MakeFiles As Boolean = True)
        Dim NObs As Integer = (nT + NC) * ni
        Dim probTreatment = CDbl(nT) / (nT + NC)
        Dim VarYBet As Double = (iccy / (1 - iccy))
        Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                            "MONTECARLO:" & vbCrLf & _
                            "NAMES ARE y x;" & vbCrLf & _
                            "BETWEEN = x;" & vbCrLf & _
                            "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                            "NCSIZES = 1;" & vbCrLf & _
                            "CSIZES = " & (nT + NC) & " (" & ni & ");" & vbCrLf & _
                            "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                            "NREP = " & NRep & ";" & vbCrLf
        If MakeFiles = True Then
            first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
            '"SAVE = " & Path & "\raw*.dat;" & vbCrLf

        End If
        first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                 "MODEL POPULATION:" & vbCrLf
        Dim model As String = "%WITHIN%" & vbCrLf & _
                                "y*1;" & vbCrLf & _
                                "%BETWEEN%" & vbCrLf & _
                                "y ON x*" & es & ";" & vbCrLf & _
                                "y*" & VarYBet.ToString("f5") & ";"
        Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
        FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
        PrintLine(1, code)
        FileClose(1)
    End Sub

    Sub WriteMCNullCode(ByVal nJ As Integer, ByVal pT As Double, ByVal ni As Integer, ByVal iccy As Double, ByVal es As Double, ByVal NRep As Integer, Optional ByVal MakeFiles As Boolean = True)
        Dim NObs As Integer = nJ * ni
        Dim probTreatment As Double = pT
        Dim VarYBet As Double = (iccy / (1 - iccy))
        Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                            "MONTECARLO:" & vbCrLf & _
                            "NAMES ARE y x;" & vbCrLf & _
                            "BETWEEN = x;" & vbCrLf & _
                            "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                            "NCSIZES = 1;" & vbCrLf & _
                            "CSIZES = " & nJ & " (" & ni & ");" & vbCrLf & _
                            "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                            "NREP = " & NRep & ";" & vbCrLf
        If MakeFiles = True Then
            first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
            '"SAVE = " & Path & "\raw*.dat;" & vbCrLf
        End If
        first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                 "MODEL POPULATION:" & vbCrLf
        Dim model As String = "%WITHIN%" & vbCrLf & _
                                "y*1;" & vbCrLf & _
                                "%BETWEEN%" & vbCrLf & _
                                "y ON x*" & es & ";" & vbCrLf & _
                                "y*" & VarYBet.ToString("f5") & ";"
        Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
        FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
        PrintLine(1, code)
        FileClose(1)
    End Sub

    Sub WriteMCCovariateCode(ByVal nT As Integer, ByVal nC As Integer, ByVal ni As Integer, ByVal iccy As Double, ByVal es As Double, ByVal iccz As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double, ByVal NRep As Integer, ByVal MakeFiles As Boolean, ByVal ClusterOnlyZ As Boolean, ByVal ReverseZ As Boolean)
        If ClusterOnlyZ Then
            'Dim dfB As Integer = nT + nC - 1
            'Dim dfW As Integer = (nT + nC) * (ni - 1)
            Dim tauZ As Double = 1
            Dim VarYBet As Double = (1 - R2ClusterZ) * (iccy / (1 - iccy))
            Dim VarYWithin As Double = 1
            'Dim SSBZ As Double = dfB * tauZ
            'Dim SSWZ As Double = dfW * sigmaZ
            'Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
            Dim BetaB As Double = Math.Sqrt(R2ClusterZ * VarYBet / tauZ)
            Dim NObs As Integer = (nT + nC) * ni
            Dim probTreatment = CDbl(nT) / (nT + nC)

            Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                                "MONTECARLO:" & vbCrLf & _
                                "NAMES ARE y x z;" & vbCrLf & _
                                "BETWEEN = x z;" & vbCrLf & _
                                "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                                "NCSIZES = 1;" & vbCrLf & _
                                "CSIZES = " & (nT + nC) & " (" & ni & ");" & vbCrLf & _
                                "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                                "NREP = " & NRep & ";" & vbCrLf
            If MakeFiles = True Then
                first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
                '"SAVE = " & Path & "\raw*.dat;" & vbCrLf
            End If
            first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                     "MODEL POPULATION:" & vbCrLf
            Dim model As String = "%WITHIN%" & vbCrLf & _
                                    "y*" & VarYWithin.ToString("f5") & ";" & vbCrLf & _
                                    "%BETWEEN%" & vbCrLf & _
                                    "y ON x*" & es & ";" & vbCrLf & _
                                    "y ON z*" & BetaB.ToString("f5") & ";" & vbCrLf & _
                                    "z*" & tauZ.ToString("f5") & ";" & vbCrLf & _
                                    "y*" & VarYBet.ToString("f5") & ";"
            Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
            FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
            PrintLine(1, code)
            FileClose(1)
        Else
            If iccz = 0 Then
                'Dim dfB As Integer = nT + nC - 1
                'Dim dfW As Integer = (nT + nC) * (ni - 1)
                Dim sigmaZ As Double = 1
                Dim VarYBet As Double = (1 - R2ClusterZ) * (iccy / (1 - iccy))
                Dim VarYWithin As Double = 1 - R2IndivZ
                'Dim SSBZ As Double = dfB * tauZ
                'Dim SSWZ As Double = dfW * sigmaZ
                'Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
                Dim BetaW As Double = Math.Sqrt(R2IndivZ * VarYWithin / sigmaZ)
                Dim NObs As Integer = (nT + nC) * ni
                Dim probTreatment = CDbl(nT) / (nT + nC)

                Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                                    "MONTECARLO:" & vbCrLf & _
                                    "NAMES ARE y x z;" & vbCrLf & _
                                    "BETWEEN = x;" & vbCrLf & _
                                    "WITHIN = z;" & vbCrLf & _
                                    "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                                    "NCSIZES = 1;" & vbCrLf & _
                                    "CSIZES = " & (nT + nC) & " (" & ni & ");" & vbCrLf & _
                                    "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                                    "NREP = " & NRep & ";" & vbCrLf
                If MakeFiles = True Then
                    first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
                    '"SAVE = " & Path & "\raw*.dat;" & vbCrLf
                End If
                first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                         "MODEL POPULATION:" & vbCrLf
                Dim model As String = "%WITHIN%" & vbCrLf & _
                                        "y*" & VarYWithin.ToString("f5") & ";" & vbCrLf & _
                                        "z*" & sigmaZ.ToString("f5") & ";" & vbCrLf & _
                                        "y ON z*" & BetaW.ToString("f5") & ";" & vbCrLf & _
                                        "%BETWEEN%" & vbCrLf & _
                                        "y ON x*" & es & ";" & vbCrLf & _
                                        "y*" & VarYBet.ToString("f5") & ";"
                Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
                FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
                PrintLine(1, code)
                FileClose(1)
            Else
                'Dim dfB As Integer = nT + nC - 1
                'Dim dfW As Integer = (nT + nC) * (ni - 1)
                Dim tauZ As Double = iccz / (1 - iccz)
                Dim sigmaZ As Double = 1
                Dim VarYBet As Double = (1 - R2ClusterZ) * (iccy / (1 - iccy))
                Dim VarYWithin As Double = 1 - R2IndivZ
                'Dim SSBZ As Double = dfB * tauZ
                'Dim SSWZ As Double = dfW * sigmaZ
                'Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
                Dim BetaB As Double = Math.Sqrt(R2ClusterZ * VarYBet / tauZ)
                Dim BetaW As Double
                If ReverseZ Then
                    BetaW = 0 - Math.Sqrt(R2IndivZ * VarYWithin / sigmaZ)
                Else
                    BetaW = Math.Sqrt(R2IndivZ * VarYWithin / sigmaZ)
                End If
                Dim NObs As Integer = (nT + nC) * ni
                Dim probTreatment = CDbl(nT) / (nT + nC)

                Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                                    "MONTECARLO:" & vbCrLf & _
                                    "NAMES ARE y x z;" & vbCrLf & _
                                    "BETWEEN = x;" & vbCrLf & _
                                    "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                                    "NCSIZES = 1;" & vbCrLf & _
                                    "CSIZES = " & (nT + nC) & " (" & ni & ");" & vbCrLf & _
                                    "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                                    "NREP = " & NRep & ";" & vbCrLf
                If MakeFiles = True Then
                    first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
                    '"SAVE = " & Path & "\raw*.dat;" & vbCrLf
                End If
                first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                         "MODEL POPULATION:" & vbCrLf
                Dim model As String = "%WITHIN%" & vbCrLf & _
                                        "y*" & VarYWithin.ToString("f5") & ";" & vbCrLf & _
                                        "z*" & sigmaZ.ToString("f5") & ";" & vbCrLf & _
                                        "y ON z*" & BetaW.ToString("f5") & ";" & vbCrLf & _
                                        "%BETWEEN%" & vbCrLf & _
                                        "y ON x*" & es & ";" & vbCrLf & _
                                        "y ON z*" & BetaB.ToString("f5") & ";" & vbCrLf & _
                                        "z*" & tauZ.ToString("f5") & ";" & vbCrLf & _
                                        "y*" & VarYBet.ToString("f5") & ";"
                Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
                FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
                PrintLine(1, code)
                FileClose(1)
            End If
            
        End If

    End Sub

    Sub WriteMCCovariateCode(ByVal nJ As Integer, ByVal pT As Double, ByVal ni As Integer, ByVal iccy As Double, ByVal es As Double, ByVal iccz As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double, ByVal NRep As Integer, ByVal MakeFiles As Boolean, ByVal ClusterOnlyZ As Boolean, ByVal ReverseZ As Boolean)
        If ClusterOnlyZ Then
            'Dim dfB As Integer = nT + nC - 1
            'Dim dfW As Integer = (nT + nC) * (ni - 1)
            Dim tauZ As Double = 1
            Dim VarYBet As Double = (1 - R2ClusterZ) * (iccy / (1 - iccy))
            Dim VarYWithin As Double = 1
            'Dim SSBZ As Double = dfB * tauZ
            'Dim SSWZ As Double = dfW * sigmaZ
            'Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
            Dim BetaB As Double = Math.Sqrt(R2ClusterZ * VarYBet / tauZ)
            Dim NObs As Integer = nJ * ni
            Dim probTreatment As Double = pT
            Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                                "MONTECARLO:" & vbCrLf & _
                                "NAMES ARE y x z;" & vbCrLf & _
                                "BETWEEN = x z;" & vbCrLf & _
                                "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                                "NCSIZES = 1;" & vbCrLf & _
                                "CSIZES = " & nJ & " (" & ni & ");" & vbCrLf & _
                                "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                                "NREP = " & NRep & ";" & vbCrLf
            If MakeFiles = True Then
                first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
                '"SAVE = " & Path & "\raw*.dat;" & vbCrLf
            End If
            first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                     "MODEL POPULATION:" & vbCrLf
            Dim model As String = "%WITHIN%" & vbCrLf & _
                                    "y*" & VarYWithin.ToString("f5") & ";" & vbCrLf & _
                                    "%BETWEEN%" & vbCrLf & _
                                    "y ON x*" & es & ";" & vbCrLf & _
                                    "y ON z*" & BetaB.ToString("f5") & ";" & vbCrLf & _
                                    "z*" & tauZ.ToString("f5") & ";" & vbCrLf & _
                                    "y*" & VarYBet.ToString("f5") & ";"
            Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
            FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
            PrintLine(1, code)
            FileClose(1)
        Else
            If iccz = 0 Then
                'Dim dfB As Integer = nJ - 1
                'Dim dfW As Integer = nJ * (ni - 1)
                Dim sigmaZ As Double = 1
                Dim VarYBet As Double = (1 - R2ClusterZ) * (iccy / (1 - iccy))
                Dim VarYWithin As Double = 1 - R2IndivZ
                'Dim SSBZ As Double = dfB * tauZ
                'Dim SSWZ As Double = dfW * sigmaZ
                'Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
                Dim BetaW As Double= Math.Sqrt(R2IndivZ * VarYWithin / sigmaZ)
                Dim NObs As Integer = nJ * ni
                Dim probTreatment As Double = pT
                Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                                    "MONTECARLO:" & vbCrLf & _
                                    "NAMES ARE y x z;" & vbCrLf & _
                                    "BETWEEN = x;" & vbCrLf & _
                                    "WITHIN = z;" & vbCrLf & _
                                    "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                                    "NCSIZES = 1;" & vbCrLf & _
                                    "CSIZES = " & nJ & " (" & ni & ");" & vbCrLf & _
                                    "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                                    "NREP = " & NRep & ";" & vbCrLf
                If MakeFiles = True Then
                    first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
                    '"SAVE = " & Path & "\raw*.dat;" & vbCrLf
                End If
                first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                         "MODEL POPULATION:" & vbCrLf
                Dim model As String = "%WITHIN%" & vbCrLf & _
                                        "y*" & VarYWithin.ToString("f5") & ";" & vbCrLf & _
                                        "z*" & sigmaZ.ToString("f5") & ";" & vbCrLf & _
                                        "y ON z*" & BetaW.ToString("f5") & ";" & vbCrLf & _
                                        "%BETWEEN%" & vbCrLf & _
                                        "y ON x*" & es & ";" & vbCrLf & _
                                        "y*" & VarYBet.ToString("f5") & ";"
                Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
                FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
                PrintLine(1, code)
                FileClose(1)

            Else
                'Dim dfB As Integer = nJ - 1
                'Dim dfW As Integer = nJ * (ni - 1)
                Dim tauZ As Double = iccz / (1 - iccz)
                Dim sigmaZ As Double = 1
                Dim VarYBet As Double = (1 - R2ClusterZ) * (iccy / (1 - iccy))
                Dim VarYWithin As Double = 1 - R2IndivZ
                'Dim SSBZ As Double = dfB * tauZ
                'Dim SSWZ As Double = dfW * sigmaZ
                'Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
                Dim BetaB As Double = Math.Sqrt(R2ClusterZ * VarYBet / tauZ)
                Dim BetaW As Double
                If ReverseZ Then
                    BetaW = 0 - Math.Sqrt(R2IndivZ * VarYWithin / sigmaZ)
                Else
                    BetaW = Math.Sqrt(R2IndivZ * VarYWithin / sigmaZ)
                End If
                Dim NObs As Integer = nJ * ni
                Dim probTreatment As Double = pT


                Dim first As String = "TITLE: Mplus: Generate raw data for CRD" & vbCrLf & _
                                    "MONTECARLO:" & vbCrLf & _
                                    "NAMES ARE y x z;" & vbCrLf & _
                                    "BETWEEN = x;" & vbCrLf & _
                                    "NOBSERVATIONS = " & NObs & ";" & vbCrLf & _
                                    "NCSIZES = 1;" & vbCrLf & _
                                    "CSIZES = " & nJ & " (" & ni & ");" & vbCrLf & _
                                    "CUTPOINTS = x(" & Stat.InvNormalDistribution(probTreatment).ToString("f5") & ");" & vbCrLf & _
                                    "NREP = " & NRep & ";" & vbCrLf
                If MakeFiles = True Then
                    first &= "REPSAVE = ALL;" & vbCrLf & _
                            "SAVE = raw*.dat;" & vbCrLf
                    '"SAVE = " & Path & "\raw*.dat;" & vbCrLf
                End If
                first &= "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                         "MODEL POPULATION:" & vbCrLf
                Dim model As String = "%WITHIN%" & vbCrLf & _
                                        "y*" & VarYWithin.ToString("f5") & ";" & vbCrLf & _
                                        "z*" & sigmaZ.ToString("f5") & ";" & vbCrLf & _
                                        "y ON z*" & BetaW.ToString("f5") & ";" & vbCrLf & _
                                        "%BETWEEN%" & vbCrLf & _
                                        "y ON x*" & es & ";" & vbCrLf & _
                                        "y ON z*" & BetaB.ToString("f5") & ";" & vbCrLf & _
                                        "z*" & tauZ.ToString("f5") & ";" & vbCrLf & _
                                        "y*" & VarYBet.ToString("f5") & ";"
                Dim code As String = first & model & "x*1; " & vbCrLf & "MODEL:" & vbCrLf & model & vbCrLf & "OUTPUT: TECH9;"
                FileOpen(1, Path & "\simulation.inp", OpenMode.Output)
                PrintLine(1, code)
                FileClose(1)

            End If
        End If

    End Sub

    Sub WriteNullAnalysisCode(ByVal filename As String, ByVal extraCovariate As Boolean, ByVal ClusterOnlyZ As Boolean)
        Dim title As String
        If extraCovariate = True Then
            If ClusterOnlyZ Then
                title = "TITLE: Mplus: Wald CI on the effect size on CRD w/o covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                               "VARIABLE: NAMES ARE Z Y X CLUS;" & vbCrLf & _
                               "USEVARIABLES ARE Y X CLUS;" & vbCrLf '"DATA: FILE IS " & Path & "\" & filename & ".dat;" & vbCrLf & _
            Else
                title = "TITLE: Mplus: Wald CI on the effect size on CRD w/o covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                                "VARIABLE: NAMES ARE Y Z X CLUS;" & vbCrLf & _
                                "USEVARIABLES ARE Y X CLUS;" & vbCrLf
            End If

        Else
            title = "TITLE: Mplus: Wald CI on the effect size on CRD w/o covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                            "VARIABLE: NAMES ARE Y X CLUS;" & vbCrLf
        End If
        Dim code As String = title & _
                            "BETWEEN = X;" & vbCrLf & _
                            "CLUSTER = CLUS;" & vbCrLf & _
                            "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                            "MODEL:" & vbCrLf & _
                            "%WITHIN%" & vbCrLf & _
                            "y (e);" & vbCrLf & _
                            "%BETWEEN%" & vbCrLf & _
                            "y ON x (d);" & vbCrLf & _
                            "MODEL CONSTRAINT:" & vbCrLf & _
                            "NEW(es);" & vbCrLf & _
                            "es = d/sqrt(e);" & vbCrLf & _
                            "OUTPUT: CINTERVAL(symmetric);"
        Dim fullfilename As String = Path & "\" & filename & ".inp"
        FileOpen(1, fullfilename, OpenMode.Output)
        PrintLine(1, code)
        FileClose(1)
    End Sub
    Sub WriteCovariateAnalysisCode(ByVal filename As String, ByVal ClusterOnlyZ As Boolean)
        If ClusterOnlyZ Then
            Dim code As String = "TITLE: Mplus: Wald CI on the effect size on CRD with cluster-level covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                                            "VARIABLE: NAMES ARE Z Y X CLUS;" & vbCrLf & _
                                            "BETWEEN = Z X;" & vbCrLf & _
                                            "CLUSTER = CLUS;" & vbCrLf & _
                                            "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                                            "MODEL:" & vbCrLf & _
                                            "%WITHIN%" & vbCrLf & _
                                            "y (e);" & vbCrLf & _
                                            "%BETWEEN%" & vbCrLf & _
                                            "y ON x (d);" & vbCrLf & _
                                            "y ON z;" & vbCrLf & _
                                            "MODEL CONSTRAINT:" & vbCrLf & _
                                            "NEW(es);" & vbCrLf & _
                                            "es = d/sqrt(e);" & vbCrLf & _
                                            "OUTPUT: CINTERVAL(symmetric);"
            Dim fullfilename As String = Path & "\" & filename & ".inp"
            FileOpen(1, fullfilename, OpenMode.Output)
            PrintLine(1, code)
            FileClose(1)
        Else
            If ICCZValue = 0 Then
                Dim code As String = "TITLE: Mplus: Wald CI on the effect size on CRD with individual-level covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                    "VARIABLE: NAMES ARE Y Z X CLUS;" & vbCrLf & _
                    "BETWEEN = X;" & vbCrLf & _
                    "WITHIN = Z;" & vbCrLf & _
                    "CLUSTER = CLUS;" & vbCrLf & _
                    "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                    "MODEL:" & vbCrLf & _
                    "%WITHIN%" & vbCrLf & _
                    "y (e);" & vbCrLf & _
                    "y on z (g);" & vbCrLf & _
                    "z (m);" & vbCrLf & _
                    "%BETWEEN%" & vbCrLf & _
                    "y ON x (d);" & vbCrLf & _
                    "MODEL CONSTRAINT:" & vbCrLf & _
                    "NEW(es);" & vbCrLf & _
                    "es = d/sqrt((g * g * m) + e);" & vbCrLf & _
                    "OUTPUT: CINTERVAL(symmetric);"
                Dim fullfilename As String = Path & "\" & filename & ".inp"
                FileOpen(1, fullfilename, OpenMode.Output)
                PrintLine(1, code)
                FileClose(1)
            Else
                Dim code As String = "TITLE: Mplus: Wald CI on the effect size on CRD with covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                                    "VARIABLE: NAMES ARE Y Z X CLUS;" & vbCrLf & _
                                    "BETWEEN = X;" & vbCrLf & _
                                    "CLUSTER = CLUS;" & vbCrLf & _
                                    "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                                    "MODEL:" & vbCrLf & _
                                    "%WITHIN%" & vbCrLf & _
                                    "y (e);" & vbCrLf & _
                                    "y on z (g);" & vbCrLf & _
                                    "z (m);" & vbCrLf & _
                                    "%BETWEEN%" & vbCrLf & _
                                    "y ON x (d);" & vbCrLf & _
                                    "y ON z;" & vbCrLf & _
                                    "z;" & vbCrLf & _
                                    "MODEL CONSTRAINT:" & vbCrLf & _
                                    "NEW(es);" & vbCrLf & _
                                    "es = d/sqrt((g * g * m) + e);" & vbCrLf & _
                                    "OUTPUT: CINTERVAL(symmetric);"
                Dim fullfilename As String = Path & "\" & filename & ".inp"
                FileOpen(1, fullfilename, OpenMode.Output)
                PrintLine(1, code)
                FileClose(1)
            End If
        End If

    End Sub


    Sub WriteCovariateAnalysisCodeSecondTrial(ByVal filename As String, ByVal ClusterOnlyZ As Boolean)
        If ClusterOnlyZ Then
            Dim code As String = "TITLE: Mplus: Wald CI on the effect size on CRD with cluster-level covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                                            "VARIABLE: NAMES ARE Z Y X CLUS;" & vbCrLf & _
                                            "BETWEEN = Z X;" & vbCrLf & _
                                            "CLUSTER = CLUS;" & vbCrLf & _
                                            "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                                            "MODEL:" & vbCrLf & _
                                            "%WITHIN%" & vbCrLf & _
                                            "y (e);" & vbCrLf & _
                                            "%BETWEEN%" & vbCrLf & _
                                            "y ON x (d);" & vbCrLf & _
                                            "y ON z;" & vbCrLf & _
                                            "MODEL CONSTRAINT:" & vbCrLf & _
                                            "NEW(es);" & vbCrLf & _
                                            "es = d/sqrt(e);" & vbCrLf & _
                                            "OUTPUT: CINTERVAL(symmetric);"
            Dim fullfilename As String = Path & "\" & filename & ".inp"
            FileOpen(1, fullfilename, OpenMode.Output)
            PrintLine(1, code)
            FileClose(1)
        Else
            Dim code As String = "TITLE: Mplus: Wald CI on the effect size on CRD with indiviudal-level covariate" & vbCrLf & _
                               "DATA: FILE IS " & filename & ".dat;" & vbCrLf & _
                                "VARIABLE: NAMES ARE Y Z X CLUS;" & vbCrLf & _
                                "BETWEEN = X;" & vbCrLf & _
                                "WITHIN = Z;" & vbCrLf & _
                                "CLUSTER = CLUS;" & vbCrLf & _
                                "ANALYSIS: TYPE = TWOLEVEL;" & vbCrLf & _
                                "MODEL:" & vbCrLf & _
                                "%WITHIN%" & vbCrLf & _
                                "y (e);" & vbCrLf & _
                                "y on z (g);" & vbCrLf & _
                                "z (m);" & vbCrLf & _
                                "%BETWEEN%" & vbCrLf & _
                                "y ON x (d);" & vbCrLf & _
                                "MODEL CONSTRAINT:" & vbCrLf & _
                                "NEW(es);" & vbCrLf & _
                                "es = d/sqrt((g * g * m) + e);" & vbCrLf & _
                                "OUTPUT: CINTERVAL(symmetric);"
            Dim fullfilename As String = Path & "\" & filename & ".inp"
            FileOpen(1, fullfilename, OpenMode.Output)
            PrintLine(1, code)
            FileClose(1)
        End If

    End Sub

    Sub RunAnalysis(ByVal filename As String)
        'Shell("mplus.exe " & Path & "\" & filename & ".inp " & Path & "\" & filename & ".out")

        'Dim StartInfo As New ProcessStartInfo("cmd.exe")
        'StartInfo.Arguments = Path & "\" & filename & ".inp " '& Path & "\" & filename & ".out"
        'Dim process As Process = New Process
        'With process.StartInfo
        '    .FileName = "cmd.exe"
        'End With
        'process.Start()
        'process.StandardInput.WriteLine("/c Mplus.exe" & Convert.ToChar(13))
        'process.StandardInput.WriteLine("/c" & Path & "\" & filename & ".inp" & Convert.ToChar(13))
        'process.StandardInput.WriteLine("/c" & Convert.ToChar(13))

        'Dim StartInfo As New ProcessStartInfo("mplus.exe")
        'StartInfo.Arguments = Path & "\" & filename & ".inp " '& Path & "\" & filename & ".out"
        'Process.Start(StartInfo)

        'Dim psi As New System.Diagnostics.ProcessStartInfo("mplus.exe")
        'psi.Arguments = Path & "\" & filename & ".inp"
        'psi.UseShellExecute = True
        'psi.WindowStyle = ProcessWindowStyle.Hidden
        'Process.Start(psi)

        Dim process As Process
        process = New Process()
        process.StartInfo.FileName = My.Settings.MplusDirectory
        process.StartInfo.Arguments = """" & Path & "\" & filename & ".inp" & """"
        process.Start()
        process.WaitForExit(30000)
    End Sub

    Function ReadES(ByVal filename As String)
        Dim fullfilename As String = My.Application.Info.DirectoryPath & "\" & filename & ".out"
        Dim numLine As Short = 0
        Dim sr As StreamReader = ReadMplusOutputMod(fullfilename)
        Dim allText As String = sr.ReadToEnd
        sr.Close()

        Dim charInFile As Integer = allText.Length
        Dim i As Integer
        Dim letter As String
        For i = 0 To charInFile - 1
            letter = allText.Substring(i, 1)
            If letter = Chr(13) Then
                numLine += 1
            End If
        Next i

        Dim fileArray(numLine) As String
        Dim currentLine As Integer = 1
        Dim Line As String = ""
        For i = 0 To charInFile - 1
            letter = allText.Substring(i, 1)
            If letter = Chr(13) Then
                fileArray(currentLine - 1) = Line
                currentLine += 1
                Line = ""
            Else
                Line &= letter
            End If
        Next i

        'Dim DetectError As String = "ERROR in DATA command"
        'Dim Found As Boolean = False
        'Dim j As Integer = 0
        'While j < numLine
        '    If fileArray(j).Contains(DetectError) Then
        '        System.Threading.Thread.Sleep(10000)
        '        RunAnalysis(filename)
        '        ReadES(filename)
        '        Found = True
        '    End If
        '    j = j + 1
        'End While

        Dim find As String = "New/Additional"
        Dim numResult As Integer = 0
        Dim findResult(0 To 1) As Integer
        For i = 0 To numLine - 1
            If fileArray(i).Contains(find) Then
                findResult(numResult) = i
                numResult += 1
            End If
        Next i

        Dim esLine As String = fileArray(findResult(1) + 1)
        Dim splitSentence() As String = esLine.Split(" "c)

        currentLine = 0
        Dim cleanSentence(0 To 6) As String
        For i = 0 To UBound(splitSentence)
            If splitSentence(i) <> "" AndAlso splitSentence(i) <> " " Then
                cleanSentence(currentLine) = splitSentence(i)
                currentLine += 1
            End If
        Next i

        Dim es(0 To 5) As Double
        For i = 0 To 4
            es(i) = CDbl(cleanSentence(i + 2))
        Next i

        find = "Between Level"
        numResult = 0
        For i = 0 To numLine - 1
            If fileArray(i).Contains(find) Then
                findResult(numResult) = i
                numResult += 1
            End If
        Next i

        Dim powerLine As String = fileArray(findResult(0) + 3)
        Dim powerSentence() As String = powerLine.Split(" "c)

        currentLine = 0
        Dim cleanPowerSentence(0 To 5) As String
        For i = 0 To UBound(powerSentence)
            If powerSentence(i) <> "" AndAlso powerSentence(i) <> " " Then
                cleanPowerSentence(currentLine) = powerSentence(i)
                currentLine += 1
            End If
        Next i
        Dim sig As Double = 0.0
        If (CDbl(cleanPowerSentence(5)) <= 0.05) Then
            sig = 1.0
        End If
        es(5) = sig
        Return es
    End Function

    Function MAdd(ByVal A(,) As Double, ByVal B(,) As Double)
        If (A.GetLength(0) <> B.GetLength(0) And A.GetLength(1) <> B.GetLength(1)) Then
            MsgBox("The adding operation is not commutable since the A size is " & A.GetLength(0) & " * " & A.GetLength(1) & _
                   " ,but the B size is " & B.GetLength(0) & " * " & B.GetLength(1) & ".")
            End
        End If
        Dim nRow As Integer = A.GetLength(0)
        Dim nCol As Integer = A.GetLength(1)
        Dim C(0 To nRow - 1, 0 To nCol - 1) As Double
        For i = 0 To nRow - 1
            For j = 0 To nCol - 1
                C(i, j) = A(i, j) + B(i, j)
            Next j
        Next i
        Return C
    End Function

    Function MScalar(ByVal k As Double, ByVal A(,) As Double)
        Dim nRow As Integer = A.GetLength(0)
        Dim nCol As Integer = A.GetLength(1)
        Dim C(0 To nRow - 1, 0 To nCol - 1) As Double
        For i = 0 To nRow - 1
            For j = 0 To nCol - 1
                C(i, j) = k * A(i, j)
            Next j
        Next i
        Return C
    End Function

    Function MMinus(ByVal A(,) As Double, ByVal B(,) As Double)
        Dim Second(,) As Double = MScalar(-1, B)
        Dim Result(,) As Double = MAdd(A, Second)
        Return Result
    End Function

    Function MMul(ByVal A(,) As Double, ByVal B(,) As Double)
        If (A.GetLength(1) <> B.GetLength(0)) Then
            MsgBox("The inner product operation is not commutable since the A column size is " & A.GetLength(1) & _
                   " ,but the B row size is " & B.GetLength(0) & ".")
            End
        End If
        Dim nRow As Integer = A.GetLength(0)
        Dim nCol As Integer = B.GetLength(1)
        Dim nSum As Integer = A.GetLength(1)
        Dim sum As Double
        Dim C(0 To nRow - 1, 0 To nCol - 1) As Double
        For i = 0 To nRow - 1
            For j = 0 To nCol - 1
                sum = 0
                For k = 0 To nSum - 1
                    sum += A(i, k) * B(k, j)
                Next k
                C(i, j) = sum
            Next j
        Next i
        Return C
    End Function

    Function MIdentity(ByVal dimen As Integer)
        Dim C(0 To dimen - 1, 0 To dimen - 1) As Double
        For i = 0 To dimen - 1
            For j = 0 To dimen - 1
                If (i = j) Then
                    C(i, j) = 1.0
                Else
                    C(i, j) = 0.0
                End If
            Next j
        Next i
        Return C
    End Function

    Function MMake(ByVal value As Double, ByVal nRow As Integer, ByVal nCol As Integer)
        Dim C(0 To nRow - 1, 0 To nCol - 1) As Double
        For i = 0 To nRow - 1
            For j = 0 To nCol - 1
                C(i, j) = value
            Next j
        Next i
        Return C
    End Function

    Function MLU(ByVal A(,) As Double)
        If (A.GetLength(0) <> A.GetLength(1)) Then
            MsgBox("This function can make LU Decomposition for only square matrix.")
            End
        End If
        Dim length As Integer = A.GetLength(0)
        Dim L(0 To length - 1, 0 To length - 1) As Double
        Dim U(0 To length - 1, 0 To length - 1) As Double
        For i = 0 To length - 1
            For j = 0 To length - 1
                If (i < j) Then
                    L(i, j) = 0
                ElseIf (i = j) Then
                    L(i, j) = 1
                End If
            Next j
        Next i
        For i = 0 To length - 1
            For j = 0 To length - 1
                If (i > j) Then
                    U(i, j) = 0
                End If
            Next j
        Next i
        Dim num As Integer
        Dim sum As Double
        For j = 0 To length - 1
            For i = 0 To length - 1
                If (i <= j) Then
                    num = i
                    sum = 0
                    If num > 0 Then
                        For k = 0 To num - 1
                            sum += L(i, k) * U(k, j)
                        Next k
                    End If
                    U(i, j) = (A(i, j) - sum) / L(num, num)
                Else
                    num = j
                    sum = 0
                    If num > 0 Then
                        For k = 0 To num - 1
                            sum += L(i, k) * U(k, j)
                        Next k
                    End If
                    L(i, j) = (A(i, j) - sum) / U(num, num)
                End If
            Next i
        Next j
        Dim Result(,) As Double = MCAppend(L, U)
        Return Result
    End Function

    Function MDet(ByVal A(,) As Double)
        If (A.GetLength(0) <> A.GetLength(1)) Then
            MsgBox("This function can make determinant for only square matrix.")
            End
        End If
        Dim LU(,) As Double = MLU(A)
        Dim length As Integer = A.GetLength(0)
        Dim det As Double = 1
        For i = 0 To length - 1
            det *= LU(i, length + i)
        Next i
        Return det
    End Function

    Function MInverse(ByVal A(,) As Double)
        If (A.GetLength(0) <> A.GetLength(1)) Then
            MsgBox("This function can make inverse for only square matrix.")
            End
        End If
        Dim LU(,) As Double = MLU(A)
        Dim length As Integer = A.GetLength(0)
        Dim L(0 To length - 1, 0 To length - 1), U(0 To length - 1, 0 To length - 1) As Double
        Dim i, j As Integer
        For i = 0 To length - 1
            For j = 0 To length - 1
                L(i, j) = LU(i, j)
                U(i, j) = LU(i, length + j)
            Next j
        Next i
        Dim det As Double = 1
        For i = 0 To length - 1
            det *= U(i, i)
        Next i
        If (Math.Abs(det) < 0.00000000000001) Then
            MsgBox("This matrix cannot find inverse because det = 0")
            End
        End If
        Dim Iden(,) As Double = MIdentity(length)
        Dim invL(,) As Double = MGaussJforL(L, Iden)
        Dim invA(,) As Double = MGaussJForU(U, invL)
        Return invA
    End Function

    Function MRAppend(ByVal A(,) As Double, ByVal B(,) As Double)
        If (A.GetLength(1) <> B.GetLength(1)) Then
            MsgBox("Row append function can be made for only matricies with equal columns.")
            End
        End If
        Dim nCol As Integer = A.GetLength(1)
        Dim nRow1 As Integer = A.GetLength(0)
        Dim nRow2 As Integer = B.GetLength(0)
        Dim C(0 To nRow1 + nRow2 - 1, 0 To nCol - 1) As Double
        For i = 0 To nRow1 - 1
            For j = 0 To nCol - 1
                C(i, j) = A(i, j)
            Next j
        Next i
        For i = 0 To nRow2 - 1
            For j = 0 To nCol - 1
                C(nRow1 + i, j) = B(i, j)
            Next j
        Next i
        Return C
    End Function

    Function MCAppend(ByVal A(,) As Double, ByVal B(,) As Double)
        If (A.GetLength(0) <> B.GetLength(0)) Then
            MsgBox("Row append function can be made for only matricies with equal rows.")
            End
        End If
        Dim nCol1 As Integer = A.GetLength(1)
        Dim nCol2 As Integer = B.GetLength(1)
        Dim nRow As Integer = A.GetLength(0)
        Dim C(0 To nRow - 1, nCol1 + nCol2 - 1) As Double
        For i = 0 To nRow - 1
            For j = 0 To nCol1 - 1
                C(i, j) = A(i, j)
            Next j
            For j = 0 To nCol2 - 1
                C(i, nCol1 + j) = B(i, j)
            Next j
        Next i
        Return C
    End Function

    Private Function MGaussJforL(ByVal A(,) As Double, ByVal B(,) As Double)
        Dim length As Integer = A.GetLength(0)
        Dim C(,) As Double = MCAppend(A, B)
        For j = 0 To length - 1
            Rconstant(C, j, 1 / C(j, j))
            For i = length - 1 To j + 1 Step -1
                ROperation(C, i, j, C(i, j))
            Next i
        Next j
        Dim Result(0 To length - 1, 0 To length - 1) As Double
        For i = 0 To length - 1
            For j = 0 To length - 1
                Result(i, j) = C(i, length + j)
            Next j
        Next i
        Return (Result)
    End Function

    Private Function MGaussJForU(ByVal A(,) As Double, ByVal B(,) As Double)
        Dim length As Integer = A.GetLength(0)
        Dim C(,) As Double = MCAppend(A, B)
        For j = length - 1 To 0 Step -1
            Rconstant(C, j, 1 / C(j, j))
            For i = 0 To j - 1
                ROperation(C, i, j, C(i, j))
            Next i
        Next j
        Dim Result(0 To length - 1, 0 To length - 1) As Double
        For i = 0 To length - 1
            For j = 0 To length - 1
                Result(i, j) = C(i, length + j)
            Next j
        Next i
        Return (Result)
    End Function

    Private Sub ROperation(ByRef Matrix(,) As Double, ByVal targetRow As Integer, ByVal operRow As Integer, ByVal k As Double)
        Dim length As Integer = Matrix.GetLength(1)
        For i = 0 To length - 1
            Matrix(targetRow, i) = Matrix(targetRow, i) - (k * Matrix(operRow, i))
        Next i
    End Sub

    Private Sub Rconstant(ByRef Matrix(,) As Double, ByVal targetRow As Integer, ByVal k As Double)
        For i = 0 To Matrix.GetLength(1) - 1
            Matrix(targetRow, i) = k * Matrix(targetRow, i)
        Next
    End Sub

    Function IsDoubleInRange(ByVal textbox As TextBox, ByVal LBound As Double, ByVal UBound As Double, ByVal Name As String) As Boolean
        Dim temp As Double
        Try
            temp = Convert.ToDecimal(textbox.Text)
        Catch ex As Exception
            MsgBox(Name & " is required field and must be a decimal value.")
            Return False
        End Try
        temp = Convert.ToDecimal(textbox.Text)
        If (temp > UBound Or temp < LBound) Then
            MsgBox(Name & " must be in range of " & LBound.ToString("f3") & " to " & UBound.ToString("f3") & ".")
            Return False
        Else
            Return True
        End If
    End Function

    Function IsIntegerLBound(ByVal textbox As TextBox, ByVal LBound As Integer, ByVal Name As String) As Boolean
        Dim temp As Integer
        Try
            temp = Convert.ToInt32(textbox.Text)
            If temp < LBound Then
                MsgBox(Name & " cannot be less than " & LBound & ".")
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(Name & " is required field and must be integer.")
            Return False
        End Try
    End Function

    Function IsDoubleLBound(ByVal textbox As TextBox, ByVal LBound As Double, ByVal Name As String) As Boolean
        Dim temp As Double
        Try
            temp = Convert.ToDecimal(textbox.Text)
            If temp < LBound Then
                MsgBox(Name & " cannot be less than " & LBound.ToString("f3") & ".")
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(Name & " is required field and must be number in decimals.")
            Return False
        End Try
    End Function

    Public Function NormsDist(ByVal x As Double) As Double

        Dim t As Double
        Const b1 = 0.31938153
        Const b2 = -0.356563782
        Const b3 = 1.781477937
        Const b4 = -1.821255978
        Const b5 = 1.330274429
        Const p = 0.2316419
        Const c = 0.39894228

        If x >= 0 Then
            t = 1.0# / (1.0# + p * x)
            NormsDist = (1.0# - c * Math.Exp(-x * x / 2.0#) * t * _
            (t * (t * (t * (t * b5 + b4) + b3) + b2) + b1))
        Else
            t = 1.0# / (1.0# - p * x)
            NormsDist = (c * Math.Exp(-x * x / 2.0#) * t * _
            (t * (t * (t * (t * b5 + b4) + b3) + b2) + b1))
        End If
        'Reference http://www.sitmo.com/doc/Calculating_the_Cumulative_Normal_Distribution
        'Accurate for 6 decimal places
    End Function

    Public Function NormsInv(ByVal x As Double)
        Dim result, temp1, temp2, temp3 As Double
        Dim z As Double = 0
        For i = 1 To 8
            temp1 = -1
            temp2 = -1
            temp3 = 1
            Do While CInt(temp1 * (10 ^ i)) <> CInt(temp3 * (10 ^ i))
                result = NormsDist(z)
                temp1 = temp2
                temp2 = temp3
                temp3 = z
                If (result - x > 0) Then
                    z = z - (0.1 ^ i)
                Else
                    z = z + (0.1 ^ i)
                End If
            Loop
        Next i
        Return z
        'Accurate for only 4 decimal places
    End Function

    Public Function FindVar(ByVal nT As Integer, ByVal nC As Integer, ByVal nIndiv As Integer, ByVal residualBetween As Double, ByVal residualWithin As Double) As Double
        'Dim INV() As Decimal = MInverseV(residualBetween, residualWithin, nIndiv)
        'Dim f As Decimal = INV(1)
        'Dim d As Decimal = INV(0) - INV(1)

        Dim J As Decimal = nT + nC
        Dim VarGamma11 As Decimal = (residualWithin + (nIndiv * residualBetween)) * J / (nIndiv * nT * nC)

        'Dim VarGamma11 As Decimal = J / (nIndiv * nT * nC * (d + (nIndiv * f)))
        'Dim V(,) As Double = Identity(nIndiv)
        'MScalar(residualWithin, V)
        'Dim T(,) As Double = MakeM(nIndiv, nIndiv, residualBetween)
        'V = MAdd(V, T)
        'Dim InvV(,) As Double = MInverse(V)
        'Dim f As Double = InvV(1, 0)
        'Dim d As Double = InvV(0, 0) - f
        'Dim J As Double = nT + nC
        'Dim VarGamma11 As Double = J / (nIndiv * nT * nC * (d + (nIndiv * f)))
        'Dim InvVarGamma(0 To 1, 0 To 1) As Double
        'InvVarGamma(0, 0) = J * nIndiv * (d + (nIndiv * f))
        'InvVarGamma(1, 0) = InvVarGamma(0, 1) = InvVarGamma(1, 1) = nT * nIndiv * (d + (nIndiv * f))
        'If covariate = True Then
        '    Dim InvVarGammaC(,) As Double = MakeM(4, 4, 0.0)
        '    InvVarGammaC(0, 0) = InvVarGamma(0, 0)
        '    InvVarGammaC(1, 0) = InvVarGammaC(0, 1) = InvVarGammaC(1, 1) = InvVarGamma(1, 0)
        '    InvVarGammaC(2, 2) = d * SSWZ
        '    InvVarGammaC(3, 3) = nIndiv * (d + (nIndiv * f)) * SSBZ
        '    Dim VarGamma(,) As Double = MInverse(InvVarGammaC)
        '    Return VarGamma(1, 1)
        'Else
        '    Dim VarGamma(,) As Double = MInverse(InvVarGamma)
        '    Return VarGamma(1, 1)
        'End If
        Return CDbl(VarGamma11)
    End Function
    Public Function FindVar(ByVal J As Integer, ByVal pT As Double, ByVal nIndiv As Integer, ByVal residualBetween As Double, ByVal residualWithin As Double) As Double
        'Dim INV() As Decimal = MInverseV(residualBetween, residualWithin, nIndiv)
        'Dim f As Decimal = INV(1)
        'Dim d As Double = INV(0) - INV(1)
        'Dim VarGamma11 As Decimal = J / (nIndiv * J * pT * J * (1 - pT) * (d + (nIndiv * f)))

        Dim VarGamma11 As Decimal = (residualWithin + (nIndiv * residualBetween)) / (nIndiv * J * pT * (1 - pT))

        'Dim V(,) As Double = Identity(nIndiv)
        'MScalar(residualWithin, V)
        'Dim T(,) As Double = MakeM(nIndiv, nIndiv, residualBetween)
        'V = MAdd(V, T)
        'Dim InvV(,) As Double = MInverse(V)
        'Dim f As Double = InvV(1, 0)
        'Dim d As Double = InvV(0, 0) - f
        'Dim VarGamma11 As Double = 1 / (nIndiv * J * pT * (1 - pT) * (d + (nIndiv * f)))
        'Dim InvVarGamma(0 To 1, 0 To 1) As Double
        'InvVarGamma(0, 0) = J * nIndiv * (d + (nIndiv * f))
        'InvVarGamma(1, 0) = InvVarGamma(0, 1) = InvVarGamma(1, 1) = nT * nIndiv * (d + (nIndiv * f))
        'If covariate = True Then
        '    Dim InvVarGammaC(,) As Double = MakeM(4, 4, 0.0)
        '    InvVarGammaC(0, 0) = InvVarGamma(0, 0)
        '    InvVarGammaC(1, 0) = InvVarGammaC(0, 1) = InvVarGammaC(1, 1) = InvVarGamma(1, 0)
        '    InvVarGammaC(2, 2) = d * SSWZ
        '    InvVarGammaC(3, 3) = nIndiv * (d + (nIndiv * f)) * SSBZ
        '    Dim VarGamma(,) As Double = MInverse(InvVarGammaC)
        '    Return VarGamma(1, 1)
        'Else
        '    Dim VarGamma(,) As Double = MInverse(InvVarGamma)
        '    Return VarGamma(1, 1)
        'End If
        Return CDbl(VarGamma11)
    End Function
    Public Function FindPower(ByVal ES As Double, ByVal Var As Double) As Double
        Dim CritValue As Double = InvNormalDistribution(0.975)
        Dim z As Double = Math.Abs(ES / Math.Sqrt(Var))
        Dim ValueH1 As Double = CritValue - z
        Return 1 - NormalDistribution(ValueH1)
    End Function

    Public Function FindWidthCIES(ByVal Var As Double, Optional ByVal WidthLevel As Double = 0.95) As Double
        Dim ProportionFromLeft As Double = 1 - ((1 - WidthLevel) / 2)
        Dim CritValue As Double = InvNormalDistribution(ProportionFromLeft)
        Return 2 * (CritValue * Math.Sqrt(Var))
    End Function

    Public Function FindVarForPower(ByVal ES As Double, ByVal PowerLevel As Double) As Double
        Dim CritValue As Double = InvNormalDistribution(0.975)
        Dim PowerValue As Double = InvNormalDistribution(1 - PowerLevel)
        Dim z As Double = CritValue - PowerValue
        Return (ES / z) ^ 2
    End Function

    Public Function FindVarForWidth(ByVal Width As Double, ByVal WidthLevel As Double) As Double
        Dim SideEdge As Double = (1 - WidthLevel) / 2
        Dim CritValue As Double = InvNormalDistribution(1 - SideEdge)
        Return (Width / (2 * CritValue)) ^ 2
    End Function

    Public Function FindTotalCostExact(ByVal nT As Integer, ByVal nC As Integer, ByVal nIndiv As Integer, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Return (nT * (TGroupcost + (nIndiv * TIndivCost))) + (nC * (CGroupCost + (nIndiv * CIndivCost)))
    End Function

    Public Function FindTotalCostExact(ByVal nJ As Integer, ByVal pT As Double, ByVal nIndiv As Integer, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Return nJ * ((pT * (TGroupcost + (nIndiv * TIndivCost))) + ((1 - pT) * (CGroupCost + (nIndiv * CIndivCost))))
    End Function

    Public Function FindNIndiv_PJ_Var(ByVal covariate As Boolean, ByVal pT As Double, ByVal J As Integer, ByVal VarCriterion As Double, ByVal ES As Double, ByVal ICCY As Double, Optional ByVal ICCZ As Double = 0, Optional ByVal R2ClusterZ As Double = 0, Optional ByVal R2IndivZ As Double = 0) As Integer
        Dim NIndiv As Integer = 100
        Dim residual() As Double = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
        Dim CurrentVar As Double = FindVar(J, pT, NIndiv, residual(0), residual(1))
        If (CurrentVar < VarCriterion) Then 'NIndiv too large
            Do While CurrentVar < VarCriterion
                NIndiv -= 10
                residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                If (residual(0) <= 0 Or residual(1) <= 0) Then Return 100000
                CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
                If (NIndiv <= 10) Then Exit Do
            Loop
            If (CurrentVar > VarCriterion) Then
                Do While CurrentVar > VarCriterion
                    NIndiv += 1
                    residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
                Loop
                Return NIndiv
            Else
                Do While CurrentVar < VarCriterion
                    If (NIndiv < MinIndiv) Then Return MinIndiv
                    NIndiv -= 1
                    residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
                Loop
                Return NIndiv + 1
            End If

        ElseIf (CurrentVar > VarCriterion) Then 'NIndiv too small
            Do While CurrentVar > VarCriterion
                NIndiv += 10000
                residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
                If (NIndiv >= 100000) Then Return 100000
            Loop
            Do While CurrentVar < VarCriterion
                NIndiv -= 1000
                residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
            Loop
            Do While CurrentVar > VarCriterion
                NIndiv += 100
                residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
            Loop
            Do While CurrentVar < VarCriterion
                NIndiv -= 10
                residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
            Loop
            Do While CurrentVar > VarCriterion
                NIndiv += 1
                residual = FindResidual(covariate, pT, J, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, NIndiv, residual(0), residual(1))
            Loop
            Return NIndiv
        Else
            Return NIndiv
        End If
    End Function

    Public Function FindNIndiv_PJ_TotalCostExact(ByVal pT As Double, ByVal J As Integer, ByVal TotalCost As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Dim AverageIndivCost = (pT * TIndivCost) + ((1 - pT) * CIndivCost)
        Dim AverageGroupCost = (pT * TGroupcost) + ((1 - pT) * CGroupCost)
        Dim EachGroupCost = TotalCost / J
        Dim result As Integer = Math.Truncate((EachGroupCost - AverageGroupCost) / AverageIndivCost)
        If (result < MinIndiv) Then
            Return -1
        Else
            Return result
        End If
    End Function

    Public Function FindNIndiv_PJ_TotalCostMax(ByVal pT As Double, ByVal J As Integer, ByVal TotalCost As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Dim AverageIndivCost1 = (pT * TIndivCost) + ((1 - pT) * CIndivCost)
        Dim AverageGroupCost1 = (pT * TGroupcost) + ((1 - pT) * CGroupCost)
        Dim EachGroupCost1 = TotalCost / J
        Dim nIndiv1 As Integer = Math.Truncate((EachGroupCost1 - AverageGroupCost1) / AverageIndivCost1)
        Dim AverageIndivCost2 = ((1 - pT) * TIndivCost) + (pT * CIndivCost)
        Dim AverageGroupCost2 = ((1 - pT) * TGroupcost) + (pT * CGroupCost)
        Dim EachGroupCost2 = TotalCost / J
        Dim nIndiv2 As Integer = Math.Truncate((EachGroupCost2 - AverageGroupCost2) / AverageIndivCost2)
        If (nIndiv1 < MinIndiv) And (nIndiv2 < MinIndiv) Then Return -1
        If nIndiv1 >= nIndiv2 Then
            Return nIndiv1
        Else
            Return nIndiv2
        End If
    End Function

    Public Function FindJ_PNIndiv_Var(ByVal covariate As Boolean, ByVal pT As Double, ByVal nIndiv As Integer, ByVal VarCriterion As Double, ByVal ES As Double, ByVal ICCY As Double, ByVal ICCZ As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double)
        Dim J As Integer = 50
        Dim residual() As Double = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
        Dim CurrentVar As Double = FindVar(J, pT, nIndiv, residual(0), residual(1))
        If (CurrentVar < VarCriterion) Then 'NIndiv too large
            Do While CurrentVar < VarCriterion And J > 10
                J -= 10
                residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
            Loop
            If (CurrentVar > VarCriterion) Then
                Do While CurrentVar > VarCriterion
                    J += 1
                    residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
                Loop
                Return J
            Else
                Do While CurrentVar < VarCriterion
                    J -= 1
                    residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
                    If (J < 4) Then Return 4
                Loop
                Return J + 1
            End If
        ElseIf (CurrentVar > VarCriterion) Then 'NIndiv too small
            Do While CurrentVar > VarCriterion
                J += 5000
                residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
                If (J >= 20000) Then Return 20000
            Loop
            Do While CurrentVar < VarCriterion And J > 10
                J -= 1000
                residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
            Loop
            Do While CurrentVar > VarCriterion
                J += 100
                residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
            Loop
            Do While CurrentVar < VarCriterion And J > 10
                J -= 10
                residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
            Loop
            Do While CurrentVar > VarCriterion
                J += 1
                residual = FindResidual(covariate, pT, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(J, pT, nIndiv, residual(0), residual(1))
            Loop
            Return J
        Else
            Return J
        End If
    End Function
    Public Function FindJ_PNIndiv_TotalCostExact(ByVal pT As Double, ByVal nIndiv As Integer, ByVal TotalCost As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Dim EachGroupCost = (pT * (TGroupcost + (nIndiv * TIndivCost))) + ((1 - pT) * (CGroupCost + (nIndiv * CIndivCost)))
        Dim result As Integer = Math.Truncate(TotalCost / EachGroupCost)
        If result < 4 Then
            Return -1
        Else
            Return result
        End If
    End Function

    Public Function FindJ_PNIndiv_TotalCostMax(ByVal pT As Double, ByVal nIndiv As Integer, ByVal TotalCost As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Dim EachGroupCost1 = (pT * (TGroupcost + (nIndiv * TIndivCost))) + ((1 - pT) * (CGroupCost + (nIndiv * CIndivCost)))
        Dim J1 As Integer = Math.Truncate(TotalCost / EachGroupCost1)
        Dim EachGroupCost2 = ((1 - pT) * (TGroupcost + (nIndiv * TIndivCost))) + (pT * (CGroupCost + (nIndiv * CIndivCost)))
        Dim J2 As Integer = Math.Truncate(TotalCost / EachGroupCost2)
        If (J1 < 4) And (J2 < 4) Then Return -1
        If J1 >= J2 Then
            Return J1
        Else
            Return J2
        End If
    End Function

    Public Function FindNT_NIndivJ_Var(ByVal covariate As Boolean, ByVal J As Integer, ByVal nIndiv As Integer, ByVal VarCriterion As Double, ByVal ES As Double, ByVal ICCY As Double, ByVal ICCZ As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double) As Integer
        Dim startNT As Integer
        If J Mod 2 = 1 Then
            startNT = Math.Truncate(J / 2) + 1
        Else
            startNT = J / 2
        End If
        Dim startNC As Integer = J - startNT
        Dim residual() As Double = FindResidual(covariate, CDbl(startNT) / J, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
        Dim CurrentVar As Double = FindVar(startNT, startNC, nIndiv, residual(0), residual(1))
        If (CurrentVar <= VarCriterion) Then 'p too far from .50
            Return startNT
        Else 'p too close to .50
            Do While CurrentVar > VarCriterion
                startNT += 1
                startNC -= 1
                If (startNT = J) Then
                    Return startNT - 1
                End If
                residual = FindResidual(covariate, CDbl(startNT) / J, J, nIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                CurrentVar = FindVar(startNT, startNC, nIndiv, residual(0), residual(1))
            Loop
            Return startNT - 1
        End If
    End Function

    Public Function FindNT_NIndivJ_TotalCostExact(ByVal J As Integer, ByVal nIndiv As Integer, ByVal TotalCost As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Integer
        Dim EachControlGroupCost As Double = CGroupCost + (CIndivCost * nIndiv)
        Dim GroupCostDiff As Double = TGroupcost - CGroupCost
        Dim IndivCostDiff As Double = TIndivCost - CIndivCost
        Dim Numerator As Double = TotalCost - (J * EachControlGroupCost)
        Dim Denominator As Double = GroupCostDiff + (nIndiv * IndivCostDiff)
        Return Math.Round(Numerator / Denominator)
        ' May have problem since not convert pT for the greater cost
    End Function

    Public Function FindNT_NIndivJ_TotalCostMax(ByVal J As Integer, ByVal nIndiv As Integer, ByVal TotalCost As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Integer
        Dim EachControlGroupCost As Double = CGroupCost + (CIndivCost * nIndiv)
        Dim GroupCostDiff As Double = TGroupcost - CGroupCost
        Dim IndivCostDiff As Double = TIndivCost - CIndivCost
        Dim Numerator As Double = TotalCost - (J * EachControlGroupCost)
        Dim Denominator As Double = GroupCostDiff + (nIndiv * IndivCostDiff)
        Dim nT1 = Math.Round(Numerator / Denominator)
        Dim EachControlGroupCostR As Double = TGroupcost + (TIndivCost * nIndiv)
        Dim GroupCostDiffR As Double = CGroupCost - TGroupcost
        Dim IndivCostDiffR As Double = CIndivCost - TIndivCost
        Dim NumeratorR As Double = TotalCost - (J * EachControlGroupCostR)
        Dim DenominatorR As Double = GroupCostDiffR + (nIndiv * IndivCostDiffR)
        Dim nT2 = Math.Round(NumeratorR / DenominatorR)
        Dim VarNT1 As Double = (CDbl(nT1) / J) * ((J - CDbl(nT1)) / J)
        Dim VarNT2 As Double = (CDbl(nT2) / J) * ((J - CDbl(nT2)) / J)
        If (nT1 >= J) And (nT2 >= J) Then Return -1
        If (nT1 >= J) Then Return nT2
        If (nT2 >= J) Then Return nT1
        If VarNT1 >= VarNT2 Then
            Return nT2
        Else
            Return nT1
        End If
    End Function

    Public Function FindTotalCostRevertible(ByVal nT As Integer, ByVal nC As Integer, ByVal nIndiv As Integer, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Dim cost1 = (nT * (TGroupcost + (nIndiv * TIndivCost))) + (nC * (CGroupCost + (nIndiv * CIndivCost)))
        Dim cost2 = (nC * (TGroupcost + (nIndiv * TIndivCost))) + (nT * (CGroupCost + (nIndiv * CIndivCost)))
        If cost1 >= cost2 Then
            Return cost2
        Else
            Return cost1
        End If
    End Function

    Public Function FindTotalCostRevertible(ByVal nJ As Integer, ByVal pT As Double, ByVal nIndiv As Integer, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double) As Double
        Dim cost1 = nJ * ((pT * (TGroupcost + (nIndiv * TIndivCost))) + ((1 - pT) * (CGroupCost + (nIndiv * CIndivCost))))
        Dim cost2 = nJ * (((1 - pT) * (TGroupcost + (nIndiv * TIndivCost))) + (pT * (CGroupCost + (nIndiv * CIndivCost))))
        If cost1 >= cost2 Then
            Return cost2
        Else
            Return cost1
        End If
    End Function

    Public Function FindOption1(ByVal covariate As Boolean, ByVal DesiredVar As Double, ByVal ES As Double, ByVal ICCY As Double, Optional ByVal ICCZ As Double = 0, Optional ByVal R2ClusterZ As Double = 0, Optional ByVal R2IndivZ As Double = 0, Optional ByVal pT2 As Double = -1)
        Dim StartJ() As Integer = {10, 30, 50, 70, 90, 110, 130, 150, 170, 190}
        Dim nIndiv(9) As Integer
        Dim TNIndiv(9) As Double
        Dim result(0 To 3) As Double 'NT NC And NINDIV
        Dim pT As Double
        If pT2 = -1 Then
            pT = 0.5
        Else
            pT = pT2
        End If
        For i = 0 To 9
            nIndiv(i) = FindNIndiv_PJ_Var(covariate, pT, StartJ(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
            TNIndiv(i) = CDbl(StartJ(i)) * nIndiv(i)
        Next i
        Dim FindNewJ As Boolean = AllElementEqual(nIndiv, 100000)
        Dim MinPosition As Integer = FindMinimum(TNIndiv)
        If (FindNewJ = False) AndAlso MinPosition = 0 Then
            Dim Range As Integer = StartJ(1) - StartJ(0)
            Dim JScope(0 To Range) As Integer
            For i = 0 To Range
                JScope(i) = StartJ(0) + i
            Next i
            Dim nIndiv2(0 To Range) As Integer
            Dim TNIndiv2(0 To Range) As Double
            For i = 0 To Range
                nIndiv2(i) = FindNIndiv_PJ_Var(covariate, pT, JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                TNIndiv2(i) = CDbl(JScope(i)) * nIndiv2(i)
            Next
            Dim MinPosition2 As Integer = FindMinimum(TNIndiv2)
            If (MinPosition2 > 0) Then
                result(0) = Math.Round(JScope(MinPosition2) * pT)
                result(1) = Math.Round(JScope(MinPosition2) * (1 - pT))
                result(2) = nIndiv2(MinPosition2)
                result(3) = pT
                If (JScope(MinPosition2) > (result(0) + result(1))) Then
                    result(0) += 1
                ElseIf (JScope(MinPosition2) < (result(0) + result(1))) Then
                    result(1) -= 1
                End If
            Else
                Dim nTnC(,) As Integer = pJ11
                Dim nIndiv3(0 To 35) As Integer
                Dim pNTNC(0 To 35) As Double
                Dim TNIndiv3(0 To 35) As Double
                For i = 0 To 35
                    pNTNC(i) = CDbl(nTnC(i, 0)) / (nTnC(i, 0) + nTnC(i, 1))
                    If (pT2 = -1) Then
                        nIndiv3(i) = FindNIndiv_PJ_Var(covariate, pNTNC(i), nTnC(i, 0) + nTnC(i, 1), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    Else
                        If Math.Abs(pNTNC(i) - pT) <= 0.1 Then
                            nIndiv3(i) = FindNIndiv_PJ_Var(covariate, pNTNC(i), nTnC(i, 0) + nTnC(i, 1), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                        Else
                            nIndiv3(i) = 100000
                        End If
                    End If
                    TNIndiv3(i) = CDbl(nTnC(i, 0) + nTnC(i, 1)) * nIndiv3(i)
                Next
                Dim MinPosition3 As Integer = FindMinimum(TNIndiv3)
                result(0) = nTnC(MinPosition3, 0)
                result(1) = nTnC(MinPosition3, 1)
                result(2) = nIndiv3(MinPosition3)
                result(3) = pNTNC(MinPosition3)
            End If

        ElseIf (FindNewJ = False) AndAlso MinPosition < 9 Then
            Dim Range As Integer = StartJ(MinPosition + 1) - StartJ(MinPosition - 1)
            Dim JScope(0 To Range) As Integer
            For i = 0 To Range
                JScope(i) = StartJ(MinPosition - 1) + i
            Next i
            Dim nIndiv2(0 To Range) As Integer
            Dim TNIndiv2(0 To Range) As Double
            For i = 0 To Range
                nIndiv2(i) = FindNIndiv_PJ_Var(covariate, pT, JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                TNIndiv2(i) = CDbl(JScope(i)) * nIndiv2(i)
            Next
            Dim MinPosition2 As Integer = FindMinimum(TNIndiv2)
            result(0) = Math.Round(JScope(MinPosition2) * pT)
            result(1) = Math.Round(JScope(MinPosition2) * (1 - pT))
            result(2) = nIndiv2(MinPosition2)
            result(3) = pT
            If (JScope(MinPosition2) > (result(0) + result(1))) Then
                result(0) += 1
            ElseIf (JScope(MinPosition2) < (result(0) + result(1))) Then
                result(1) -= 1
            End If
        Else
            Dim RunJ() As Integer = {170, 190, 210}
            Dim RunNIndiv(0 To 2) As Integer
            Dim RunTNIndiv(0 To 2) As Double
            For i = 0 To 2
                RunNIndiv(i) = FindNIndiv_PJ_Var(covariate, pT, RunJ(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                RunTNIndiv(i) = CDbl(RunJ(i)) * RunNIndiv(i)
            Next i
            Do While (AllElementEqual(RunNIndiv, 100000) = True) OrElse (FindMinimum(RunTNIndiv) <> 1)
                RunJ(0) = RunJ(1)
                RunJ(1) = RunJ(2)
                RunJ(2) = RunJ(2) + 20
                RunNIndiv(0) = RunNIndiv(1)
                RunNIndiv(1) = RunNIndiv(2)
                RunNIndiv(2) = 0
                RunNIndiv(2) = FindNIndiv_PJ_Var(covariate, pT, RunJ(2), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                RunTNIndiv(0) = RunTNIndiv(1)
                RunTNIndiv(1) = RunTNIndiv(2)
                RunTNIndiv(2) = CDbl(RunJ(2)) * RunNIndiv(2)
            Loop
            Dim Range As Integer = RunJ(2) - RunJ(0)
            Dim JScope(0 To Range) As Integer
            For i = 0 To Range
                JScope(i) = RunJ(0) + i
            Next i
            Dim nIndiv2(0 To Range) As Integer
            Dim TNIndiv2(0 To Range) As Double
            For i = 0 To Range
                nIndiv2(i) = FindNIndiv_PJ_Var(covariate, pT, JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                TNIndiv2(i) = CDbl(JScope(i)) * nIndiv2(i)
            Next
            Dim MinPosition2 As Integer = FindMinimum(TNIndiv2)
            result(0) = Math.Round(JScope(MinPosition2) * pT)
            result(1) = Math.Round(JScope(MinPosition2) * (1 - pT))
            result(2) = nIndiv2(MinPosition2)
            result(3) = pT
            If (JScope(MinPosition2) > (result(0) + result(1))) Then
                result(0) += 1
            ElseIf (JScope(MinPosition2) < (result(0) + result(1))) Then
                result(1) -= 1
            End If
        End If
        Return result
    End Function

    Public Function FindMaximum(ByVal array() As Double)
        Dim Position As Integer = 0
        Dim value As Double = array(0)
        For i = 1 To UBound(array)
            If (array(i) = value) Then
                Position = -1
            End If
            If (array(i) > value) Then
                value = array(i)
                Position = i
            End If
        Next i
        If (Position = -1) Then
            For i = 0 To UBound(array)
                If (array(i) = value) Then
                    Return i
                End If
            Next
        End If
        Return Position
    End Function

    Public Function FindMaximum(ByVal array(,) As Double)
        Dim Position() As Integer = {0, 0}
        Dim value As Double = array(0, 0)
        For j = 1 To array.GetLength(1) - 1
            If (array(0, j) = value) Then
                Position(0) = -1
                Position(1) = -1
            End If
            If (array(0, j) > value) Then
                value = array(0, j)
                Position(0) = 0
                Position(1) = j
            End If

        Next j
        For i = 1 To array.GetLength(0) - 1
            For j = 0 To array.GetLength(1) - 1
                If (array(i, j) = value) Then
                    Position(0) = -1
                    Position(1) = -1
                End If
                If (array(i, j) > value) Then
                    value = array(i, j)
                    Position(0) = i
                    Position(1) = j
                End If
            Next j
        Next i
        If (Position(0) = -1) And (Position(1) = -1) Then
            For i = 0 To array.GetLength(0) - 1
                For j = 0 To array.GetLength(1) - 1
                    If (array(i, j) = value) Then
                        Dim Position2() As Integer = {i, j}
                        Return Position2
                    End If
                Next
            Next
        End If
        Return Position
    End Function

    Public Function FindMinimum(ByVal array() As Double)
        Dim Position As Integer = 0
        Dim value As Double = array(0)
        For i = 1 To UBound(array)
            If (array(i) = value) Then
                Position = -1
            End If
            If (array(i) < value) Then
                value = array(i)
                Position = i
            End If
        Next i
        If (Position = -1) Then
            For i = 0 To UBound(array)
                If (array(i) = value) Then
                    Return i
                End If
            Next
        End If
        Return Position
    End Function

    Public Function FindMinimum(ByVal array(,) As Double)
        Dim Position() As Integer = {0, 0}
        Dim value As Double = array(0, 0)
        For j = 1 To array.GetLength(1) - 1
            If (array(0, j) = value) Then
                Position(0) = -1
                Position(1) = -1
            End If
            If (array(0, j) < value) Then
                value = array(0, j)
                Position(0) = 0
                Position(1) = j
            End If

        Next j
        For i = 1 To array.GetLength(0) - 1
            For j = 0 To array.GetLength(1) - 1
                If (array(i, j) = value) Then
                    Position(0) = -1
                    Position(1) = -1
                End If
                If (array(i, j) < value) Then
                    value = array(i, j)
                    Position(0) = i
                    Position(1) = j
                End If

            Next j
        Next i
        If (Position(0) = -1) And (Position(1) = -1) Then
            For i = 0 To array.GetLength(0) - 1
                For j = 0 To array.GetLength(1) - 1
                    If (array(i, j) = value) Then
                        Dim Position2() As Integer = {i, j}
                        Return Position2
                    End If
                Next
            Next
        End If
        Return Position
    End Function

    Public Function FindMinimumCells(ByVal array() As Double)
        Dim Position As Integer = 0
        Dim value As Double = array(0)
        For i = 1 To UBound(array)
            If (array(i) = value) Then
                Position = -1
            End If
            If (array(i) < value) Then
                value = array(i)
                Position = i
            End If
        Next i
        Dim MinCells(0 To UBound(array)) As Integer
        For i = 0 To UBound(array)
            If (array(i) = value) Then
                MinCells(i) = 1
            Else
                MinCells(i) = 0
            End If
        Next
        Return MinCells
    End Function

    Public Function FindMinimumCells(ByVal array(,) As Double)
        Dim Position() As Integer = {0, 0}
        Dim value As Double = array(0, 0)
        For j = 1 To array.GetLength(1) - 1
            If (array(0, j) = value) Then
                Position(0) = -1
                Position(1) = -1
            End If
            If (array(0, j) < value) Then
                value = array(0, j)
                Position(0) = 0
                Position(1) = j
            End If

        Next j
        For i = 1 To array.GetLength(0) - 1
            For j = 0 To array.GetLength(1) - 1
                If (array(i, j) = value) Then
                    Position(0) = -1
                    Position(1) = -1
                End If
                If (array(i, j) < value) Then
                    value = array(i, j)
                    Position(0) = i
                    Position(1) = j
                End If

            Next j
        Next i
        Dim MinCells(0 To array.GetLength(0) - 1, 0 To array.GetLength(1) - 1) As Integer
        For i = 0 To array.GetLength(0) - 1
            For j = 0 To array.GetLength(1) - 1
                If (array(i, j) = value) Then
                    MinCells(i, j) = 1
                Else
                    MinCells(i, j) = 0
                End If
            Next
        Next
        Return MinCells
    End Function

    Public Function FindResidual(ByVal covariate As Boolean, ByVal pT As Double, ByVal J As Integer, ByVal NIndiv As Integer, ByVal ES As Double, ByVal ICCY As Double, Optional ByVal ICCZ As Double = 0, Optional ByVal R2ClusterZ As Double = 0, Optional ByVal R2IndivZ As Double = 0)
        Dim result(0 To 1) As Double 'Between and Within
        If covariate = True Then
            'Dim dfB As Integer = J - 1
            'Dim dfW As Integer = J * (NIndiv - 1)
            'Dim tauZ As Double = (dfB + dfW) / (dfB + (dfW * (1 - ICCZ) / ICCZ))
            'Dim sigmaZ As Double = tauZ * (1 - ICCZ) / ICCZ
            'Dim SSBZ As Double = dfB * tauZ
            'Dim SSWZ As Double = dfW * sigmaZ
            'Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
            'Dim BetaB As Double = TEZ + ((1 - etaZ) * CEZ)
            'Dim BetaW As Double = BetaB - CEZ
            'result(1) = 1 - ((BetaW ^ 2) * sigmaZ)
            'result(0) = (ICCY / (1 - ICCY)) - ((BetaB ^ 2) * tauZ)
            result(0) = (1 - R2ClusterZ) * (ICCY / (1 - ICCY))
            result(1) = 1 - R2IndivZ
        Else
            result(1) = 1
            result(0) = (ICCY / (1 - ICCY))
        End If
        Return result
    End Function

    Public Function FindOption2(ByVal covariate As Boolean, ByVal DesiredVar As Double, ByVal ES As Double, ByVal ICCY As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double, Optional ByVal ICCZ As Double = 0, Optional ByVal R2ClusterZ As Double = 0, Optional ByVal R2IndivZ As Double = 0, Optional ByVal pT2 As Double = -1, Optional ByVal NJ2 As Integer = -1, Optional ByVal NIndiv2 As Integer = -1)
        Dim StartJ() As Integer = {10, 30, 50, 70, 90, 110, 130, 150, 170, 190}
        Dim StartPT() As Double = {0.5, 0.6, 0.7, 0.8, 0.9, 0.95}
        Dim result(0 To 3) As Double 'NT NC And NINDIV


        If (pT2 <> -1) Then
            If (NIndiv2 <> -1) Then
                'P & NINDIV
                Dim pT As Double = pT2
                Dim nIndiv As Integer = NIndiv2
                Dim NJ As Integer = FindJ_PNIndiv_Var(covariate, pT, nIndiv, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                If NJ >= 20000 Then
                    result(0) = -1
                    result(1) = -1
                    result(2) = nIndiv
                    result(3) = pT
                    MsgBox("The algorithm was terminated because number of groups was increasing over 20000.")
                Else
                    result(0) = Math.Round(NJ * pT)
                    result(1) = Math.Round(NJ * (1 - pT))
                    result(2) = nIndiv
                    result(3) = pT
                    Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                    If (NJ > (result(0) + result(1))) Then
                        If (TCostLess = True) Then
                            result(0) += 1
                        Else
                            result(1) += 1
                        End If
                    ElseIf (NJ < (result(0) + result(1))) Then
                        If (TCostLess = True) Then
                            result(1) -= 1
                        Else
                            result(0) -= 1
                        End If
                    End If
                End If
            Else
                If (NJ2 <> -1) Then
                    'P & NG
                    Dim pT As Double = pT2
                    Dim NJ As Integer = NJ2
                    Dim NIndiv As Integer = FindNIndiv_PJ_Var(covariate, pT, NJ, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    If (NIndiv >= 100000) Then
                        result(0) = Math.Round(NJ * pT)
                        result(1) = Math.Round(NJ * (1 - pT))
                        result(2) = -1
                        result(3) = pT
                        MsgBox("The algorithm was terminated because number of individuals in each group over 100000.")
                    Else
                        result(0) = Math.Round(NJ * pT)
                        result(1) = Math.Round(NJ * (1 - pT))
                        result(2) = NIndiv
                        result(3) = pT
                        If (NJ > (result(0) + result(1))) Then
                            result(0) += 1
                        ElseIf (NJ < (result(0) + result(1))) Then
                            result(1) -= 1
                        End If
                    End If
                Else
                    'P only
                    Dim pT As Double = pT2
                    Dim NJ() As Integer = StartJ
                    Dim NIndiv(0 To UBound(NJ)) As Integer
                    Dim TCost(0 To UBound(NJ)) As Double
                    For i = 0 To UBound(NJ)
                        NIndiv(i) = FindNIndiv_PJ_Var(covariate, pT, NJ(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                        TCost(i) = FindTotalCostExact(NJ(i), pT, NIndiv(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                    Next i
                    Dim FindNewJ As Boolean = AllElementEqual(NIndiv, 100000)
                    Dim MinimumCost = FindMinimum(TCost)
                    If (FindNewJ = False) AndAlso (MinimumCost = 0) Then
                        Dim Range As Integer = NJ(1) - NJ(0)
                        Dim JScope(0 To Range) As Integer
                        For i = 0 To Range
                            JScope(i) = NJ(0) + i
                        Next i
                        Dim nIndiv4(0 To Range) As Integer
                        Dim TNIndiv4(0 To Range) As Double
                        For i = 0 To Range
                            nIndiv4(i) = FindNIndiv_PJ_Var(covariate, pT, JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TNIndiv4(i) = FindTotalCostExact(JScope(i), pT, nIndiv4(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next
                        Dim MinPosition2 As Integer = FindMinimum(TNIndiv4)
                        If (MinPosition2 > 0) Then
                            result(0) = Math.Round(JScope(MinPosition2) * pT)
                            result(1) = Math.Round(JScope(MinPosition2) * (1 - pT))
                            result(2) = nIndiv4(MinPosition2)
                            result(3) = pT
                            Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                            If (JScope(MinPosition2) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (JScope(MinPosition2) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            Dim nTnC(,) As Integer = pJ11
                            Dim nIndiv3(0 To 35) As Integer
                            Dim pNTNC(0 To 35) As Double
                            Dim TNIndiv3(0 To 35) As Double
                            For i = 0 To 35
                                pNTNC(i) = CDbl(nTnC(i, 0)) / (nTnC(i, 0) + nTnC(i, 1))
                                If Math.Abs(pNTNC(i) - pT) <= 0.1 Then
                                    nIndiv3(i) = FindNIndiv_PJ_Var(covariate, pNTNC(i), nTnC(i, 0) + nTnC(i, 1), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Else
                                    nIndiv3(i) = 100000
                                End If
                                TNIndiv3(i) = FindTotalCostExact(nTnC(i, 0), nTnC(i, 1), nIndiv3(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            Next
                            Dim MinPosition3 As Integer = FindMinimum(TNIndiv3)
                            result(0) = nTnC(MinPosition3, 0)
                            result(1) = nTnC(MinPosition3, 1)
                            result(2) = nIndiv3(MinPosition3)
                            result(3) = pNTNC(MinPosition3)
                        End If
                    ElseIf (FindNewJ = False) AndAlso (MinimumCost < UBound(NJ)) Then
                        Dim Range As Integer = NJ(MinimumCost + 1) - NJ(MinimumCost - 1)
                        Dim JScope(0 To Range) As Integer
                        For i = 0 To Range
                            JScope(i) = NJ(MinimumCost - 1) + i
                        Next i
                        Dim nIndiv3(0 To Range) As Integer
                        Dim TNIndiv3(0 To Range) As Double
                        For i = 0 To Range
                            nIndiv3(i) = FindNIndiv_PJ_Var(covariate, pT, JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TNIndiv3(i) = FindTotalCostExact(JScope(i), pT, nIndiv3(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next
                        Dim MinPosition3 As Integer = FindMinimum(TNIndiv3)
                        result(0) = Math.Round(JScope(MinPosition3) * pT)
                        result(1) = Math.Round(JScope(MinPosition3) * (1 - pT))
                        result(2) = nIndiv3(MinPosition3)
                        result(3) = pT
                        Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                        If (JScope(MinPosition3) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim RunJ() As Integer = {170, 190, 210}
                        Dim RunNIndiv(0 To 2) As Integer
                        Dim RunTNIndiv(0 To 2) As Double
                        For i = 0 To 2
                            RunNIndiv(i) = FindNIndiv_PJ_Var(covariate, pT, RunJ(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            RunTNIndiv(i) = FindTotalCostExact(RunJ(i), pT, RunNIndiv(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next i
                        Do While (AllElementEqual(RunNIndiv, 100000) = True) OrElse (FindMinimum(RunTNIndiv) <> 1)
                            RunJ(0) = RunJ(1)
                            RunJ(1) = RunJ(2)
                            RunJ(2) = RunJ(2) + 20
                            RunNIndiv(0) = RunNIndiv(1)
                            RunNIndiv(1) = RunNIndiv(2)
                            RunNIndiv(2) = FindNIndiv_PJ_Var(covariate, pT, RunJ(2), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            RunTNIndiv(0) = RunTNIndiv(1)
                            RunTNIndiv(1) = RunTNIndiv(2)
                            RunTNIndiv(2) = FindTotalCostExact(RunJ(2), pT, RunNIndiv(2), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Loop
                        Dim Range As Integer = RunJ(2) - RunJ(0)
                        Dim JScope(0 To Range) As Integer
                        For i = 0 To Range
                            JScope(i) = RunJ(0) + i
                        Next i
                        Dim nIndiv3(0 To Range) As Integer
                        Dim TNIndiv3(0 To Range) As Double
                        For i = 0 To Range
                            nIndiv3(i) = FindNIndiv_PJ_Var(covariate, pT, JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TNIndiv3(i) = FindTotalCostExact(JScope(i), pT, nIndiv3(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next
                        Dim MinPosition3 As Integer = FindMinimum(TNIndiv3)
                        result(0) = Math.Round(JScope(MinPosition3) * pT)
                        result(1) = Math.Round(JScope(MinPosition3) * (1 - pT))
                        result(2) = nIndiv3(MinPosition3)
                        result(3) = pT
                        Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                        If (JScope(MinPosition3) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If (NIndiv2 <> -1) Then
                If (NJ2 <> -1) Then
                    'NIndiv & NG
                    Dim NIndiv As Integer = NIndiv2
                    Dim NJ As Integer = NJ2
                    Dim residual() As Double = FindResidual(covariate, 0.5, NJ, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    Dim VarPossibleLower = FindVar(NJ, 0.5, NIndiv, residual(0), residual(1))
                    residual = FindResidual(covariate, 0.95, NJ, NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                    Dim VarPossibleUpper = FindVar(NJ, 0.95, NIndiv, residual(0), residual(1))
                    If (DesiredVar > VarPossibleLower) And (DesiredVar < VarPossibleUpper) Then
                        Dim nLarger As Integer = FindNT_NIndivJ_Var(covariate, NJ, NIndiv, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                        Dim nT, nC As Integer
                        Dim EachTGroupCost As Double = TGroupcost + (NIndiv * TIndivCost)
                        Dim EachCGroupCost As Double = CGroupCost + (NIndiv * CIndivCost)
                        If (EachTGroupCost >= EachCGroupCost) Then
                            nT = nLarger
                            nC = NJ - nLarger
                        Else
                            nC = nLarger
                            nT = NJ - nLarger
                        End If
                        result(0) = nT
                        result(1) = nC
                        result(2) = NIndiv
                        result(3) = CDbl(nT) / (nT + nC)
                    Else
                        MsgBox("The number of individuals in each group and number of groups specified is impossible to make variance to achieve specified power of width of CI of ES")
                        result(0) = -1
                        result(1) = -1
                        result(2) = NIndiv
                        result(3) = -1
                    End If
                Else
                    'NIndiv Only
                    Dim NIndiv As Integer = NIndiv2
                    Dim pT() As Double = StartPT
                    Dim NJ(0 To UBound(pT)) As Integer
                    Dim TCost(0 To UBound(pT)) As Double
                    For i = 0 To UBound(pT)
                        NJ(i) = FindJ_PNIndiv_Var(covariate, pT(i), NIndiv, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                        TCost(i) = FindTotalCostRevertible(NJ(i), pT(i), NIndiv, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                    Next i
                    Dim AllCeilingJ As Boolean = AllElementEqual(NJ, 20000)
                    Dim MinimumCost As Integer = FindMinimum(TCost)
                    If AllCeilingJ = True Then
                        result(0) = -1
                        result(1) = -1
                        result(2) = NIndiv
                        result(3) = -1
                        MsgBox("The algorithm was terminated because number of groups was increasing over 20000.")
                    ElseIf (MinimumCost = 0) Then
                        Dim Range As Double = pT(1) - pT(0)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(0) + (i * 0.005)
                        Next i
                        Dim NJ3(0 To numElem) As Integer
                        Dim TCost3(0 To numElem) As Double
                        For i = 0 To numElem
                            NJ3(i) = FindJ_PNIndiv_Var(covariate, pTScope(i), NIndiv, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TCost3(i) = FindTotalCostRevertible(NJ3(i), pTScope(i), NIndiv, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(TCost3)
                        If (MinPosition3 > 0) Then
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv
                            If (NJ3(MinPosition3) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(MinPosition3) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            result(0) = Math.Round(0.5 * NJ3(0))
                            result(1) = Math.Round(0.5 * NJ3(0))
                            result(2) = NIndiv
                            result(3) = 0.5
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                            If (NJ3(0) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(0) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    ElseIf (MinimumCost < UBound(TCost)) Then
                        Dim Range As Double = pT(MinimumCost + 1) - pT(MinimumCost - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(MinimumCost - 1) + (i * 0.005)
                        Next i
                        Dim NJ3(0 To numElem) As Integer
                        Dim TCost3(0 To numElem) As Double
                        For i = 0 To numElem
                            NJ3(i) = FindJ_PNIndiv_Var(covariate, pTScope(i), NIndiv, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TCost3(i) = FindTotalCostRevertible(NJ3(i), pTScope(i), NIndiv, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(TCost3)
                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                            result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                            result(3) = pTScope(MinPosition3)
                        Else
                            result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                            result(1) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                            result(3) = 1 - pTScope(MinPosition3)
                        End If
                        result(2) = NIndiv
                        If (NJ3(MinPosition3) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (NJ3(MinPosition3) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim Range As Double = pT(UBound(TCost)) - pT(UBound(TCost) - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(UBound(TCost) - 1) + (i * 0.005)
                        Next i
                        Dim NJ3(0 To numElem) As Integer
                        Dim TCost3(0 To numElem) As Double
                        For i = 0 To numElem
                            NJ3(i) = FindJ_PNIndiv_Var(covariate, pTScope(i), NIndiv, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TCost3(i) = FindTotalCostRevertible(NJ3(i), pTScope(i), NIndiv, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(TCost3)
                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                        If (MinPosition3 < UBound(TCost3)) Then
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv
                            If (NJ3(MinPosition3) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(MinPosition3) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            If TCostLess = True Then
                                result(0) = Math.Round(0.95 * NJ3(UBound(NJ3)))
                                result(1) = Math.Round(0.05 * NJ3(UBound(NJ3)))
                                result(3) = 0.95
                            Else
                                result(0) = Math.Round(0.05 * NJ3(UBound(NJ3)))
                                result(1) = Math.Round(0.95 * NJ3(UBound(NJ3)))
                                result(3) = 0.05
                            End If
                            result(2) = NIndiv
                            If (NJ3(UBound(NJ3)) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(UBound(NJ3)) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If (NJ2 <> -1) Then
                    'NG Only
                    Dim NJ As Integer = NJ2
                    Dim pT() As Double = StartPT
                    Dim NIndiv(0 To UBound(pT)) As Integer
                    Dim TCost(0 To UBound(pT)) As Double
                    For i = 0 To UBound(pT)
                        NIndiv(i) = FindNIndiv_PJ_Var(covariate, pT(i), NJ, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                        TCost(i) = FindTotalCostRevertible(NJ, pT(i), NIndiv(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                    Next i
                    Dim AllCeiling As Boolean = AllElementEqual(NIndiv, 100000)
                    Dim MinimumCost As Integer = FindMinimum(TCost)
                    If AllCeiling = True Then
                        result(0) = -1
                        result(1) = -1
                        result(2) = -1
                        result(3) = -1
                        MsgBox("The algorithm was terminated because number of individuals in each group was increasing over 100000.")
                    ElseIf (MinimumCost = 0) Then
                        Dim Range As Double = pT(1) - pT(0)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(0) + (i * 0.005)
                        Next i
                        Dim NIndiv3(0 To numElem) As Integer
                        Dim TCost3(0 To numElem) As Double
                        For i = 0 To numElem
                            NIndiv3(i) = FindNIndiv_PJ_Var(covariate, pTScope(i), NJ, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TCost3(i) = FindTotalCostRevertible(NJ, pTScope(i), NIndiv3(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(TCost3)
                        If (MinPosition3 > 0) Then
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(MinPosition3) * TIndivCost)) <= (CGroupCost + (NIndiv3(MinPosition3) * CIndivCost))
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv3(MinPosition3)
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            result(0) = Math.Round(0.5 * NJ)
                            result(1) = Math.Round(0.5 * NJ)
                            result(2) = NIndiv3(0)
                            result(3) = 0.5
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(0) * TIndivCost)) <= (CGroupCost + (NIndiv3(0) * CIndivCost))
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    ElseIf (MinimumCost < UBound(TCost)) Then
                        Dim Range As Double = pT(MinimumCost + 1) - pT(MinimumCost - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(MinimumCost - 1) + (i * 0.005)
                        Next i
                        Dim NIndiv3(0 To numElem) As Integer
                        Dim TCost3(0 To numElem) As Double
                        For i = 0 To numElem
                            NIndiv3(i) = FindNIndiv_PJ_Var(covariate, pTScope(i), NJ, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TCost3(i) = FindTotalCostRevertible(NJ, pTScope(i), NIndiv3(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(TCost3)
                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(MinPosition3) * TIndivCost)) <= (CGroupCost + (NIndiv3(MinPosition3) * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(pTScope(MinPosition3) * NJ)
                            result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                            result(3) = pTScope(MinPosition3)
                        Else
                            result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                            result(1) = Math.Round(pTScope(MinPosition3) * NJ)
                            result(3) = 1 - pTScope(MinPosition3)
                        End If
                        result(2) = NIndiv3(MinPosition3)
                        If (NJ > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (NJ < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim Range As Double = pT(UBound(TCost)) - pT(UBound(TCost) - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(UBound(TCost) - 1) + (i * 0.005)
                        Next i
                        Dim NIndiv3(0 To numElem) As Integer
                        Dim TCost3(0 To numElem) As Double
                        For i = 0 To numElem
                            NIndiv3(i) = FindNIndiv_PJ_Var(covariate, pTScope(i), NJ, DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TCost3(i) = FindTotalCostRevertible(NJ, pTScope(i), NIndiv3(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(TCost3)

                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(MinPosition3) * TIndivCost)) <= (CGroupCost + (NIndiv3(MinPosition3) * CIndivCost))
                        If (MinPosition3 < UBound(TCost3)) Then
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv3(MinPosition3)
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            If TCostLess = True Then
                                result(0) = Math.Round(0.95 * NJ)
                                result(1) = Math.Round(0.05 * NJ)
                                result(3) = 0.95
                            Else
                                result(0) = Math.Round(0.05 * NJ)
                                result(1) = Math.Round(0.95 * NJ)
                                result(3) = 0.05
                            End If
                            result(2) = NIndiv3(MinPosition3)
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    End If
                Else
                    'None specified
                    Dim NJ() As Integer = StartJ
                    Dim pT() As Double = StartPT
                    Dim numJ As Integer = UBound(NJ)
                    Dim numPT As Integer = UBound(pT)
                    Dim NIndiv(0 To numJ, 0 To numPT) As Integer
                    Dim TCost(0 To numJ, 0 To numPT) As Double
                    For i = 0 To numJ
                        For j = 0 To numPT
                            NIndiv(i, j) = FindNIndiv_PJ_Var(covariate, pT(j), NJ(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            TCost(i, j) = FindTotalCostRevertible(NJ(i), pT(j), NIndiv(i, j), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        Next j
                    Next i
                    Dim FindNewJ As Boolean = AllElementEqual(NIndiv, 100000)
                    Dim MinimumCost() As Integer = FindMinimum(TCost)
                    If (FindNewJ = False) AndAlso (MinimumCost(0) = 0) Then
                        'Use PJ11
                        Dim RangeJ As Integer = NJ(1) - NJ(0)
                        Dim JScope(0 To RangeJ) As Integer
                        For i = 0 To RangeJ
                            JScope(i) = NJ(0) + i
                        Next i
                        Dim RangeP As Double
                        Dim BaseP As Double
                        If (MinimumCost(1) = 0) Then
                            RangeP = pT(1) - pT(0)
                            BaseP = pT(0)
                        ElseIf (MinimumCost(1) < UBound(pT)) Then
                            RangeP = pT(MinimumCost(1) + 1) - pT(MinimumCost(1) - 1)
                            BaseP = pT(MinimumCost(1) - 1)
                        Else
                            RangeP = pT(UBound(pT)) - pT(UBound(pT) - 1)
                            BaseP = pT(UBound(pT) - 1)
                        End If
                        Dim numElemP As Integer = Math.Round(RangeP / 0.005)
                        Dim PTScope(0 To numElemP) As Double
                        For j = 0 To numElemP
                            PTScope(j) = BaseP + (j * 0.005)
                        Next j

                        Dim nIndiv3(0 To RangeJ, 0 To numElemP) As Integer
                        Dim TNIndiv3(0 To RangeJ, 0 To numElemP) As Double
                        For i = 0 To RangeJ
                            For j = 0 To numElemP
                                nIndiv3(i, j) = FindNIndiv_PJ_Var(covariate, PTScope(j), JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                TNIndiv3(i, j) = FindTotalCostRevertible(JScope(i), PTScope(j), nIndiv3(i, j), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            Next j
                        Next i
                        Dim MinPosition3() As Integer = FindMinimum(TNIndiv3)
                        If (MinPosition3(0) > 0) Then
                            Dim TCostLess As Boolean = (TGroupcost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * TIndivCost)) <= (CGroupCost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * CIndivCost))
                            If TCostLess = True Then
                                result(0) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                                result(1) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                                result(3) = PTScope(MinPosition3(1))
                            Else
                                result(0) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                                result(1) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                                result(3) = 1 - PTScope(MinPosition3(1))
                            End If
                            result(2) = nIndiv3(MinPosition3(0), MinPosition3(1))
                            If (JScope(MinPosition3(0)) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (JScope(MinPosition3(0)) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            Dim nTnC(,) As Integer = pJ11
                            Dim nIndiv4(0 To 35) As Integer
                            Dim pNTNC(0 To 35) As Double
                            Dim TNIndiv4(0 To 35) As Double
                            For i = 0 To 35
                                pNTNC(i) = CDbl(nTnC(i, 0)) / (nTnC(i, 0) + nTnC(i, 1))
                                nIndiv4(i) = FindNIndiv_PJ_Var(covariate, pNTNC(i), nTnC(i, 0) + nTnC(i, 1), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                TNIndiv4(i) = FindTotalCostRevertible(nTnC(i, 0), nTnC(i, 1), nIndiv4(i), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            Next
                            Dim MinPosition4 As Integer = FindMinimum(TNIndiv4)
                            result(0) = nTnC(MinPosition4, 0)
                            result(1) = nTnC(MinPosition4, 1)
                            result(2) = nIndiv4(MinPosition4)
                            result(3) = pNTNC(MinPosition4)
                        End If
                    ElseIf (FindNewJ = False) AndAlso (MinimumCost(0) < TCost.GetLength(0) - 1) Then
                        Dim RangeJ As Integer = NJ(MinimumCost(0) + 1) - NJ(MinimumCost(0) - 1)
                        Dim JScope(0 To RangeJ) As Integer
                        For i = 0 To RangeJ
                            JScope(i) = NJ(MinimumCost(0) - 1) + i
                        Next i
                        Dim RangeP As Double
                        Dim BaseP As Double
                        If (MinimumCost(1) = 0) Then
                            RangeP = pT(1) - pT(0)
                            BaseP = pT(0)
                        ElseIf (MinimumCost(1) < UBound(pT)) Then
                            RangeP = pT(MinimumCost(1) + 1) - pT(MinimumCost(1) - 1)
                            BaseP = pT(MinimumCost(1) - 1)
                        Else
                            RangeP = pT(UBound(pT)) - pT(UBound(pT) - 1)
                            BaseP = pT(UBound(pT) - 1)
                        End If
                        Dim numElemP As Integer = Math.Round(RangeP / 0.005)
                        Dim PTScope(0 To numElemP) As Double
                        For j = 0 To numElemP
                            PTScope(j) = BaseP + (j * 0.005)
                        Next j

                        Dim nIndiv3(0 To RangeJ, 0 To numElemP) As Integer
                        Dim TNIndiv3(0 To RangeJ, 0 To numElemP) As Double
                        For i = 0 To RangeJ
                            For j = 0 To numElemP
                                nIndiv3(i, j) = FindNIndiv_PJ_Var(covariate, PTScope(j), JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                TNIndiv3(i, j) = FindTotalCostRevertible(JScope(i), PTScope(j), nIndiv3(i, j), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            Next j
                        Next i
                        Dim MinPosition3() As Integer = FindMinimum(TNIndiv3)
                        Dim TCostLess As Boolean = (TGroupcost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * TIndivCost)) <= (CGroupCost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(3) = PTScope(MinPosition3(1))
                        Else
                            result(0) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(3) = 1 - PTScope(MinPosition3(1))
                        End If
                        result(2) = nIndiv3(MinPosition3(0), MinPosition3(1))
                        If (JScope(MinPosition3(0)) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3(0)) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim RunNJ() As Integer = {170, 190, 210}
                        numJ = 2
                        Dim RunNIndiv(0 To numJ, 0 To numPT) As Integer
                        Dim RunTCost(0 To numJ, 0 To numPT) As Double
                        For i = 0 To numJ
                            For j = 0 To numPT
                                RunNIndiv(i, j) = FindNIndiv_PJ_Var(covariate, pT(j), RunNJ(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                RunTCost(i, j) = FindTotalCostRevertible(RunNJ(i), pT(j), RunNIndiv(i, j), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            Next j
                        Next i
                        MinimumCost = FindMinimum(RunTCost)
                        Do While (AllElementEqual(RunNIndiv, 100000)) OrElse (MinimumCost(0) <> 1)
                            For i = 0 To 1
                                RunNJ(i) = RunNJ(i + 1)
                                For j = 0 To numPT
                                    RunNIndiv(i, j) = RunNIndiv(i + 1, j)
                                    RunTCost(i, j) = RunTCost(i + 1, j)
                                Next j
                            Next i
                            RunNJ(2) = RunNJ(2) + 20
                            For j = 0 To numPT
                                RunNIndiv(2, j) = FindNIndiv_PJ_Var(covariate, pT(j), RunNJ(2), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                RunTCost(2, j) = FindTotalCostRevertible(RunNJ(2), pT(j), RunNIndiv(2, j), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            Next
                            MinimumCost = FindMinimum(RunTCost)
                        Loop
                        Dim RangeJ As Integer = RunNJ(2) - RunNJ(0)
                        Dim JScope(0 To RangeJ) As Integer
                        For i = 0 To RangeJ
                            JScope(i) = RunNJ(0) + i
                        Next i
                        Dim RangeP As Double
                        Dim BaseP As Double
                        If (MinimumCost(1) = 0) Then
                            RangeP = pT(1) - pT(0)
                            BaseP = pT(0)
                        ElseIf (MinimumCost(1) < UBound(pT)) Then
                            RangeP = pT(MinimumCost(1) + 1) - pT(MinimumCost(1) - 1)
                            BaseP = pT(MinimumCost(1) - 1)
                        Else
                            RangeP = pT(UBound(pT)) - pT(UBound(pT) - 1)
                            BaseP = pT(UBound(pT) - 1)
                        End If
                        Dim numElemP As Integer = Math.Round(RangeP / 0.005)
                        Dim PTScope(0 To numElemP) As Double
                        For j = 0 To numElemP
                            PTScope(j) = BaseP + (j * 0.005)
                        Next j
                        Dim nIndiv3(0 To RangeJ, 0 To numElemP) As Integer
                        Dim TNIndiv3(0 To RangeJ, 0 To numElemP) As Double
                        For i = 0 To RangeJ
                            For j = 0 To numElemP
                                nIndiv3(i, j) = FindNIndiv_PJ_Var(covariate, PTScope(j), JScope(i), DesiredVar, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                TNIndiv3(i, j) = FindTotalCostRevertible(JScope(i), PTScope(j), nIndiv3(i, j), TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            Next j
                        Next i
                        Dim MinPosition3() As Integer = FindMinimum(TNIndiv3)
                        Dim TCostLess As Boolean = (TGroupcost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * TIndivCost)) <= (CGroupCost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(3) = PTScope(MinPosition3(1))
                        Else
                            result(0) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(3) = 1 - PTScope(MinPosition3(1))
                        End If
                        result(2) = nIndiv3(MinPosition3(0), MinPosition3(1))
                        If (JScope(MinPosition3(0)) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3(0)) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return result
    End Function

    Public Function MInverseV(ByVal tau As Double, ByVal sigma As Double, ByVal nIndiv As Integer)
        Dim result(0 To 1) As Decimal
        Dim Comn As Decimal = (sigma + (nIndiv * tau))
        Dim Comn1 As Decimal = (sigma + ((nIndiv - 1) * tau))
        result(0) = Comn1 / (sigma * Comn)
        result(1) = -tau / (sigma * Comn)
        Return result
    End Function

    Public Function PrintM(ByVal M(,) As Double)
        Dim result As String = ""
        For i = 0 To M.GetLength(0) - 1
            For j = 0 To M.GetLength(1) - 1
                result &= M(i, j).ToString("f4") & vbTab
            Next j
            result &= vbCrLf
        Next i
        Return result
    End Function

    Public Function AllElementEqual(ByVal vector() As Integer, ByVal Value As Integer)
        For i = 0 To UBound(vector)
            If (vector(i) <> Value) Then
                Return False
            End If
        Next i
        Return True
    End Function

    Public Function AllElementEqual(ByVal array(,) As Integer, ByVal Value As Integer)
        For i = 0 To array.GetLength(0) - 1
            For j = 0 To array.GetLength(1) - 1
                If (array(i, j) <> Value) Then
                    Return False
                End If
            Next j
        Next i
        Return True
    End Function

    Public Function FindOption3(ByVal covariate As Boolean, ByVal ES As Double, ByVal ICCY As Double, ByVal TotalCost As Double, ByVal TGroupcost As Double, ByVal TIndivCost As Double, ByVal CGroupCost As Double, ByVal CIndivCost As Double, Optional ByVal ICCZ As Double = 0, Optional ByVal R2ClusterZ As Double = 0, Optional ByVal R2IndivZ As Double = 0, Optional ByVal pT2 As Double = -1, Optional ByVal NJ2 As Integer = -1, Optional ByVal NIndiv2 As Integer = -1)
        Dim StartJ() As Integer = {10, 30, 50, 70, 90, 110, 130, 150, 170, 190}
        Dim StartPT() As Double = {0.5, 0.6, 0.7, 0.8, 0.9, 0.95}
        Dim result(0 To 3) As Double 'NT NC And NINDIV

        ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! Return - 1 in J and nIndiv
        If (pT2 <> -1) Then
            If (NIndiv2 <> -1) Then
                'P & NINDIV
                Dim pT As Double = pT2
                Dim nIndiv As Integer = NIndiv2
                Dim NJ As Integer = FindJ_PNIndiv_TotalCostExact(pT, nIndiv, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                result(2) = nIndiv
                result(3) = pT
                If NJ = -1 Then
                    result(0) = -1
                    result(1) = -1
                    MsgBox("The totalcost is not enough to collect at least two groups from treatment and control conditions.")
                Else
                    result(0) = Math.Round(NJ * pT)
                    result(1) = Math.Round(NJ * (1 - pT))
                    Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                    If (NJ > (result(0) + result(1))) Then
                        If (TCostLess = True) Then
                            result(0) += 1
                        Else
                            result(1) += 1
                        End If
                    ElseIf (NJ < (result(0) + result(1))) Then
                        If (TCostLess = True) Then
                            result(1) -= 1
                        Else
                            result(0) -= 1
                        End If
                    End If
                End If
            Else
                If (NJ2 <> -1) Then
                    'P & NG
                    Dim pT As Double = pT2
                    Dim NJ As Integer = NJ2
                    Dim NIndiv As Integer = FindNIndiv_PJ_TotalCostExact(pT, NJ, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                    result(0) = Math.Round(NJ * pT)
                    result(1) = Math.Round(NJ * (1 - pT))
                    If NIndiv = -1 Then MsgBox("The totalcost is not enough to collect number of individuals more than or equal to 2 in each group.")
                    result(2) = NIndiv
                    result(3) = pT
                    Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                    If (NJ > (result(0) + result(1))) Then
                        If (TCostLess = True) Then
                            result(0) += 1
                        Else
                            result(1) += 1
                        End If
                    ElseIf (NJ < (result(0) + result(1))) Then
                        If (TCostLess = True) Then
                            result(1) -= 1
                        Else
                            result(0) -= 1
                        End If
                    End If
                Else
                    'P only
                    Dim pT As Double = pT2
                    Dim NJ() As Integer = StartJ
                    Dim NIndiv(0 To UBound(NJ)) As Integer
                    Dim Var(0 To UBound(NJ)) As Double
                    For i = 0 To UBound(NJ)
                        NIndiv(i) = FindNIndiv_PJ_TotalCostExact(pT, NJ(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        If NIndiv(i) = -1 Then
                            Var(i) = 100000
                        Else
                            Dim residual() As Double = FindResidual(covariate, pT, NJ(i), NIndiv(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            Var(i) = FindVar(NJ(i), pT, NIndiv(i), residual(0), residual(1))
                        End If
                    Next i
                    Dim FindCheaperDesign As Boolean = AllElementEqual(NIndiv, -1)
                    Dim MinimumVar = FindMinimum(Var)
                    If (FindCheaperDesign = True) Or (MinimumVar = 0) Then
                        Dim Range As Integer = NJ(1) - NJ(0)
                        Dim JScope(0 To Range) As Integer
                        For i = 0 To Range
                            JScope(i) = NJ(0) + i
                        Next i
                        Dim nIndiv4(0 To Range) As Integer
                        Dim Var4(0 To Range) As Double
                        For i = 0 To Range
                            nIndiv4(i) = FindNIndiv_PJ_TotalCostExact(pT, JScope(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If nIndiv4(i) = -1 Then
                                Var4(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pT, JScope(i), nIndiv4(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var4(i) = FindVar(JScope(i), pT, nIndiv4(i), residual(0), residual(1))
                            End If
                        Next
                        Dim FindCheaperDesign4 As Boolean = AllElementEqual(nIndiv4, -1)
                        Dim MinPosition4 As Integer = FindMinimum(Var4)
                        If (FindCheaperDesign4 = False) And (MinPosition4 > 0) Then
                            result(0) = Math.Round(JScope(MinPosition4) * pT)
                            result(1) = Math.Round(JScope(MinPosition4) * (1 - pT))
                            result(2) = nIndiv4(MinPosition4)
                            result(3) = pT
                            Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                            If (JScope(MinPosition4) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (JScope(MinPosition4) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            Dim nTnC(,) As Integer = pJ11
                            Dim nIndiv3(0 To 35) As Integer
                            Dim pNTNC(0 To 35) As Double
                            Dim Var3(0 To 35) As Double
                            For i = 0 To 35
                                pNTNC(i) = CDbl(nTnC(i, 0)) / (nTnC(i, 0) + nTnC(i, 1))
                                If Math.Abs(pNTNC(i) - pT) <= 0.1 Then
                                    nIndiv3(i) = FindNIndiv_PJ_TotalCostExact(pNTNC(i), nTnC(i, 0) + nTnC(i, 1), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                                Else
                                    nIndiv3(i) = -1
                                End If
                                If nIndiv3(i) = -1 Then
                                    Var3(i) = 100000
                                Else
                                    Dim residual() As Double = FindResidual(covariate, pNTNC(i), nTnC(i, 0) + nTnC(i, 1), nIndiv3(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                    Var3(i) = FindVar(nTnC(i, 0), nTnC(i, 1), nIndiv3(i), residual(0), residual(1))
                                End If
                            Next
                            Dim FindCheaperDesign3 As Boolean = AllElementEqual(nIndiv3, -1)
                            Dim MinPosition3 As Integer = FindMinimum(Var3)
                            If (FindCheaperDesign3) Then
                                result(0) = -1
                                result(1) = -1
                                result(2) = -1
                                result(3) = -1
                                MsgBox("The total cost is too low to make a cluster randomized design.")
                            Else
                                result(0) = nTnC(MinPosition3, 0)
                                result(1) = nTnC(MinPosition3, 1)
                                result(2) = nIndiv3(MinPosition3)
                                result(3) = pNTNC(MinPosition3)
                            End If
                        End If
                    ElseIf (MinimumVar < UBound(NJ)) Then
                        Dim Range As Integer = NJ(MinimumVar + 1) - NJ(MinimumVar - 1)
                        Dim JScope(0 To Range) As Integer
                        For i = 0 To Range
                            JScope(i) = NJ(MinimumVar - 1) + i
                        Next i
                        Dim nIndiv3(0 To Range) As Integer
                        Dim Var3(0 To Range) As Double
                        For i = 0 To Range
                            nIndiv3(i) = FindNIndiv_PJ_TotalCostExact(pT, JScope(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If nIndiv3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pT, JScope(i), nIndiv3(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(JScope(i), pT, nIndiv3(i), residual(0), residual(1))
                            End If
                        Next
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        result(0) = Math.Round(JScope(MinPosition3) * pT)
                        result(1) = Math.Round(JScope(MinPosition3) * (1 - pT))
                        result(2) = nIndiv3(MinPosition3)
                        result(3) = pT
                        Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                        If (JScope(MinPosition3) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim RunJ() As Integer = {170, 190, 210}
                        Dim RunNIndiv(0 To 2) As Integer
                        Dim RunVarNIndiv(0 To 2) As Double
                        For i = 0 To 2
                            RunNIndiv(i) = FindNIndiv_PJ_TotalCostExact(pT, RunJ(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If RunNIndiv(i) = -1 Then
                                RunVarNIndiv(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pT, RunJ(i), RunNIndiv(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                RunVarNIndiv(i) = FindVar(RunJ(i), pT, RunNIndiv(i), residual(0), residual(1))
                            End If
                        Next i
                        Do While (FindMinimum(RunVarNIndiv) <> 1)
                            RunJ(0) = RunJ(1)
                            RunJ(1) = RunJ(2)
                            RunJ(2) = RunJ(2) + 20
                            RunNIndiv(0) = RunNIndiv(1)
                            RunNIndiv(1) = RunNIndiv(2)
                            RunNIndiv(2) = FindNIndiv_PJ_TotalCostExact(pT, RunJ(2), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            RunVarNIndiv(0) = RunVarNIndiv(1)
                            RunVarNIndiv(1) = RunVarNIndiv(2)
                            If RunNIndiv(2) = -1 Then
                                RunVarNIndiv(2) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pT, RunJ(2), RunNIndiv(2), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                RunVarNIndiv(2) = FindVar(RunJ(2), pT, RunNIndiv(2), residual(0), residual(1))
                            End If
                        Loop
                        Dim Range As Integer = RunJ(2) - RunJ(0)
                        Dim JScope(0 To Range) As Integer
                        For i = 0 To Range
                            JScope(i) = RunJ(0) + i
                        Next i
                        Dim nIndiv3(0 To Range) As Integer
                        Dim Var3(0 To Range) As Double
                        For i = 0 To Range
                            nIndiv3(i) = FindNIndiv_PJ_TotalCostExact(pT, JScope(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If nIndiv3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pT, JScope(i), nIndiv3(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(JScope(i), pT, nIndiv3(i), residual(0), residual(1))
                            End If
                        Next
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        result(0) = Math.Round(JScope(MinPosition3) * pT)
                        result(1) = Math.Round(JScope(MinPosition3) * (1 - pT))
                        result(2) = nIndiv3(MinPosition3)
                        result(3) = pT
                        Dim TCostLess As Boolean = (TGroupcost + (result(2) * TIndivCost)) <= (CGroupCost + (result(2) * CIndivCost))
                        If (JScope(MinPosition3) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If (NIndiv2 <> -1) Then
                If (NJ2 <> -1) Then
                    'NIndiv & NG
                    Dim NIndiv As Integer = NIndiv2
                    Dim NJ As Integer = NJ2
                    Dim pTScope() As Double = {0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55, 0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95}
                    Dim TCostPTScope(0 To UBound(pTScope)) As Double
                    For i = 0 To UBound(pTScope)
                        TCostPTScope(i) = FindTotalCostExact(NJ, pTScope(i), NIndiv, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                    Next
                    Dim MinCostPosition As Integer = FindMinimum(TCostPTScope)
                    Dim MaxCostPosition As Integer = FindMaximum(TCostPTScope)
                    If (TotalCost > MinCostPosition) And (TotalCost < MaxCostPosition) Then
                        Dim nT As Integer = FindNT_NIndivJ_TotalCostMax(NJ, NIndiv, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        result(0) = nT
                        result(1) = NJ - nT
                        result(2) = NIndiv
                        result(3) = CDbl(nT) / NJ
                    Else
                        MsgBox("The number of individuals in each group and number of groups specified is impossible to make total cost specified.")
                        result(0) = -1
                        result(1) = -1
                        result(2) = NIndiv
                        result(3) = -1
                    End If
                Else
                    'NIndiv Only
                    Dim NIndiv As Integer = NIndiv2
                    Dim pT() As Double = StartPT
                    Dim NJ(0 To UBound(pT)) As Integer
                    Dim Var(0 To UBound(pT)) As Double
                    For i = 0 To UBound(pT)
                        NJ(i) = FindJ_PNIndiv_TotalCostMax(pT(i), NIndiv, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        If NJ(i) = -1 Then
                            Var(i) = 100000
                        Else
                            Dim residual() As Double = FindResidual(covariate, pT(i), NJ(i), NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            Var(i) = FindVar(NJ(i), pT(i), NIndiv, residual(0), residual(1))
                        End If
                    Next i
                    If AllElementEqual(NJ, -1) = True Then
                        result(0) = -1
                        result(1) = -1
                        result(2) = NIndiv
                        result(3) = -1
                        MsgBox("The total cost specified is not enough to collect at least two clusters for both treatment and control conditions.")
                        Return result
                    End If
                    Dim MinimumVar As Integer = FindMinimum(Var)
                    If (MinimumVar = 0) Then
                        Dim Range As Double = pT(1) - pT(0)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(0) + (i * 0.005)
                        Next i
                        Dim NJ3(0 To numElem) As Integer
                        Dim Var3(0 To numElem) As Double
                        For i = 0 To numElem
                            NJ3(i) = FindJ_PNIndiv_TotalCostMax(pTScope(i), NIndiv, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If NJ3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pTScope(i), NJ3(i), NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(NJ3(i), pTScope(i), NIndiv, residual(0), residual(1))
                            End If
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        If (MinPosition3 > 0) Then
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv
                            If (NJ3(MinPosition3) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(MinPosition3) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            result(0) = Math.Round(0.5 * NJ3(0))
                            result(1) = Math.Round(0.5 * NJ3(0))
                            result(2) = NIndiv
                            result(3) = 0.5
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                            If (NJ3(0) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(0) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    ElseIf (MinimumVar < UBound(Var)) Then
                        Dim Range As Double = pT(MinimumVar + 1) - pT(MinimumVar - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(MinimumVar - 1) + (i * 0.005)
                        Next i
                        Dim NJ3(0 To numElem) As Integer
                        Dim Var3(0 To numElem) As Double
                        For i = 0 To numElem
                            NJ3(i) = FindJ_PNIndiv_TotalCostMax(pTScope(i), NIndiv, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If NJ3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pTScope(i), NJ3(i), NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(NJ3(i), pTScope(i), NIndiv, residual(0), residual(1))
                            End If
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                            result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                            result(3) = pTScope(MinPosition3)
                        Else
                            result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                            result(1) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                            result(3) = 1 - pTScope(MinPosition3)
                        End If
                        result(2) = NIndiv
                        If (NJ3(MinPosition3) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (NJ3(MinPosition3) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim Range As Double = pT(UBound(Var)) - pT(UBound(Var) - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(UBound(Var) - 1) + (i * 0.005)
                        Next i
                        Dim NJ3(0 To numElem) As Integer
                        Dim Var3(0 To numElem) As Double
                        For i = 0 To numElem
                            NJ3(i) = FindJ_PNIndiv_TotalCostMax(pTScope(i), NIndiv, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If NJ3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pTScope(i), NJ3(i), NIndiv, ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(NJ3(i), pTScope(i), NIndiv, residual(0), residual(1))
                            End If
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv * TIndivCost)) <= (CGroupCost + (NIndiv * CIndivCost))
                        If (MinPosition3 < UBound(Var3)) Then
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ3(MinPosition3))
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ3(MinPosition3))
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv
                            If (NJ3(MinPosition3) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(MinPosition3) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            If TCostLess = True Then
                                result(0) = Math.Round(0.95 * NJ3(UBound(NJ3)))
                                result(1) = Math.Round(0.05 * NJ3(UBound(NJ3)))
                                result(3) = 0.95
                            Else
                                result(0) = Math.Round(0.05 * NJ3(UBound(NJ3)))
                                result(1) = Math.Round(0.95 * NJ3(UBound(NJ3)))
                                result(3) = 0.05
                            End If
                            result(2) = NIndiv
                            If (NJ3(UBound(NJ3)) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ3(UBound(NJ3)) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If (NJ2 <> -1) Then
                    'NG Only
                    Dim NJ As Integer = NJ2
                    Dim pT() As Double = StartPT
                    Dim NIndiv(0 To UBound(pT)) As Integer
                    Dim Var(0 To UBound(pT)) As Double
                    For i = 0 To UBound(pT)
                        NIndiv(i) = FindNIndiv_PJ_TotalCostMax(pT(i), NJ, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                        If NIndiv(i) = -1 Then
                            Var(i) = 100000
                        Else
                            Dim residual() As Double = FindResidual(covariate, pT(i), NJ, NIndiv(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                            Var(i) = FindVar(NJ, pT(i), NIndiv(i), residual(0), residual(1))
                        End If
                    Next i
                    If AllElementEqual(NIndiv, -1) Then
                        result(0) = -1
                        result(1) = -1
                        result(2) = -1
                        result(3) = -1
                        MsgBox("The total cost is not enough to make " & NJ.ToString("n0") & " clusters.")
                        Return result
                    End If
                    Dim MinimumVar As Integer = FindMinimum(Var)
                    If (MinimumVar = 0) Then
                        Dim Range As Double = pT(1) - pT(0)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(0) + (i * 0.005)
                        Next i
                        Dim NIndiv3(0 To numElem) As Integer
                        Dim Var3(0 To numElem) As Double
                        For i = 0 To numElem
                            NIndiv3(i) = FindNIndiv_PJ_TotalCostMax(pTScope(i), NJ, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If NIndiv3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pTScope(i), NJ, NIndiv3(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(NJ, pTScope(i), NIndiv3(i), residual(0), residual(1))
                            End If
                        Next
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        If (MinPosition3 > 0) Then
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(MinPosition3) * TIndivCost)) <= (CGroupCost + (NIndiv3(MinPosition3) * CIndivCost))
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv3(MinPosition3)
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            result(0) = Math.Round(0.5 * NJ)
                            result(1) = Math.Round(0.5 * NJ)
                            result(2) = NIndiv3(0)
                            result(3) = 0.5
                            Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(0) * TIndivCost)) <= (CGroupCost + (NIndiv3(0) * CIndivCost))
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    ElseIf (MinimumVar < UBound(Var)) Then
                        Dim Range As Double = pT(MinimumVar + 1) - pT(MinimumVar - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(MinimumVar - 1) + (i * 0.005)
                        Next i
                        Dim NIndiv3(0 To numElem) As Integer
                        Dim Var3(0 To numElem) As Double
                        For i = 0 To numElem
                            NIndiv3(i) = FindNIndiv_PJ_TotalCostMax(pTScope(i), NJ, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If NIndiv3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pTScope(i), NJ, NIndiv3(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(NJ, pTScope(i), NIndiv3(i), residual(0), residual(1))
                            End If
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(MinPosition3) * TIndivCost)) <= (CGroupCost + (NIndiv3(MinPosition3) * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(pTScope(MinPosition3) * NJ)
                            result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                            result(3) = pTScope(MinPosition3)
                        Else
                            result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                            result(1) = Math.Round(pTScope(MinPosition3) * NJ)
                            result(3) = 1 - pTScope(MinPosition3)
                        End If
                        result(2) = NIndiv3(MinPosition3)
                        If (NJ > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (NJ < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim Range As Double = pT(UBound(Var)) - pT(UBound(Var) - 1)
                        Dim numElem As Integer = Math.Round(Range / 0.005)
                        Dim pTScope(numElem) As Double
                        For i = 0 To numElem
                            pTScope(i) = pT(UBound(Var) - 1) + (i * 0.005)
                        Next i
                        Dim NIndiv3(0 To numElem) As Integer
                        Dim Var3(0 To numElem) As Double
                        For i = 0 To numElem
                            NIndiv3(i) = FindNIndiv_PJ_TotalCostMax(pTScope(i), NJ, TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If NIndiv3(i) = -1 Then
                                Var3(i) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pTScope(i), NJ, NIndiv3(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var3(i) = FindVar(NJ, pTScope(i), NIndiv3(i), residual(0), residual(1))
                            End If
                        Next i
                        Dim MinPosition3 As Integer = FindMinimum(Var3)
                        Dim TCostLess As Boolean = (TGroupcost + (NIndiv3(MinPosition3) * TIndivCost)) <= (CGroupCost + (NIndiv3(MinPosition3) * CIndivCost))
                        If (MinPosition3 < UBound(Var3)) Then
                            If TCostLess = True Then
                                result(0) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(1) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(3) = pTScope(MinPosition3)
                            Else
                                result(0) = Math.Round((1 - pTScope(MinPosition3)) * NJ)
                                result(1) = Math.Round(pTScope(MinPosition3) * NJ)
                                result(3) = 1 - pTScope(MinPosition3)
                            End If
                            result(2) = NIndiv3(MinPosition3)
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            If TCostLess = True Then
                                result(0) = Math.Round(0.95 * NJ)
                                result(1) = Math.Round(0.05 * NJ)
                                result(3) = 0.95
                            Else
                                result(0) = Math.Round(0.05 * NJ)
                                result(1) = Math.Round(0.95 * NJ)
                                result(3) = 0.05
                            End If
                            result(2) = NIndiv3(MinPosition3)
                            If (NJ > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (NJ < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        End If
                    End If
                Else
                    'None specified
                    Dim NJ() As Integer = StartJ
                    Dim pT() As Double = StartPT
                    Dim numJ As Integer = UBound(NJ)
                    Dim numPT As Integer = UBound(pT)
                    Dim NIndiv(0 To numJ, 0 To numPT) As Integer
                    Dim Var(0 To numJ, 0 To numPT) As Double
                    For i = 0 To numJ
                        For j = 0 To numPT
                            NIndiv(i, j) = FindNIndiv_PJ_TotalCostMax(pT(j), NJ(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                            If NIndiv(i, j) = -1 Then
                                Var(i, j) = 100000
                            Else
                                Dim residual() As Double = FindResidual(covariate, pT(j), NJ(i), NIndiv(i, j), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                Var(i, j) = FindVar(NJ(i), pT(j), NIndiv(i, j), residual(0), residual(1))
                            End If
                        Next j
                    Next i
                    Dim FindCheaperDesign As Boolean = AllElementEqual(NIndiv, -1)
                    Dim MinimumVar() As Integer = FindMinimum(Var)
                    If FindCheaperDesign = True Or (MinimumVar(0) = 0) Then
                        Dim RangeJ As Integer = NJ(1) - NJ(0)
                        Dim JScope(0 To RangeJ) As Integer
                        For i = 0 To RangeJ
                            JScope(i) = NJ(0) + i
                        Next i
                        Dim RangeP As Double
                        Dim BaseP As Double
                        If (MinimumVar(1) = 0) Then
                            RangeP = pT(1) - pT(0)
                            BaseP = pT(0)
                        ElseIf (MinimumVar(1) < UBound(pT)) Then
                            RangeP = pT(MinimumVar(1) + 1) - pT(MinimumVar(1) - 1)
                            BaseP = pT(MinimumVar(1) - 1)
                        Else
                            RangeP = pT(UBound(pT)) - pT(UBound(pT) - 1)
                            BaseP = pT(UBound(pT) - 1)
                        End If
                        Dim numElemP As Integer = Math.Round(RangeP / 0.005)
                        Dim PTScope(0 To numElemP) As Double
                        For j = 0 To numElemP
                            PTScope(j) = BaseP + (j * 0.005)
                        Next j
                        Dim nIndiv3(0 To RangeJ, 0 To numElemP) As Integer
                        Dim Var3(0 To RangeJ, 0 To numElemP) As Double
                        For i = 0 To RangeJ
                            For j = 0 To numElemP
                                nIndiv3(i, j) = FindNIndiv_PJ_TotalCostMax(PTScope(j), JScope(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                                If nIndiv3(i, j) = -1 Then
                                    Var3(i, j) = 100000
                                Else
                                    Dim residual() As Double = FindResidual(covariate, PTScope(j), JScope(i), nIndiv3(i, j), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                    Var3(i, j) = FindVar(JScope(i), PTScope(j), nIndiv3(i, j), residual(0), residual(1))
                                End If
                            Next j
                        Next i
                        Dim FindCheaperDesign3 As Boolean = AllElementEqual(nIndiv3, -1)
                        Dim MinPosition3() As Integer = FindMinimum(Var3)
                        If (FindCheaperDesign3 = False) And (MinPosition3(0) > 0) Then
                            Dim TCostLess As Boolean = (TGroupcost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * TIndivCost)) <= (CGroupCost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * CIndivCost))
                            If TCostLess = True Then
                                result(0) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                                result(1) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                                result(3) = PTScope(MinPosition3(1))
                            Else
                                result(0) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                                result(1) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                                result(3) = 1 - PTScope(MinPosition3(1))
                            End If
                            result(2) = nIndiv3(MinPosition3(0), MinPosition3(1))
                            If (JScope(MinPosition3(0)) > (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(0) += 1
                                Else
                                    result(1) += 1
                                End If
                            ElseIf (JScope(MinPosition3(0)) < (result(0) + result(1))) Then
                                If (TCostLess = True) Then
                                    result(1) -= 1
                                Else
                                    result(0) -= 1
                                End If
                            End If
                        Else
                            Dim nTnC(,) As Integer = pJ11
                            Dim nIndiv4(0 To 35) As Integer
                            Dim pNTNC(0 To 35) As Double
                            Dim VarNIndiv4(0 To 35) As Double
                            For i = 0 To 35
                                nIndiv4(i) = FindNIndiv_PJ_TotalCostMax(pNTNC(i), nTnC(i, 0) + nTnC(i, 1), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                                If nIndiv4(i) = -1 Then
                                    VarNIndiv4(i) = 100000
                                Else
                                    Dim residual() As Double = FindResidual(covariate, pNTNC(i), nTnC(i, 0) + nTnC(i, 1), nIndiv4(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                    VarNIndiv4(i) = FindVar(nTnC(i, 0), nTnC(i, 1), nIndiv4(i), residual(0), residual(1))
                                End If
                            Next
                            If AllElementEqual(nIndiv4, -1) Then
                                result(0) = -1
                                result(1) = -1
                                result(2) = -1
                                result(3) = -1
                                MsgBox("The total cost specified is not enough to make clustered randomized design with group and individual costs specified.")
                            Else
                                Dim MinPosition4 As Integer = FindMinimum(VarNIndiv4)
                                result(0) = nTnC(MinPosition4, 0)
                                result(1) = nTnC(MinPosition4, 1)
                                result(2) = nIndiv4(MinPosition4)
                                result(3) = pNTNC(MinPosition4)
                            End If
                        End If
                    ElseIf (MinimumVar(0) < Var.GetLength(0) - 1) Then
                        Dim RangeJ As Integer = NJ(MinimumVar(0) + 1) - NJ(MinimumVar(0) - 1)
                        Dim JScope(0 To RangeJ) As Integer
                        For i = 0 To RangeJ
                            JScope(i) = NJ(MinimumVar(0)) + i
                        Next i
                        Dim RangeP As Double
                        Dim BaseP As Double
                        If (MinimumVar(1) = 0) Then
                            RangeP = pT(1) - pT(0)
                            BaseP = pT(0)
                        ElseIf (MinimumVar(1) < UBound(pT)) Then
                            RangeP = pT(MinimumVar(1) + 1) - pT(MinimumVar(1) - 1)
                            BaseP = pT(MinimumVar(1) - 1)
                        Else
                            RangeP = pT(UBound(pT)) - pT(UBound(pT) - 1)
                            BaseP = pT(UBound(pT) - 1)
                        End If
                        Dim numElemP As Integer = Math.Round(RangeP / 0.005)
                        Dim PTScope(0 To numElemP) As Double
                        For j = 0 To numElemP
                            PTScope(j) = BaseP + (j * 0.005)
                        Next j

                        Dim nIndiv3(0 To RangeJ, 0 To numElemP) As Integer
                        Dim Var3(0 To RangeJ, 0 To numElemP) As Double
                        For i = 0 To RangeJ
                            For j = 0 To numElemP
                                nIndiv3(i, j) = FindNIndiv_PJ_TotalCostMax(PTScope(j), JScope(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                                If nIndiv3(i, j) = -1 Then
                                    Var3(i, j) = 100000
                                Else
                                    Dim residual() As Double = FindResidual(covariate, PTScope(j), JScope(i), nIndiv3(i, j), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                    Var3(i, j) = FindVar(JScope(i), PTScope(j), nIndiv3(i, j), residual(0), residual(1))
                                End If
                            Next j
                        Next i
                        Dim MinPosition3() As Integer = FindMinimum(Var3)
                        Dim TCostLess As Boolean = (TGroupcost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * TIndivCost)) <= (CGroupCost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(3) = PTScope(MinPosition3(1))
                        Else
                            result(0) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(3) = 1 - PTScope(MinPosition3(1))
                        End If
                        result(2) = nIndiv3(MinPosition3(0), MinPosition3(1))
                        If (JScope(MinPosition3(0)) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3(0)) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    Else
                        Dim RunNJ() As Integer = {170, 190, 210}
                        numJ = 2
                        Dim RunNIndiv(0 To numJ, 0 To numPT) As Integer
                        Dim RunVarNIndiv(0 To numJ, 0 To numPT) As Double
                        For i = 0 To numJ
                            For j = 0 To numPT
                                RunNIndiv(i, j) = FindNIndiv_PJ_TotalCostMax(pT(j), RunNJ(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                                If RunNIndiv(i, j) = -1 Then
                                    RunVarNIndiv(i, j) = 100000
                                Else
                                    Dim residual() As Double = FindResidual(covariate, pT(j), RunNJ(i), RunNIndiv(i, j), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                    RunVarNIndiv(i, j) = FindVar(RunNJ(i), pT(j), RunNIndiv(i, j), residual(0), residual(1))
                                End If
                            Next j
                        Next i
                        MinimumVar = FindMinimum(RunVarNIndiv)
                        Do While (MinimumVar(0) <> 1)
                            For i = 0 To 1
                                RunNJ(i) = RunNJ(i + 1)
                                For j = 0 To numPT
                                    RunNIndiv(i, j) = RunNIndiv(i + 1, j)
                                    RunVarNIndiv(i, j) = RunVarNIndiv(i + 1, j)
                                Next j
                            Next i
                            RunNJ(2) = RunNJ(2) + 20
                            For j = 0 To numPT
                                RunNIndiv(2, j) = FindNIndiv_PJ_TotalCostMax(pT(j), RunNJ(2), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                                If RunNIndiv(2, j) = -1 Then
                                    RunVarNIndiv(2, j) = 100000
                                Else
                                    Dim residual() As Double = FindResidual(covariate, pT(j), RunNJ(2), RunNIndiv(2, j), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                    RunVarNIndiv(2, j) = FindVar(RunNJ(2), pT(j), RunNIndiv(2, j), residual(0), residual(1))
                                End If
                            Next
                            MinimumVar = FindMinimum(RunVarNIndiv)
                        Loop
                        Dim RangeJ As Integer = RunNJ(2) - RunNJ(0)
                        Dim JScope(0 To RangeJ) As Integer
                        For i = 0 To RangeJ
                            JScope(i) = RunNJ(0) + i
                        Next i
                        Dim RangeP As Double
                        Dim BaseP As Double
                        If (MinimumVar(1) = 0) Then
                            RangeP = pT(1) - pT(0)
                            BaseP = pT(0)
                        ElseIf (MinimumVar(1) < UBound(pT)) Then
                            RangeP = pT(MinimumVar(1) + 1) - pT(MinimumVar(1) - 1)
                            BaseP = pT(MinimumVar(1) - 1)
                        Else
                            RangeP = pT(UBound(pT)) - pT(UBound(pT) - 1)
                            BaseP = pT(UBound(pT) - 1)
                        End If
                        Dim numElemP As Integer = Math.Round(RangeP / 0.005)
                        Dim PTScope(0 To numElemP) As Double
                        For j = 0 To numElemP
                            PTScope(j) = BaseP + (j * 0.005)
                        Next j
                        Dim nIndiv3(0 To RangeJ, 0 To numElemP) As Integer
                        Dim Var3(0 To RangeJ, 0 To numElemP) As Double
                        For i = 0 To RangeJ
                            For j = 0 To numElemP
                                nIndiv3(i, j) = FindNIndiv_PJ_TotalCostMax(PTScope(j), JScope(i), TotalCost, TGroupcost, TIndivCost, CGroupCost, CIndivCost)
                                If nIndiv3(i, j) = -1 Then
                                    Var3(i, j) = 100000
                                Else
                                    Dim residual() As Double = FindResidual(covariate, PTScope(j), JScope(i), nIndiv3(i, j), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
                                    Var3(i, j) = FindVar(JScope(i), PTScope(j), nIndiv3(i, j), residual(0), residual(1))
                                End If
                            Next j
                        Next i
                        Dim MinPosition3() As Integer = FindMinimum(Var3)
                        Dim TCostLess As Boolean = (TGroupcost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * TIndivCost)) <= (CGroupCost + (nIndiv3(MinPosition3(0), MinPosition3(1)) * CIndivCost))
                        If TCostLess = True Then
                            result(0) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(3) = PTScope(MinPosition3(1))
                        Else
                            result(0) = Math.Round(JScope(MinPosition3(0)) * (1 - PTScope(MinPosition3(1))))
                            result(1) = Math.Round(JScope(MinPosition3(0)) * PTScope(MinPosition3(1)))
                            result(3) = 1 - PTScope(MinPosition3(1))
                        End If
                        result(2) = nIndiv3(MinPosition3(0), MinPosition3(1))
                        If (JScope(MinPosition3(0)) > (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(0) += 1
                            Else
                                result(1) += 1
                            End If
                        ElseIf (JScope(MinPosition3(0)) < (result(0) + result(1))) Then
                            If (TCostLess = True) Then
                                result(1) -= 1
                            Else
                                result(0) -= 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return result
    End Function

    Public Function FindNIndiv_pJ_Var_MonteCarlo(ByVal PowerOrWidth As Integer, ByVal covariate As Boolean, ByVal NJ As Integer, ByVal pT As Double, ByVal NIndiv As Integer, ByVal ICCY As Double, ByVal ES As Double, ByVal ICCZ As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double, ByVal Rep As Integer, ByVal Power As Double, ByVal CILevel As Double, ByVal Width As Double, ByVal ClusterOnlyZ As Boolean, ByVal ReverseZ As Boolean, ByVal IsDegreeOfCertainty As Boolean, Optional ByVal DegreeOfCertainty As Double = 0)
        ' Check possiblity in writing
        Dim RunNIndiv(0 To 2) As Integer
        If NIndiv > MinIndiv Then
            RunNIndiv(0) = NIndiv - 1
            RunNIndiv(1) = NIndiv
            RunNIndiv(2) = NIndiv + 1
        Else
            RunNIndiv(0) = MinIndiv
            RunNIndiv(1) = MinIndiv + 1
            RunNIndiv(2) = MinIndiv + 2
        End If
        Dim Result(0 To 2) As Double
        For i = 0 To 2
            Dim possible As Boolean = CheckPossibleNJpTNIndiv(covariate, NJ, pT, RunNIndiv(i), ES, ICCY, ICCZ, R2ClusterZ, R2IndivZ)
            If possible Then
                If PowerOrWidth = 1 Then
                    Result(i) = RunSimulationAndReadPower(covariate, NJ, pT, RunNIndiv(i), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, ClusterOnlyZ, ReverseZ)
                Else
                    If IsDegreeOfCertainty = True Then
                        Result(i) = RunSimulationAndReadDegreeOfCertainty(covariate, NJ, pT, RunNIndiv(i), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, CILevel, ClusterOnlyZ, ReverseZ, Width)
                    Else
                        Result(i) = RunSimulationAndReadCIES(covariate, NJ, pT, RunNIndiv(i), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, CILevel, ClusterOnlyZ, ReverseZ)
                    End If
                End If
            Else
                Result(i) = -1
            End If
            If (Main.ProgressBar1.Value < 750) Then
                Main.ProgressBar1.Value += 50
                Main.StatusNum.Text = CInt(Main.StatusNum.Text) + 5
            End If
        Next i

        If PowerOrWidth = 1 Then
            Dim Index As Integer = FindIndexPass(Result, Power, True)
            If (Result(0) = -1) Or Index = -1 Then
                'Increment
                Do
                    RunNIndiv(0) = RunNIndiv(1)
                    RunNIndiv(1) = RunNIndiv(2)
                    RunNIndiv(2) = RunNIndiv(2) + 1
                    Result(0) = Result(1)
                    Result(1) = Result(2)
                    Result(2) = RunSimulationAndReadPower(covariate, NJ, pT, RunNIndiv(2), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, ClusterOnlyZ, ReverseZ)
                    Index = FindIndexPass(Result, Power, True)
                Loop While Index = -1
            ElseIf (Result(2) = -1) Or Index = 0 Then
                'Decrease
                ' IF IT IS < 2 DONE
                If RunNIndiv(0) > MinIndiv Then
                    Do
                        RunNIndiv(2) = RunNIndiv(1)
                        RunNIndiv(1) = RunNIndiv(0)
                        RunNIndiv(0) = RunNIndiv(0) - 1
                        If RunNIndiv(1) = MinIndiv Then Return MinIndiv
                        Result(2) = Result(1)
                        Result(1) = Result(0)
                        Result(0) = RunSimulationAndReadPower(covariate, NJ, pT, RunNIndiv(0), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, ClusterOnlyZ, ReverseZ)
                        Index = FindIndexPass(Result, Power, True)
                    Loop While Index = 0
                End If
            End If
            Return RunNIndiv(Index)
        Else
            If IsDegreeOfCertainty = True Then
                Dim Index As Integer = FindIndexPass(Result, DegreeOfCertainty, True)
                If (Result(0) = -1) Or Index = -1 Then
                    'Increment
                    Do
                        RunNIndiv(0) = RunNIndiv(1)
                        RunNIndiv(1) = RunNIndiv(2)
                        RunNIndiv(2) = RunNIndiv(2) + 1
                        Result(0) = Result(1)
                        Result(1) = Result(2)
                        Result(2) = RunSimulationAndReadDegreeOfCertainty(covariate, NJ, pT, RunNIndiv(2), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, CILevel, ClusterOnlyZ, ReverseZ, Width)
                        Index = FindIndexPass(Result, DegreeOfCertainty, True)
                    Loop While Index = -1
                ElseIf (Result(2) = -1) Or Index = 0 Then
                    'Decrease
                    ' IF IT IS < 2 DONE
                    If RunNIndiv(0) > MinIndiv Then
                        Do
                            RunNIndiv(2) = RunNIndiv(1)
                            RunNIndiv(1) = RunNIndiv(0)
                            RunNIndiv(0) = RunNIndiv(0) - 1
                            If RunNIndiv(1) = MinIndiv Then Return MinIndiv
                            Result(2) = Result(1)
                            Result(1) = Result(0)
                            Result(0) = RunSimulationAndReadDegreeOfCertainty(covariate, NJ, pT, RunNIndiv(0), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, CILevel, ClusterOnlyZ, ReverseZ, Width)
                            Index = FindIndexPass(Result, DegreeOfCertainty, True)
                        Loop While Index = 0
                    End If
                End If
                Return RunNIndiv(Index)
            Else
                Dim Index As Integer = FindIndexPass(Result, Width, False)
                If (Result(0) = -1) Or Index = -1 Then
                    'Increment
                    Do
                        RunNIndiv(0) = RunNIndiv(1)
                        RunNIndiv(1) = RunNIndiv(2)
                        RunNIndiv(2) = RunNIndiv(2) + 1
                        Result(0) = Result(1)
                        Result(1) = Result(2)
                        Result(2) = RunSimulationAndReadCIES(covariate, NJ, pT, RunNIndiv(2), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, CILevel, ClusterOnlyZ, ReverseZ)
                        Index = FindIndexPass(Result, Width, False)
                    Loop While Index = -1
                ElseIf (Result(2) = -1) Or Index = 0 Then
                    'Decrease
                    If RunNIndiv(0) > MinIndiv Then
                        Do
                            RunNIndiv(2) = RunNIndiv(1)
                            RunNIndiv(1) = RunNIndiv(0)
                            RunNIndiv(0) = RunNIndiv(0) - 1
                            If RunNIndiv(1) = MinIndiv Then Return MinIndiv
                            Result(2) = Result(1)
                            Result(1) = Result(0)
                            Result(0) = RunSimulationAndReadCIES(covariate, NJ, pT, RunNIndiv(0), ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, CILevel, ClusterOnlyZ, ReverseZ)
                            Index = FindIndexPass(Result, Width, False)
                        Loop While Index = 0
                    End If
                End If
                Return RunNIndiv(Index)
            End If
        End If
        Return NIndiv
    End Function

    Function ReadPower(ByVal filename As String)
        Dim fullfilename As String = My.Application.Info.DirectoryPath & "\" & filename & ".out"
        Dim numLine As Short = 0
        Dim sr As StreamReader = ReadMplusOutputMod(fullfilename)
        Dim allText As String = sr.ReadToEnd
        sr.Close()

        Dim charInFile As Integer = allText.Length
        Dim i As Integer
        Dim letter As String
        For i = 0 To charInFile - 1
            letter = allText.Substring(i, 1)
            If letter = Chr(13) Then
                numLine += 1
            End If
        Next i

        Dim fileArray(numLine) As String
        Dim currentLine As Integer = 1
        Dim Line As String = ""
        For i = 0 To charInFile - 1
            letter = allText.Substring(i, 1)
            If letter = Chr(13) Then
                fileArray(currentLine - 1) = Line
                currentLine += 1
                Line = ""
            Else
                Line &= letter
            End If
        Next i

        Dim find As String = "Between Level"
        Dim numResult As Integer = 0
        Dim findResult(0 To 1) As Integer
        For i = 0 To numLine - 1
            If fileArray(i).Contains(find) Then
                findResult(numResult) = i
                numResult += 1
            End If
        Next i

        Dim powerLine As String = fileArray(findResult(0) + 3)
        Dim powerSentence() As String = powerLine.Split(" "c)

        currentLine = 0
        Dim cleanPowerSentence(0 To 8) As String
        For i = 0 To UBound(powerSentence)
            If powerSentence(i) <> "" AndAlso powerSentence(i) <> " " Then
                cleanPowerSentence(currentLine) = powerSentence(i)
                currentLine += 1
            End If
        Next i
        Dim power As Double = CDbl(cleanPowerSentence(8))
        Return power
    End Function

    Public Function CheckPossibleNJpTNIndiv(ByVal covariate As Boolean, ByVal NJ As Integer, ByVal pT As Double, ByVal NIndiv As Integer, ByVal ES As Double, ByVal ICCY As Double, Optional ByVal ICCZ As Double = 0, Optional ByVal R2ClusterZ As Double = 0, Optional ByVal R2IndivZ As Double = 0)
        Dim result As Boolean = True
        'Dim varYonX As Double = (pT * (1 - pT)) * (ES * ES)
        'Dim ResidualYBetween, ResidualYWithin As Double
        'If (covariate = True) Then
        '    Dim dfB As Integer = NJ - 1
        '    Dim dfW As Integer = NJ * (NIndiv - 1)
        '    Dim tauZ As Double = (dfB + dfW) / (dfB + (dfW * (1 - ICCZ) / ICCZ))
        '    Dim sigmaZ As Double = tauZ * (1 - ICCZ) / ICCZ
        '    Dim SSBZ As Double = dfB * tauZ
        '    Dim SSWZ As Double = dfW * sigmaZ
        '    Dim etaZ As Double = SSBZ / (SSBZ + SSWZ)
        '    Dim BetaB As Double = TEZ + ((1 - etaZ) * CEZ)
        '    Dim BetaW As Double = BetaB - CEZ
        '    If (BetaB > 1.0) Or (BetaB < -1.0) Then
        '        result = False
        '    End If
        '    If (BetaW > 1.0) Or (BetaW < -1.0) Then
        '        result = False
        '    End If
        '    ResidualYBetween = (ICCY / (1 - ICCY)) - varYonX - (BetaB * BetaB * tauZ)
        '    ResidualYWithin = 1 - (BetaW * BetaW * sigmaZ)
        '    If (ResidualYBetween < 0) Then
        '        result = False
        '    End If
        '    If (ResidualYWithin < 0) Then
        '        result = False
        '    End If
        'Else
        '    ResidualYBetween = (ICCY / (1 - ICCY)) - varYonX
        '    If (ResidualYBetween < 0) Then
        '        result = False
        '    End If
        'End If
        Return result 'It does not need to check now because TEZ and CEZ are changed to proportion error variance explained
    End Function

    Public Function CompareAllElementWithScalar(ByVal array() As Double, ByVal scalar As Double)
        Dim compareArray(0 To UBound(array)) As Integer
        For i = 0 To UBound(array)
            If array(i) = scalar Then
                Return 2 'Some elements are exactly equal
            ElseIf array(i) > scalar Then
                compareArray(i) = 1
            Else
                compareArray(i) = -1
            End If
        Next i
        If AllElementEqual(compareArray, 1) Then Return 1
        If AllElementEqual(compareArray, -1) Then Return -1
        Return 0
    End Function

    Public Function RunSimulationAndReadPower(ByVal covariate As Boolean, ByVal NJ As Integer, ByVal pT As Double, ByVal NIndiv As Integer, ByVal ICCY As Double, ByVal ES As Double, ByVal ICCZ As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double, ByVal Rep As Integer, ByVal ClusterOnlyZ As Boolean, ByVal ReverseZ As Boolean)
        If (covariate = True) Then
            WriteMCCovariateCode(NJ, pT, NIndiv, ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, False, ClusterOnlyZ, ReverseZ)
        Else
            WriteMCNullCode(NJ, pT, NIndiv, ICCY, ES, Rep, MakeFiles:=False)
        End If
        Dim sim As String = "simulation"
        RunAnalysis(sim)
        While DetectProcess("Mplus")
            System.Threading.Thread.Sleep(3000)
        End While
        Dim Power As Double = ReadPower(sim)
        WriteMonteCarloArchieve(NJ, pT, NIndiv, Power)
        My.Computer.FileSystem.DeleteFile(Path & "\simulation.inp")
        My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\simulation.out")
        Return Power
    End Function

    Public Function RunSimulationAndReadCIES(ByVal covariate As Boolean, ByVal NJ As Integer, ByVal pT As Double, ByVal NIndiv As Integer, ByVal ICCY As Double, ByVal ES As Double, ByVal ICCZ As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double, ByVal Rep As Integer, ByVal CILevel As Double, ByVal ClusterOnlyZ As Boolean, ByVal ReverseZ As Boolean)
        If (covariate = True) Then
            WriteMCCovariateCode(NJ, pT, NIndiv, ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, True, ClusterOnlyZ, ReverseZ)
        Else
            WriteMCNullCode(NJ, pT, NIndiv, ICCY, ES, Rep, MakeFiles:=True)
        End If
        Dim sim As String = "simulation"
        RunAnalysis(sim)
        While DetectProcess("Mplus")
            System.Threading.Thread.Sleep(500)
        End While
        Dim Width As Double
        Dim result(,) As Double = RunEachSimulatedData(covariate, NRep, False, ClusterOnlyZ, NJ, NIndiv)
        Dim sumWidth95 As Double = 0, sumWidth99 As Double = 0, sumEst As Double = 0, sumPower As Double = 0
        Dim width95, width99 As Double
        Dim CorrectedResult(,) As Double
        For i = 0 To Rep - 1
            While result(i, 5) = -1
                If (covariate = True) Then
                    WriteMCCovariateCode(NJ, pT, NIndiv, ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, 1, True, ClusterOnlyZ, ReverseZ)
                Else
                    WriteMCNullCode(NJ, pT, NIndiv, ICCY, ES, 1, MakeFiles:=True)
                End If
                RunAnalysis(sim)
                CorrectedResult = RunEachSimulatedData(covariate, 1, False, ClusterOnlyZ, NJ, NIndiv)
                For k = 0 To 5
                    result(i, k) = CorrectedResult(0, k)
                Next k
            End While
            width95 = result(i, 3) - result(i, 1)
            width99 = result(i, 4) - result(i, 0)
            sumEst += result(i, 2)
            sumPower += result(i, 5)
            sumWidth95 += width95
            sumWidth99 += width99
        Next i
        If CILevel = 0.95 Then
            Width = sumWidth95 / Rep
        Else
            Width = sumWidth99 / Rep
        End If
        WriteMonteCarloArchieve(NJ, pT, NIndiv, Width)
        My.Computer.FileSystem.DeleteFile(Path & "\simulation.inp")
        My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\simulation.out")
        My.Computer.FileSystem.DeleteFile(Path & "\rawlist.dat")
        ClearAllData(NRep)
        Return Width
    End Function

    Public Function RunSimulationAndReadDegreeOfCertainty(ByVal covariate As Boolean, ByVal NJ As Integer, ByVal pT As Double, ByVal NIndiv As Integer, ByVal ICCY As Double, ByVal ES As Double, ByVal ICCZ As Double, ByVal R2ClusterZ As Double, ByVal R2IndivZ As Double, ByVal Rep As Integer, ByVal CILevel As Double, ByVal ClusterOnlyZ As Boolean, ByVal ReverseZ As Boolean, ByVal SpecifiedWidth As Double)
        If (covariate = True) Then
            WriteMCCovariateCode(NJ, pT, NIndiv, ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, Rep, True, ClusterOnlyZ, ReverseZ)
        Else
            WriteMCNullCode(NJ, pT, NIndiv, ICCY, ES, Rep, MakeFiles:=True)
        End If
        Dim sim As String = "simulation"
        RunAnalysis(sim)
        While DetectProcess("Mplus")
            System.Threading.Thread.Sleep(500)
        End While
        Dim DegreeOfCertainty As Double
        Dim result(,) As Double = RunEachSimulatedData(covariate, NRep, False, ClusterOnlyZ, NJ, NIndiv)
        Dim sumWidth95 As Double = 0, sumWidth99 As Double = 0, sumEst As Double = 0, sumPower As Double = 0
        Dim width95, width99 As Double
        Dim NumLowerWidth95 As Integer = 0, NumLowerWidth99 As Integer = 0
        Dim CorrectedResult(,) As Double
        For i = 0 To Rep - 1
            While result(i, 5) = -1
                If (covariate = True) Then
                    WriteMCCovariateCode(NJ, pT, NIndiv, ICCY, ES, ICCZ, R2ClusterZ, R2IndivZ, 1, True, ClusterOnlyZ, ReverseZ)
                Else
                    WriteMCNullCode(NJ, pT, NIndiv, ICCY, ES, 1, MakeFiles:=True)
                End If
                RunAnalysis(sim)
                CorrectedResult = RunEachSimulatedData(covariate, 1, False, ClusterOnlyZ, NJ, NIndiv)
                For k = 0 To 5
                    result(i, k) = CorrectedResult(0, k)
                Next k
            End While
            width95 = result(i, 3) - result(i, 1)
            width99 = result(i, 4) - result(i, 0)
            If (width95 <= SpecifiedWidth) Then NumLowerWidth95 += 1
            If (width99 <= SpecifiedWidth) Then NumLowerWidth99 += 1
        Next i
        If CILevel = 0.95 Then
            DegreeOfCertainty = NumLowerWidth95 / Rep

        Else
            DegreeOfCertainty = NumLowerWidth99 / Rep
        End If
        WriteMonteCarloArchieve(NJ, pT, NIndiv, DegreeOfCertainty)
        My.Computer.FileSystem.DeleteFile(Path & "\simulation.inp")
        My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\simulation.out")
        My.Computer.FileSystem.DeleteFile(Path & "\rawlist.dat")
        ClearAllData(NRep)
        Return DegreeOfCertainty
    End Function

    Public Function RunEachSimulatedData(ByVal covariate As Boolean, ByVal rep As Integer, ByVal runCovaraiteDataByNull As Boolean, ByVal ClusterOnlyZ As Boolean, ByVal NGroups As Integer, ByVal NIndiv As Integer)
        Dim i, j As Integer
        Dim TempText = Main.status.Text
        Dim counter As Integer = 0
        While DetectProcess("Mplus")
            System.Threading.Thread.Sleep(500)
            counter += 5
            Main.status.Text = "Waiting for Mplus Creating Data (" & counter.ToString("n0") & " seconds)"
        End While
        Main.status.Text = TempText
        Dim ncase As Integer = NGroups * NIndiv
        For i = 1 To rep
            Dim filename As String = "raw" & i
            If covariate = True Then
                'If ClusterOnlyZ = True Then
                '    addY(filename, ncase, 4, 2)
                'Else
                '    addY(filename, ncase, 4, 1)
                'End If
                WriteCovariateAnalysisCode(filename, ClusterOnlyZ)
            Else
                'If runCovaraiteDataByNull = False Then addY(filename, ncase, 3, 1)
                WriteNullAnalysisCode(filename, runCovaraiteDataByNull, ClusterOnlyZ)
            End If
            RunAnalysis(filename)
            'System.Threading.Thread.Sleep(500)
        Next i
        TempText = Main.status.Text
        counter = 0
        While DetectProcess("Mplus")
            System.Threading.Thread.Sleep(500)
            counter += 5
            Main.status.Text = "Waiting for Mplus Analyzing Simulated Data (" & counter.ToString("n0") & " seconds)"
        End While
        Main.status.Text = TempText
        Dim result(0 To rep - 1, 5) As Double
        Dim temp(0 To 5) As Double
        For i = 1 To rep
            Dim filename As String = "raw" & i
            'RunAnalysis(filename)
            Try
                temp = ReadES(filename)
            Catch ex As Exception
                My.Computer.FileSystem.DeleteFile(Path & "\" & filename & ".inp")
                WriteCovariateAnalysisCodeSecondTrial(filename, ClusterOnlyZ)
                RunAnalysis(filename)
                While DetectProcess("Mplus")
                    System.Threading.Thread.Sleep(500)
                End While
                Try
                    temp = ReadES(filename)
                Catch ex2 As Exception
                    For k = 0 To 5
                        temp(k) = -1
                    Next k
                End Try
            End Try
            For j = 0 To 5
                result(i - 1, j) = temp(j)
            Next j
            My.Computer.FileSystem.DeleteFile(Path & "\" & filename & ".inp")
            My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\" & filename & ".out")
        Next i
        Return result
    End Function

    Public Sub ClearAllData(ByVal NRep As Integer)
        For i = 1 To NRep
            Dim filename As String = "raw" & i
            My.Computer.FileSystem.DeleteFile(Path & "\" & filename & ".dat")
        Next i
    End Sub

    Public Function FindCloset(ByVal array() As Double, ByVal scalar As Double, ByVal OverSide As Boolean)
        Dim Distance(0 To UBound(array)) As Double
        For i = 0 To UBound(array)
            Distance(i) = array(i) - scalar
        Next i
        Dim Position As Integer = -1
        Dim MinDistance As Double = 10000000
        For i = 0 To UBound(array)
            If OverSide = True Then
                If (Distance(i) > 0) AndAlso (Math.Abs(Distance(i)) < MinDistance) Then
                    Position = i
                    MinDistance = Math.Abs(Distance(i))
                End If
            Else
                If (Distance(i) < 0) AndAlso (Math.Abs(Distance(i)) < MinDistance) Then
                    Position = i
                    MinDistance = Math.Abs(Distance(i))
                End If
            End If
        Next
        Return Position
    End Function

    Public Sub SaveFile(ByVal filename As String)
        Dim Tab0 As String = "Tab0" & vbCrLf & _
                            "NEXP=" & vbCrLf & Main.NTGroups.Text & vbCrLf & _
                            "NCTRL=" & vbCrLf & Main.NCGroups.Text & vbCrLf & _
                            "NIndiv=" & vbCrLf & Main.NIndiv.Text & vbCrLf & _
                            "ICCY=" & vbCrLf & Main.ICCY.Text & vbCrLf & _
                            "ES=" & vbCrLf & Main.ES.Text & vbCrLf & _
                            "ICCZ=" & vbCrLf & Main.ICCZ.Text & vbCrLf & _
                            "TEZ=" & vbCrLf & Main.R2ClusterZ.Text & vbCrLf & _
                            "CEZ=" & vbCrLf & Main.R2IndivZ.Text & vbCrLf & _
                            "WidthCI95=" & vbCrLf & Main.WidthCI95.Text & vbCrLf & _
                            "WidthCI99=" & vbCrLf & Main.WidthCI99.Text & vbCrLf & _
                            "Power=" & vbCrLf & Main.Power.Text & vbCrLf & _
                            "WidthCI95C=" & vbCrLf & Main.WidthCI95C.Text & vbCrLf & _
                            "WidthCI99C=" & vbCrLf & Main.WidthCI99C.Text & vbCrLf & _
                            "PowerC=" & vbCrLf & Main.PowerC.Text & vbCrLf
        Dim Tab1 As String = "Tab1" & vbCrLf & _
                            "ICCY2=" & vbCrLf & Main.ICCY2.Text & vbCrLf & _
                            "ES2=" & vbCrLf & Main.ES2.Text & vbCrLf & _
                            "ICCZ2=" & vbCrLf & Main.ICCZ2.Text & vbCrLf & _
                            "TEZ2=" & vbCrLf & Main.R2ClusterZ2.Text & vbCrLf & _
                            "CEZ2=" & vbCrLf & Main.R2IndivZ2.Text & vbCrLf & _
                            "NIndiv2=" & vbCrLf & Main.NIndiv2.Text & vbCrLf & _
                            "NEXP2=" & vbCrLf & Main.NTGroups2.Text & vbCrLf & _
                            "NCTRL2=" & vbCrLf & Main.NCGroups2.Text & vbCrLf & _
                            "TOTALCOST2=" & vbCrLf & Main.TotalCost.Text & vbCrLf & _
                            "Power2=" & vbCrLf & Main.ObtainedPower.Text & vbCrLf & _
                            "WidthCI952=" & vbCrLf & Main.ObtainedWidth95.Text & vbCrLf & _
                            "WidthCI992=" & vbCrLf & Main.ObtainedWidth99.Text & vbCrLf & _
                            "PowerC2=" & vbCrLf & Main.ObtainedPowerC.Text & vbCrLf & _
                            "WidthCI95C2=" & vbCrLf & Main.ObtainedWidth95C.Text & vbCrLf & _
                            "WidthCI99C2=" & vbCrLf & Main.ObtainedWidth99C.Text & vbCrLf
        Dim InstanceVariables As String = "InstanceVariables" & vbCrLf & _
                                            "NTValue=" & vbCrLf & NTValue & vbCrLf & _
                                            "NCValue=" & vbCrLf & NCValue & vbCrLf & _
                                            "NGValue=" & vbCrLf & NGValue & vbCrLf & _
                                            "PTValue=" & vbCrLf & PTValue & vbCrLf & _
                                            "NIndivValue=" & vbCrLf & NIndivValue & vbCrLf & _
                                            "OptionValue=" & vbCrLf & OptionValue & vbCrLf & _
                                            "CILevelValue=" & vbCrLf & CILevelValue & vbCrLf & _
                                            "PowerOrWidthValue=" & vbCrLf & PowerOrWidthValue & vbCrLf & _
                                            "PowerValue=" & vbCrLf & PowerValue & vbCrLf & _
                                            "Widthvalue=" & vbCrLf & WidthValue & vbCrLf & _
                                            "TGroupCostValue=" & vbCrLf & TGroupCostValue & vbCrLf & _
                                            "CGroupCostValue=" & vbCrLf & CGroupCostValue & vbCrLf & _
                                            "TIndivCostValue=" & vbCrLf & TIndivCostValue & vbCrLf & _
                                            "CIndivCostValue=" & vbCrLf & CIndivCostValue & vbCrLf & _
                                            "TotalCostValue=" & vbCrLf & TotalCostValue & vbCrLf
        Dim CriteriaCheck As String = "CriteriaCheck" & vbCrLf & _
                                        "PCheckBox" & vbCrLf & WriteBool(Criteria.PCheckBox.Checked) & vbCrLf & _
                                        "NGCheckBox" & vbCrLf & WriteBool(Criteria.NGCheckBox.Checked) & vbCrLf & _
                                        "NIndivCheckBox" & vbCrLf & WriteBool(Criteria.NIndivCheckBox.Checked) & vbCrLf
        Dim MainCheck As String = "MainCheck" & vbCrLf & _
                                    "checkCovariate" & vbCrLf & WriteBool(Main.checkCovariate.Checked) & vbCrLf & _
                                    "nullModel" & vbCrLf & WriteBool(Main.RunNullModel.Checked) & vbCrLf & _
                                    "montecarlo" & vbCrLf & WriteBool(Main.MonteCarlo.Checked) & vbCrLf
        Dim CriteriaText As String = "CriteriaText" & vbCrLf & Main.SSCriteria.Text
        Dim SaveText As String = InstanceVariables & MainCheck & Tab0 & Tab1 & CriteriaCheck & CriteriaText
        FileOpen(1, filename, OpenMode.Output)
        PrintLine(1, SaveText)
        FileClose(1)
    End Sub

    

    Public Sub Reset()
        'Model instance variables
        ICCYValue = -1
        ESValue = -1
        ICCZValue = -1
        R2ClusterZValue = -1
        R2IndivZValue = -1
        'Sample Size Instance Variables
        NTValue = -1
        NCValue = -1
        NGValue = -1
        PTValue = -1.0
        NIndivValue = -1
        'Option variable
        OptionValue = -1
        'Power or width instance variables
        CILevelValue = -1
        PowerOrWidthValue = -1
        PowerValue = -1 ' 1 = Power, 2 = Width
        WidthValue = -1
        IsDegreeOfCertaintyValue = False
        DegreeOfCertaintyValue = -1
        'Cost instance variable
        TGroupCostValue = -1
        CGroupCostValue = -1
        TIndivCostValue = -1
        CIndivCostValue = -1
        TotalCostValue = -1
        StartingValue(0) = 0
        StartingValue(1) = 0
        StartingValue(2) = 0
        MonteCarloLine = 0
        MonteCarloMaxLine = 0
        Main.NTGroups.Text = ""
        Main.NCGroups.Text = ""
        Main.NIndiv.Text = ""
        Main.ICCY.Text = ""
        Main.ES.Text = ""
        Main.ICCZ.Text = ""
        Main.R2ClusterZ.Text = ""
        Main.R2IndivZ.Text = ""
        Main.WidthCI95.Text = ""
        Main.WidthCI99.Text = ""
        Main.Power.Text = ""
        Main.WidthCI95C.Text = ""
        Main.WidthCI99C.Text = ""
        Main.PowerC.Text = ""
        Main.ICCY2.Text = ""
        Main.ES2.Text = ""
        Main.ICCZ2.Text = ""
        Main.R2ClusterZ2.Text = ""
        Main.R2IndivZ2.Text = ""
        Main.NIndiv2.Text = ""
        Main.NTGroups2.Text = ""
        Main.NCGroups2.Text = ""
        Main.ObtainedPower.Text = ""
        Main.ObtainedWidth95.Text = ""
        Main.ObtainedWidth99.Text = ""
        Main.ObtainedPowerC.Text = ""
        Main.ObtainedWidth95C.Text = ""
        Main.ObtainedWidth99C.Text = ""
        Main.checkCovariate.Checked = False
        Main.RunNullModel.Checked = False
        If Main.IsMplusInstalled = True Then
            Main.MonteCarlo.Checked = True
        Else
            Main.MonteCarlo.Checked = False
        End If
        Main.CheckCovariateChanged()
        Main.SSCriteria.Text = "Please Click 'Define' to Specify Sample Size Characteristics"
        Main.TotalCostLabel.Text = "Total Cost"
        Main.IsOutputDegree.Enabled = False
        Main.IsOutputDegree2.Enabled = False
        Main.OutputCalcDegree.Enabled = False
        Main.OutputCalcDegree2.Enabled = False
        Main.OutputDegreeLevel.Enabled = False
        Main.OutputDegreeLevel2.Enabled = False
        Main.OutputDegreeLevel.Text = ""
        Main.OutputDegreeLevel2.Text = ""
        Main.IsOutputDegree.ForeColor = Color.Black
        Main.IsOutputDegree.ForeColor = Color.Black
        Main.Width95OutputLabel.ForeColor = Color.Black
        Main.Width99OutputLabel.ForeColor = Color.Black
        Main.Width95OutputLabelC.ForeColor = Color.Black
        Main.Width99OutputLabelC.ForeColor = Color.Black
        Main.Width95OutputLabel2.ForeColor = Color.Black
        Main.Width99OutputLabel2.ForeColor = Color.Black
        Main.Width95OutputLabel2C.ForeColor = Color.Black
        Main.Width99OutputLabel2C.ForeColor = Color.Black
        Main.TotalCost.Text = ""
        Main.IsOutputDegree.Checked = False
        Main.IsOutputDegree2.Checked = False
        Main.StatusNum.Text = "0"
        Main.ProgressBar1.Value = "0"
        Main.status.Text = "Have not started yet"
        Criteria.Option1.Checked = False
        Criteria.Option2.Checked = False
        Criteria.Option3.Checked = False
        Criteria.PCheckBox.Checked = False
        Criteria.NGCheckBox.Checked = False
        Criteria.NIndivCheckBox.Checked = False
        Criteria.ActivateNGNIndivCheckedBox(False)
        Criteria.ActivateCostCriteria(False)
        Criteria.ActivatePowerWidthCriteria(False)
        Criteria.ActivateTotalCost(False)
        Criteria.SSLabel.Text = "Please Choose Input Criteria Above"
        Criteria.PCheckBox.Enabled = False
        Criteria.CILevel.SelectedIndex = -1
        Criteria.TotalCost.Text = ""
        Criteria.TGroupCost.Text = ""
        Criteria.CGroupCost.Text = ""
        Criteria.TIndivCost.Text = ""
        Criteria.CIndivCost.Text = ""
        Criteria.PT.Text = ""
        Criteria.NGroups.Text = ""
        Criteria.NIndiv.Text = ""
        Criteria.Power.Text = ""
        Criteria.DesiredWidth.Text = ""
        Criteria.RPowerWidth1.Checked = False
        Criteria.RPowerWidth2.Checked = False
        Criteria.IsDegreeOfCertainty.SelectedIndex = -1
        Criteria.DegreeOfCertainty.Text = ""
    End Sub

    Public Function FindIndexPass(ByVal array() As Double, ByVal scalar As Double, ByVal over As Boolean)
        Dim Index As Integer = -1
        If over = True Then
            For i = 0 To UBound(array)
                If array(i) >= scalar Then Return i
            Next i
        ElseIf over = False Then
            For i = 0 To UBound(array)
                If array(i) <= scalar Then Return i
            Next i
        End If
        Return Index
    End Function

    Public Function ReadMplusOutputMod(ByVal fullfilename As String)
        Dim sr As StreamReader
        Try
            sr = New StreamReader(fullfilename)
        Catch ex As Exception
            System.Threading.Thread.Sleep(3000)
            sr = ReadMplusOutputMod(fullfilename)
        End Try
        Return sr
    End Function

    Private Function CBoolean(ByVal Text As String) As Boolean
        If CInt(Text) = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function WriteBool(ByVal bool As Boolean) As String
        If bool = True Then
            Return "1"
        Else
            Return "0"
        End If
    End Function

    Public Function checkMplus()
        Dim alltext As String = ""
        Dim uninstallkey As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" 'Import the program list into the alltext
        Dim rk As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(uninstallkey)
        Dim sk As Microsoft.Win32.RegistryKey
        Dim skname() = rk.GetSubKeyNames
        For counter As Integer = 0 To skname.Length - 1
            sk = rk.OpenSubKey(skname(counter))
            If sk.GetValue("DisplayName") <> "" Then
                alltext &= sk.GetValue("DisplayName") & ControlChars.CrLf & vbCrLf
            End If
        Next
        Dim numLine = 0 'Count number of line
        Dim charInFile As Integer = alltext.Length
        Dim i As Integer
        Dim letter As String
        For i = 0 To charInFile - 1
            letter = alltext.Substring(i, 1)
            If letter = Chr(13) Then
                numLine += 1
            End If
        Next i

        Dim fileArray(numLine) As String 'Build Array of Program list
        Dim currentLine As Integer = 1
        Dim Line As String = ""
        For i = 0 To charInFile - 1
            letter = alltext.Substring(i, 1)
            If letter = Chr(13) Then
                fileArray(currentLine - 1) = Line
                currentLine += 1
                Line = ""
            Else
                Line &= letter
            End If
        Next i
        Dim find As String = "Mplus" 'Find Mplus in the program list
        Dim numResult As Integer = 0
        Dim findResult(0 To 8) As Integer
        For i = 0 To numLine - 1
            If fileArray(i).Contains(find) Then
                findResult(numResult) = i
                numResult += 1
            End If
        Next i
        If numResult = 0 Then Return False 'If not found, return false which tell that Mplus is not available
        For i = 0 To numResult - 1
            If fileArray(findResult(i)).Contains("Multilevel") Or fileArray(findResult(i)).Contains("Combination") Then 'Check whether the Mplus is Multilevel/Combination version
                Return True
            End If
        Next i
        Return False
    End Function

    Public Function DetectProcess(ByVal processName As String)

        Dim Result As Boolean = False
        On Error GoTo ErrHandler

        Dim oWMI
        Dim ret
        Dim sService
        Dim oWMIServices
        Dim oWMIService
        Dim oServices
        Dim oService
        Dim servicename

        oWMI = GetObject("winmgmts:")
        oServices = oWMI.InstancesOf("win32_process")

        For Each oService In oServices

            servicename = LCase(Trim(CStr(oService.Name) & ""))

            If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
                Result = True
            End If

        Next

        oServices = Nothing
        oWMI = Nothing

ErrHandler:
        Err.Clear()
        Return Result
    End Function


    Public Sub addY(ByVal filename As String, ByVal ncase As Integer, ByVal ncol As Integer, ByVal yposition As Integer)
        Dim fullfilename As String = Path & "\" & filename & ".dat"
        Dim sr As StreamReader = New StreamReader(fullfilename)
        Dim allText As String = sr.ReadToEnd
        sr.Close()

        Dim charInFile As Integer = allText.Length
        Dim letter As String

        Dim fileArray(0 To ncase - 1) As String
        Dim currentLine As Integer = 0
        Dim Line As String = ""
        For i = 0 To charInFile - 1
            letter = allText.Substring(i, 1)
            If letter = Chr(13) Then
                fileArray(currentLine) = Line
                currentLine += 1
                Line = ""
            Else
                Line &= letter
            End If
        Next i

        Dim data(0 To ncase - 1, 0 To ncol - 1) As String
        Dim charInLine As Integer
        Dim LastValue As String, CurrentValue As String
        Dim StartPoint As Integer
        Dim CurrentColumn As Integer
        For i = 0 To ncase - 1
            charInLine = fileArray(i).Length
            StartPoint = 1
            CurrentColumn = 0
            While fileArray(i).Substring(StartPoint, 1) = Chr(32) Or fileArray(i).Substring(StartPoint, 1) = ""
                StartPoint += 1
            End While
            data(i, 0) = fileArray(i).Substring(StartPoint, 1)
            LastValue = data(i, 0)
            For j = StartPoint + 1 To charInLine - 1
                CurrentValue = fileArray(i).Substring(j, 1)
                If CurrentValue = Chr(32) Then
                    If LastValue <> Chr(32) Then
                        CurrentColumn += 1
                    End If
                Else
                    data(i, CurrentColumn) &= CurrentValue
                End If
                LastValue = CurrentValue
            Next j
        Next i

        'Dim code As String = ""
        'For i = 0 To ncase - 1
        '    code &= fileArray(i)
        '    code &= "   "
        '    code &= data(i, yposition - 1)
        '    code &= vbCrLf
        'Next i


        My.Computer.FileSystem.DeleteFile(Path & "\" & filename & ".dat")

        Dim textOut As New StreamWriter(New FileStream(Path & "\" & filename & ".dat", FileMode.Create, FileAccess.Write))

        For i = 0 To ncase - 1
            textOut.Write(fileArray(i) & "   ")
            textOut.WriteLine(data(i, yposition - 1))
        Next i

        textOut.Close()

        'FileOpen(1, Path & "\" & filename & ".dat", OpenMode.Output)
        'PrintLine(1, code)
        'FileClose(1)
    End Sub

    Public Sub WriteOutput(ByVal filePath As String)
        Dim textOut As New StreamWriter(New FileStream(filePath, FileMode.Create, FileAccess.Write))
        Dim PostHoc As Boolean = Main.TabControl1.SelectedIndex = 0
        textOut.WriteLine(Banner & vbCrLf)
        textOut.Write("TYPE OF ANALYSIS:" & vbTab & "Two-level Cluster Randomized Design ")
        If PostHoc Then
            textOut.WriteLine("(Post Hoc analysis)")
        Else
            textOut.WriteLine("(A priori analysis)")
        End If
        textOut.WriteLine()
        WriteSpecifiedInput(textOut, PostHoc)
        textOut.WriteLine()
        WriteResult(textOut, PostHoc)
        textOut.WriteLine()
        If Main.MonteCarlo.Checked = True And PostHoc = False And (Criteria.Option1.Checked = True Or Criteria.Option2.Checked = True) Then
            textOut.WriteLine("######################A PRIORI MONTE CARLO HISTORY######################")
            textOut.WriteLine()
            textOut.Write("#Treatment Groups" & vbTab & "#Control Groups" & vbTab & "Group Size" & vbTab)
            If (Criteria.RPowerWidth1.Checked = True) Then
                textOut.Write("Power")
            Else
                If Criteria.IsDegreeOfCertainty.SelectedIndex = 1 Then textOut.Write("Proportion of ")
                If Criteria.CILevel.SelectedIndex = 0 Then textOut.Write("95% CI ")
                If Criteria.CILevel.SelectedIndex = 1 Then textOut.Write("99% CI ")
                If Criteria.IsDegreeOfCertainty.SelectedIndex = 0 Then textOut.Write("Width")
                If Criteria.IsDegreeOfCertainty.SelectedIndex = 1 Then textOut.Write("Width Less than " & Criteria.DesiredWidth.Text)
            End If
            If Criteria.Option1.Checked = False Then
                textOut.WriteLine(vbTab & "Cost")
            Else
                textOut.WriteLine()
            End If
            textOut.WriteLine(StartingValue(0).ToString("n0") & vbTab & vbTab & vbTab & StartingValue(1).ToString("n0") & vbTab & vbTab & StartingValue(2).ToString("n0") & vbTab & vbTab & "(STARTING VALUE)")
            Dim cost As Double
            For i = 0 To MonteCarloLine - 1
                textOut.Write(MonteCarloArchive(i, 0).ToString("n0") & vbTab & vbTab & vbTab & MonteCarloArchive(i, 1).ToString("n0") & vbTab & vbTab & MonteCarloArchive(i, 2).ToString("n0") & vbTab & vbTab & MonteCarloArchive(i, 3).ToString("f4") & vbTab)
                If Criteria.Option1.Checked = False Then
                    cost = FindTotalCostExact(CInt(MonteCarloArchive(i, 0)), CInt(MonteCarloArchive(i, 1)), CInt(MonteCarloArchive(i, 2)), TGroupCostValue, TIndivCostValue, CGroupCostValue, CIndivCostValue)
                    textOut.WriteLine(cost.ToString("f2"))
                Else
                    textOut.WriteLine()
                End If
            Next i
        End If
        textOut.Close()
    End Sub

    Private Sub WriteSpecifiedInput(ByVal file As StreamWriter, ByVal postHoc As Boolean)
        If postHoc Then
            file.WriteLine("################################INPUTS################################" & vbCrLf)
            file.WriteLine("Number of Treatment Groups:" & vbTab & vbTab & vbTab & NTValue.ToString("n0"))
            file.WriteLine("Number of Control Groups:" & vbTab & vbTab & vbTab & NCValue.ToString("n0"))
            file.WriteLine("Number of Individual in Each Group:" & vbTab & vbTab & NIndivValue.ToString("n0"))
            file.WriteLine("Effect Size:" & vbTab & vbTab & vbTab & vbTab & vbTab & ESValue.ToString("f3"))
            file.WriteLine("Intraclass Correlation of Y:" & vbTab & vbTab & vbTab & ICCYValue.ToString("f3"))
            If Main.checkCovariate.Checked = True Then
                file.WriteLine("Covariate")
                If Main.ClusterOnlyZ.Checked = True Then
                    file.WriteLine(vbTab & "Cluster-level covariate")
                Else
                    file.WriteLine(vbTab & "Individual-level covariate")
                    file.WriteLine(vbTab & "Intraclass Correlation of the covariate:" & vbTab & vbTab & vbTab & ICCZValue.ToString("f3"))
                End If
                file.WriteLine(vbTab & "Proportion of Error Variance Explained in Cluster Level:" & vbTab & R2ClusterZValue.ToString("f3"))
                If Main.ClusterOnlyZ.Checked = False Then file.WriteLine(vbTab & "Proportion of Error Variance Explained in Individual Level:" & vbTab & R2IndivZValue.ToString("f3"))
            Else
                file.WriteLine("No Covariate")
            End If
        Else
            file.WriteLine("################################CRITERIA################################" & vbCrLf)
            file.Write("Criterion Used:" & vbTab)
            If Criteria.Option1.Checked = True Then
                file.WriteLine("1. " & Criteria.Option1.Text)
            ElseIf Criteria.Option2.Checked = True Then
                file.WriteLine("2. " & Criteria.Option2.Text)
            ElseIf Criteria.Option3.Checked = True Then
                file.WriteLine("3. " & Criteria.Option3.Text)
            End If
            If Criteria.Option1.Checked = True Or Criteria.Option2.Checked = True Then
                file.WriteLine("Power/CIES")
                If Criteria.RPowerWidth1.Checked = True Then
                    file.WriteLine(vbTab & "Power:" & vbTab & PowerValue.ToString("f3"))
                End If
                If Criteria.RPowerWidth2.Checked = True Then
                    file.WriteLine(vbTab & "Width of CI of ES:" & vbTab & WidthValue.ToString("f3"))
                    file.WriteLine(vbTab & "CI Level:" & vbTab & CILevelValue)
                    file.Write(vbTab & "Degree of Certainty:" & vbTab)
                    If (IsDegreeOfCertaintyValue = True) Then
                        file.WriteLine("Yes (" & DegreeOfCertaintyValue.ToString("f3") & ")")
                    Else
                        file.WriteLine("No")
                    End If
                End If
            End If
            If Criteria.Option2.Checked = True Or Criteria.Option3.Checked = True Then
                file.WriteLine("Cost")
                file.WriteLine(vbTab & "Treatment Group Cost:" & vbTab & vbTab & TGroupCostValue.ToString("f3"))
                file.WriteLine(vbTab & "Treatment Individual Cost:" & vbTab & TIndivCostValue.ToString("f3"))
                file.WriteLine(vbTab & "Control Group Cost:" & vbTab & vbTab & CGroupCostValue.ToString("f3"))
                file.WriteLine(vbTab & "Control Individual Cost:" & vbTab & CIndivCostValue.ToString("f3"))
            End If
            If Criteria.Option3.Checked = True Then file.WriteLine("Total Budget:" & vbTab & TotalCostValue.ToString("f3"))
            If Criteria.PCheckBox.Checked = True Or Criteria.NGCheckBox.Checked = True Or Criteria.NIndivCheckBox.Checked = True Then
                file.WriteLine("Pre-specify result")
                If Criteria.PCheckBox.Checked = True Then file.WriteLine(vbTab & "Proportion of Treatment Groups:" & vbTab & Criteria.PT.Text)
                If Criteria.NGCheckBox.Checked = True Then file.WriteLine(vbTab & "Number of Groups:" & vbTab & vbTab & vbTab & Criteria.NGroups.Text)
                If Criteria.NIndivCheckBox.Checked = True Then file.WriteLine(vbTab & "Number of Individuals in each Cluster:" & vbTab & vbTab & Criteria.NIndiv.Text)
            End If
            file.WriteLine("Effect Size:" & vbTab & vbTab & vbTab & vbTab & vbTab & ESValue.ToString("f3"))
            file.WriteLine("Intraclass Correlation of Y:" & vbTab & vbTab & vbTab & ICCYValue.ToString("f3"))
            If Main.checkCovariate.Checked = True Then
                file.WriteLine("Covariate")
                If Main.ClusterOnlyZ.Checked = True Then
                    file.WriteLine(vbTab & "Cluster-level covariate")
                Else
                    file.WriteLine(vbTab & "Individual-level covariate")
                    file.WriteLine(vbTab & "Intraclass Correlation of the covariate:" & vbTab & vbTab & vbTab & ICCZValue.ToString("f3"))
                End If
                file.WriteLine(vbTab & "Proportion of Error Variance Explained in Cluster Level:" & vbTab & R2ClusterZValue.ToString("f3"))
                If Main.ClusterOnlyZ.Checked = False Then file.WriteLine(vbTab & "Proportion of Error Variance Explained in Individual Level:" & vbTab & R2IndivZValue.ToString("f3"))
            Else
                file.WriteLine("No Covariate")
            End If
        End If
        file.Write("Used a priori Monte Carlo simulation:" & vbTab)
        If Main.MonteCarlo.Checked = True Then
            file.WriteLine("Yes")
            file.WriteLine("Number of Replications:" & vbTab & NRep.ToString("n0"))
        Else
            file.WriteLine("No")
        End If
    End Sub

    Private Sub WriteResult(ByVal file As StreamWriter, ByVal postHoc As Boolean)
        file.WriteLine("################################RESULTS################################")
        file.WriteLine()
        Dim CertaintyLevel() As Double = {0.5, 0.6, 0.7, 0.8, 0.9, 0.95, 0.99}
        Dim runDegree() As Double
        If postHoc Then
            If Main.checkCovariate.Checked = False Then
                file.WriteLine("Power:" & vbTab & Main.Power.Text)
                file.WriteLine("95% Width of CI of ES:" & vbTab & Main.WidthCI95.Text)
                file.WriteLine("99% Width of CI of ES:" & vbTab & Main.WidthCI99.Text)
                If Main.MonteCarlo.Checked = True Then
                    file.WriteLine("Degree of Certainty")
                    file.WriteLine(vbTab & "Level" & vbTab & "95% CI" & vbTab & "99% CI")
                    For i = 0 To UBound(CertaintyLevel)
                        runDegree = FindDegreeOfCertainty(CertaintyLevel(i), False, False)
                        file.WriteLine(vbTab & CertaintyLevel(i).ToString("f2") & vbTab & runDegree(0).ToString("f3") & vbTab & runDegree(1).ToString("f3"))
                    Next i
                End If
            Else
                file.WriteLine("Result when the Model Includes the Covariate")
                file.WriteLine(vbTab & "Power:" & vbTab & Main.PowerC.Text)
                file.WriteLine(vbTab & "95% Width of CI of ES:" & vbTab & Main.WidthCI95C.Text)
                file.WriteLine(vbTab & "99% Width of CI of ES:" & vbTab & Main.WidthCI99C.Text)
                If Main.MonteCarlo.Checked = True Then
                    file.WriteLine(vbTab & "Degree of Certainty")
                    file.WriteLine(vbTab & vbTab & "Level" & vbTab & "95% CI" & vbTab & "99% CI")
                    For i = 0 To UBound(CertaintyLevel)
                        runDegree = FindDegreeOfCertainty(CertaintyLevel(i), True, False)
                        file.WriteLine(vbTab & vbTab & CertaintyLevel(i).ToString("f2") & vbTab & runDegree(0).ToString("f3") & vbTab & runDegree(1).ToString("f3"))
                    Next i
                End If
                If Main.RunNullModel.Checked = True Then
                    file.WriteLine("Result when the Model Excludes the Covariate")
                    file.WriteLine(vbTab & "Power:" & vbTab & Main.Power.Text)
                    file.WriteLine(vbTab & "95% Width of CI of ES:" & vbTab & Main.WidthCI95.Text)
                    file.WriteLine(vbTab & "99% Width of CI of ES:" & vbTab & Main.WidthCI99.Text)
                    If Main.MonteCarlo.Checked = True Then
                        file.WriteLine(vbTab & "Degree of Certainty")
                        file.WriteLine(vbTab & vbTab & "Level" & vbTab & "95% CI" & vbTab & "99% CI")
                        For i = 0 To UBound(CertaintyLevel)
                            runDegree = FindDegreeOfCertainty(CertaintyLevel(i), True, True)
                            file.WriteLine(vbTab & vbTab & CertaintyLevel(i).ToString("f2") & vbTab & runDegree(0).ToString("f3") & vbTab & runDegree(1).ToString("f3"))
                        Next i
                    End If
                End If
            End If
        Else
            file.WriteLine("Number of Treatment Groups:" & vbTab & Main.NTGroups2.Text)
            file.WriteLine("Number of Control Groups:" & vbTab & Main.NCGroups2.Text)
            file.WriteLine("Number of Individuals in Each Group:" & vbTab & Main.NIndiv2.Text)
            If Criteria.Option1.Checked = False Then file.WriteLine("Total Cost:" & vbTab & Main.TotalCost.Text)
            If Main.checkCovariate.Checked = False Then
                file.WriteLine("Power:" & vbTab & Main.ObtainedPower.Text)
                file.WriteLine("95% Width of CI of ES:" & vbTab & Main.ObtainedWidth95.Text)
                file.WriteLine("99% Width of CI of ES:" & vbTab & Main.ObtainedWidth99.Text)
                If Main.MonteCarlo.Checked = True Then
                    file.WriteLine("Degree of Certainty")
                    file.WriteLine(vbTab & "Level" & vbTab & "95% CI" & vbTab & "99% CI")
                    For i = 0 To UBound(CertaintyLevel)
                        runDegree = FindDegreeOfCertainty(CertaintyLevel(i), False, False)
                        file.WriteLine(vbTab & CertaintyLevel(i).ToString("f2") & vbTab & runDegree(0).ToString("f3") & vbTab & runDegree(1).ToString("f3"))
                    Next i
                End If
            Else
                file.WriteLine("Result when the Model Includes the Covariate")
                file.WriteLine(vbTab & "Power:" & vbTab & Main.ObtainedPowerC.Text)
                file.WriteLine(vbTab & "95% Width of CI of ES:" & vbTab & Main.ObtainedWidth95C.Text)
                file.WriteLine(vbTab & "99% Width of CI of ES:" & vbTab & Main.ObtainedWidth99C.Text)
                If Main.MonteCarlo.Checked = True Then
                    file.WriteLine(vbTab & "Degree of Certainty")
                    file.WriteLine(vbTab & vbTab & "Level" & vbTab & "95% CI" & vbTab & "99% CI")
                    For i = 0 To UBound(CertaintyLevel)
                        runDegree = FindDegreeOfCertainty(CertaintyLevel(i), True, False)
                        file.WriteLine(vbTab & vbTab & CertaintyLevel(i).ToString("f2") & vbTab & runDegree(0).ToString("f3") & vbTab & runDegree(1).ToString("f3"))
                    Next i
                End If
                If Main.RunNullModel.Checked = True Then
                    file.WriteLine("Result when the Model Excludes the Covariate")
                    file.WriteLine(vbTab & "Power:" & vbTab & Main.ObtainedPower.Text)
                    file.WriteLine(vbTab & "95% Width of CI of ES:" & vbTab & Main.ObtainedWidth95.Text)
                    file.WriteLine(vbTab & "99% Width of CI of ES:" & vbTab & Main.ObtainedWidth99.Text)
                    If Main.MonteCarlo.Checked = True Then
                        file.WriteLine(vbTab & "Degree of Certainty")
                        file.WriteLine(vbTab & vbTab & "Level" & vbTab & "95% CI" & vbTab & "99% CI")
                        For i = 0 To UBound(CertaintyLevel)
                            runDegree = FindDegreeOfCertainty(CertaintyLevel(i), True, True)
                            file.WriteLine(vbTab & vbTab & CertaintyLevel(i).ToString("f2") & vbTab & runDegree(0).ToString("f3") & vbTab & runDegree(1).ToString("f3"))
                        Next i
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub WriteMonteCarloArchieve(ByVal NJ As Integer, ByVal pT As Double, ByVal NIndiv As Integer, ByVal Statistics As Double)
        If MonteCarloLine = MonteCarloMaxLine Then
            Dim SaveResult(,) As Double = MonteCarloArchive
            MonteCarloMaxLine += 1000
            Dim NewArray(0 To (MonteCarloMaxLine - 1), 0 To 3) As Double
            MonteCarloArchive = NewArray
            For i = 0 To (MonteCarloLine - 1)
                MonteCarloArchive(i, 0) = SaveResult(i, 0)
                MonteCarloArchive(i, 1) = SaveResult(i, 1)
                MonteCarloArchive(i, 2) = SaveResult(i, 2)
                MonteCarloArchive(i, 3) = SaveResult(i, 3)
            Next i
        End If
        MonteCarloArchive(MonteCarloLine, 0) = Math.Round(NJ * pT)
        MonteCarloArchive(MonteCarloLine, 1) = Math.Round(NJ * (1 - pT))
        MonteCarloArchive(MonteCarloLine, 2) = NIndiv
        MonteCarloArchive(MonteCarloLine, 3) = Statistics
        If (NJ > (MonteCarloArchive(MonteCarloLine, 0) + MonteCarloArchive(MonteCarloLine, 1))) Then
            MonteCarloArchive(MonteCarloLine, 0) += 1
        ElseIf (NJ < (MonteCarloArchive(MonteCarloLine, 0) + MonteCarloArchive(MonteCarloLine, 1))) Then
            MonteCarloArchive(MonteCarloLine, 1) -= 1
        End If
        MonteCarloLine += 1
    End Sub

    Private Function FindDegreeOfCertainty(ByVal CertaintyLevel As Double, ByVal Covariate As Boolean, Optional ByVal CovariateWithNull As Boolean = False)
        Dim Position As Integer = Math.Ceiling(CertaintyLevel * NRep) - 1
        Dim Result(0 To 1) As Double
        If Covariate And (Not CovariateWithNull) Then
            Result(0) = Storage.Width(Position, 2)
            Result(1) = Storage.Width(Position, 3)
        Else
            Result(0) = Storage.Width(Position, 0)
            Result(1) = Storage.Width(Position, 1)
        End If
        Return Result
    End Function

    Public Function PrintWidth(ByVal ES As Double, ByVal width As Double) As String
        Dim result As String = width.ToString("f3")
        Dim MarginError As Double = width / 2
        Dim lower As Double = ES - MarginError
        Dim upper As Double = ES + MarginError
        result &= "  (" & lower.ToString("f3") & ", " & upper.ToString("f3") & ")"
        Return result
    End Function
End Module
