Public Class Options
    Dim folderBrowserDialog1 As FolderBrowserDialog
    Private Sub Okay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Okay.Click
        Dim proceed As Boolean = checkReplication()
        If proceed = True Then
            NRep = CInt(SetReplication.Text)
            Me.DialogResult = DialogResult.OK
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    Me.folderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    '    folderBrowserDialog1.ShowDialog()
    '    Me.folderBrowserDialog1.Description = "Select the directory that you want the program making a simulation."
    '    pathOptions = folderBrowserDialog1.SelectedPath
    'End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        My.Settings.Reset()
        SetReplication.Text = Storage.DefaultReplication
    End Sub

    Private Function checkReplication()
        If Main.IsMplusInstalled = True Then
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
                    Main.MonteCarlo.Enabled = False
                    Main.MonteCarlo.Checked = False
                    MplusExist = True
                End If
            End While
        End If
        If IsIntegerLBound(SetReplication, 10, "Number of replication") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Options_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetReplication.Text = NRep
    End Sub


End Class