Public Class StartUp

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.DialogResult = DialogResult.OK  'Continue to the main program

    End Sub

    Private Sub StartUp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Main.IsMplusInstalled = True Then 'Check whether the Mplus with multilevel or combination is installed
            Label4.Text = "" 'DELETE THE RED TEXT.
            Label2.Text = ""
            Label5.Text = ""
            Label6.Text = ""
            Label3.Location = New Point(38, 82)
        End If
    End Sub
End Class