Public Class Form2

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Form1.Show()
        Me.Visible = False
        Timer1.Enabled = False
    End Sub

    

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Dim pt As String
        Application.DoEvents()
        pt = Replace(Application.StartupPath + "\inner_keywords.apl", "\\", "\")
        Kill(pt)
        System.IO.File.WriteAllText(pt, TextBox1.Text)
        Application.DoEvents()


    End Sub

  
End Class