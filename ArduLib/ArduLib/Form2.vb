Public Class Form2

  
    

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Dim pt As String
        Application.DoEvents()
        pt = Replace(Application.StartupPath + "\inner_keywords.apl", "\\", "\")
        Kill(pt)
        System.IO.File.WriteAllText(pt, TextBox1.Text)
        Application.DoEvents()


    End Sub


    Private Sub Timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer.Tick
        Form1.Show()
        Me.Visible = False
        Timer.Enabled = False
    End Sub
End Class