Public Class Form1
    '  Asemani Personal Laboratory
    ' Post.asemani@gmail.com
    Dim alphabet_sum As String
    Dim word_temp As String
    Dim flag1 As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        TextBox1.Text = Replace(TextBox1.Text, Label3.Text, TextBox4.Text)
        TextBox2.Text = Replace(TextBox2.Text, Label3.Text, TextBox4.Text)
        TextBox3.Text = Replace(TextBox3.Text, Label3.Text, TextBox4.Text)
        TextBox5.Text = Replace(TextBox5.Text, Label3.Text, TextBox4.Text)
        Label3.Text = TextBox4.Text
        Call full_color_parser(TextBox1, Color.Brown)
        Call full_color_parser(TextBox2, Color.Brown)
        Call full_color_parser(TextBox5, Color.Brown)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim msg As String
        msg = MsgBox("are you sure?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "new library?")
        If msg = vbYes Then
            TextBox1.Text = txt_header_temp.Text
            TextBox2.Text = txt_source_temp.Text
            TextBox3.Text = txt_key_temp.Text
            TextBox5.Text = txt_exam_temp.Text
            TextBox6.Text = "NewExample"
            TextBox4.Text = "New"
            Label3.Text = "New"
            Call full_color_parser(TextBox1, Color.Brown)
            Call full_color_parser(TextBox2, Color.Brown)
            Call full_color_parser(TextBox5, Color.Brown)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        SaveFileDialog1.Title = "Save header file to:"
        SaveFileDialog1.FileName = TextBox4.Text
        SaveFileDialog1.Filter = "Arduino Library Header file (*.h)|*.h|All files (*.*)|*.*"
        SaveFileDialog1.ShowDialog()
        lbl_header_file.Text = SaveFileDialog1.FileName
        If lbl_header_file.Text = "" Then Exit Sub

        System.IO.File.WriteAllText(lbl_header_file.Text, TextBox1.Text)
        MessageBox.Show("Header file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        SaveFileDialog1.Title = "Save source file to:"
        SaveFileDialog1.FileName = TextBox4.Text
        SaveFileDialog1.Filter = "Arduino Library Source file (*.cpp)|*.cpp|All files (*.*)|*.*"
        SaveFileDialog1.ShowDialog()
        lbl_source_file.Text = SaveFileDialog1.FileName
        If lbl_source_file.Text = "" Then Exit Sub
        System.IO.File.WriteAllText(lbl_source_file.Text, TextBox2.Text)
        MessageBox.Show("Source file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        SaveFileDialog1.Title = "Save keyword file to:"
        SaveFileDialog1.FileName = "keywords"
        SaveFileDialog1.Filter = "Arduino Keywords Source file (*.txt)|*.txt|All files (*.*)|*.*"
        SaveFileDialog1.ShowDialog()
        lbl_keywoeds_file.Text = SaveFileDialog1.FileName
        If lbl_keywoeds_file.Text = "" Then Exit Sub

        System.IO.File.WriteAllText(lbl_keywoeds_file.Text, TextBox3.Text)
        MessageBox.Show("Keywords file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        SaveFileDialog1.Title = "Save source file to:"
        SaveFileDialog1.FileName = TextBox6.Text
        SaveFileDialog1.Filter = "Arduino Source Code file (*.ino)|*.ino|All files (*.*)|*.*"
        SaveFileDialog1.ShowDialog()
        lbl_example_file.Text = SaveFileDialog1.FileName
        If lbl_example_file.Text = "" Then Exit Sub

        System.IO.File.WriteAllText(lbl_example_file.Text, TextBox5.Text)

        MessageBox.Show("Example file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'On Error Resume Next
        Dim dirpath As String


        FolderBrowserDialog1.ShowDialog()
        dirpath = FolderBrowserDialog1.SelectedPath
        If System.IO.Directory.Exists(dirpath + "\" + TextBox4.Text) = True Then
            Dim t As String
            t = MsgBox(dirpath + "\" + TextBox4.Text + vbCrLf + "do you want to replace it?", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "same detected!")
            If t = vbOK Then
                System.IO.Directory.Delete(dirpath + "\" + TextBox4.Text, True)
            Else
                Exit Sub
            End If
        End If
        If dirpath <> "" Then
            MkDir(dirpath + "\" + TextBox4.Text)

            'header
            System.IO.File.WriteAllText(dirpath + "\" + TextBox4.Text + "\" + TextBox4.Text + ".h", TextBox1.Text)

            'source
            System.IO.File.WriteAllText(dirpath + "\" + TextBox4.Text + "\" + TextBox4.Text + ".cpp", TextBox2.Text)
            'keywords
            System.IO.File.WriteAllText(dirpath + "\" + TextBox4.Text + "\" + "keywords.txt", TextBox3.Text)
            'example

            MkDir(dirpath + "\" + TextBox4.Text + "\examples")
            MkDir(dirpath + "\" + TextBox4.Text + "\examples\" + TextBox6.Text)
            System.IO.File.WriteAllText(dirpath + "\" + TextBox4.Text + "\examples\" + TextBox6.Text + "\" + TextBox6.Text + ".ino", TextBox5.Text)

            MessageBox.Show("Library made successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub fixer(ByVal pat As String)
        Shell(Application.StartupPath + "\fixer.exe " + pat, vbNormalFocus)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        End
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call load_inner_ketwords()
        Call full_color_parser(TextBox1, Color.Brown)
        Call full_color_parser(TextBox2, Color.Brown)
        Call full_color_parser(TextBox5, Color.Brown)
    End Sub

    Private Sub load_inner_ketwords()
        On Error Resume Next
        Dim tmp As Object
        Dim pt As String
        pt = Replace(Application.StartupPath + "\inner_keywords.apl", "\\", "\")
        FileOpen(1, pt, OpenMode.Input)
        While Not EOF(1)
            Input(1, tmp)
            ListBox1.Items.Add(tmp)
        End While
        FileClose(1)

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        On Error Resume Next
        Dim header_path As String

        OpenFileDialog1.Title = "Select a Header file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library Header files (*.h)|*.h|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        header_path = OpenFileDialog1.FileName
        TextBox4.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")

        TextBox1.Text = System.IO.File.ReadAllText(header_path)
        Call full_color_parser(TextBox1, Color.Brown)
        Call full_color_parser(TextBox2, Color.Brown)
        Call full_color_parser(TextBox5, Color.Brown)
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        On Error Resume Next
        Dim source_path As String

        OpenFileDialog1.Title = "Select a Source file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library Source files (*.cpp)|*.cpp|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        source_path = OpenFileDialog1.FileName


        TextBox2.Text = System.IO.File.ReadAllText(source_path)
        Call full_color_parser(TextBox1, Color.Brown)
        Call full_color_parser(TextBox2, Color.Brown)
        Call full_color_parser(TextBox5, Color.Brown)
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        On Error Resume Next
        On Error Resume Next
        Dim keyword_path As String

        OpenFileDialog1.Title = "Select a KeyWords file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library KeyWords files (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        keyword_path = OpenFileDialog1.FileName

        TextBox3.Text = System.IO.File.ReadAllText(keyword_path)
        Call full_color_parser(TextBox1, Color.Brown)
        Call full_color_parser(TextBox2, Color.Brown)
        Call full_color_parser(TextBox5, Color.Brown)
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        On Error Resume Next
        Dim exam_path As String

        OpenFileDialog1.Title = "Select a Example file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library Example files (*.ino)|*.ino|PDE files (*.pde)|*.pde|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        exam_path = OpenFileDialog1.FileName


        TextBox5.Text = System.IO.File.ReadAllText(exam_path)
        Call full_color_parser(TextBox1, Color.Brown)
        Call full_color_parser(TextBox2, Color.Brown)
        Call full_color_parser(TextBox5, Color.Brown)
    End Sub


    Private Sub Label5_Click_4(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        Shell("explorer.exe https://www.arduino.cc/en/Hacking/LibraryTutorial", vbNormalFocus)
    End Sub



    Private Function FindIt(ByRef Box As RichTextBox, ByVal Search As String, ByVal Color As Color, Optional ByVal Start As Int32 = 0) As Int32
        Dim retval As Int32      'Instr returns a long
        Dim Source As String 'variable used in Instr
        Try

            Source = Box.Text   'put the text to search into the variable

            retval = Source.IndexOf(Search, Start) 'do the first search,
            'starting at the beginning
            'of the text

            If retval <> -1 Then 'there is at least one more occurrence of
                'the string

                'the RichTextBox doesn't support multiple active selections, so
                'this section marks the occurrences of the search string by
                'making them Bold and Red

                With Box
                    .SelectionStart = retval
                    .SelectionLength = Search.Length
                    .SelectionColor = Color
                    .DeselectAll() 'this line removes the selection highlight
                End With

                Start = retval + Search.Length 'move the starting point past the
                'first occurrence

                'FindIt calls itself with new arguments
                'this is what makes it Recursive
                FindIt = 1 + FindIt(Box, Search, Color, Start)
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try

        Return retval

    End Function


    Private Sub Button14_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Call full_color_parser(TextBox2, Color.Brown)
    End Sub

    Private Sub Button13_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Call full_color_parser(TextBox1, Color.Brown)
    End Sub


    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Call full_color_parser(TextBox5, Color.Brown)
    End Sub

    Private Function full_color_parser(ByRef TB As RichTextBox, ByVal Col As Color) As Int32
        On Error Resume Next
        Dim tmp2 As String
        'search in internal keywords
        Dim current_position As Integer
        current_position = TB.SelectionStart
        TB.Select(0, TB.TextLength)
        TB.SelectionColor = Color.Black
        TB.SelectionStart = current_position
        For i As Integer = 0 To ListBox1.Items.Count - 1
            tmp2 = ListBox1.Items.Item(i)

            FindIt(TB, tmp2, Col)

        Next
        TB.SelectionStart = current_position
    End Function

   
    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress



        If e.KeyChar = "A" Or e.KeyChar = "B" Or e.KeyChar = "C" Or e.KeyChar = "D" Or e.KeyChar = "E" Or e.KeyChar = "F" Or e.KeyChar = "G" Or e.KeyChar = "H" Or _
                        e.KeyChar = "K" Or e.KeyChar = "L" Or e.KeyChar = "M" Or e.KeyChar = "N" Or e.KeyChar = "O" Or e.KeyChar = "P" Or e.KeyChar = "Q" Or e.KeyChar = "R" Or _
                        e.KeyChar = "S" Or e.KeyChar = "T" Or e.KeyChar = "W" Or e.KeyChar = "X" Or e.KeyChar = "Y" Or e.KeyChar = "Z" Or e.KeyChar = "a" Or e.KeyChar = "b" Or _
                        e.KeyChar = "c" Or e.KeyChar = "d" Or e.KeyChar = "e" Or e.KeyChar = "f" Or e.KeyChar = "g" Or e.KeyChar = "h" Or e.KeyChar = "i" Or e.KeyChar = "j" Or _
                        e.KeyChar = "k" Or e.KeyChar = "l" Or e.KeyChar = "m" Or e.KeyChar = "n" Or e.KeyChar = "o" Or e.KeyChar = "p" Or e.KeyChar = "q" Or e.KeyChar = "r" Or _
                        e.KeyChar = "s" Or e.KeyChar = "t" Or e.KeyChar = "u" Or e.KeyChar = "U" Or e.KeyChar = "v" Or e.KeyChar = "w" Or e.KeyChar = "x" Or e.KeyChar = "y" Or _
                        e.KeyChar = "z" Or e.KeyChar = "I" Or e.KeyChar = "J" Or e.KeyChar = "V" Then
            alphabet_sum = alphabet_sum & e.KeyChar

        Else
            word_temp = alphabet_sum
        End If



    End Sub


    Private Sub find_word(ByVal word As String)
        Dim tmp2 As String
        Dim tmp_word As String
        tmp_word = word

        For i As Integer = 0 To ListBox1.Items.Count - 1
            tmp2 = ListBox1.Items.Item(i)
            If tmp_word = tmp2 Then
                flag1 = True
                ListBox1.SelectedItem = 0
                Exit For
            Else
                flag1 = False
            End If
        Next

    End Sub

    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyUp
        On Error Resume Next
        If e.KeyCode = 32 Then
            alphabet_sum = ""
            word_temp = ""
            TextBox1.SelectionColor = Color.Black
        End If

    End Sub

    Private Sub TextBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox1.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg.Visible = True

        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim startingPoint As Integer = -1
        Dim st As Integer
        st = TextBox1.SelectionStart

        Call find_word(word_temp)

        If flag1 = True Then

            If word_temp = "" Then Exit Sub
            Application.DoEvents()
            Do
                startingPoint = TextBox1.Find(word_temp, startingPoint + 1, RichTextBoxFinds.None)
                If (startingPoint >= 0) Then
                    TextBox1.SelectionStart = startingPoint
                    TextBox1.SelectionLength = word_temp.Length
                    TextBox1.SelectionColor = Color.Brown
                End If
            Loop Until startingPoint < 0
            TextBox1.DeselectAll()
            TextBox1.SelectionStart = st
            TextBox1.SelectionColor = Color.Black
            alphabet_sum = ""
            word_temp = ""
            flag1 = False
        End If

    End Sub

    Private Sub TextBox3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyUp
        If e.KeyCode = Keys.Tab Then
            TextBox3.SelectedText = Chr(9)
        End If

    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = "A" Or e.KeyChar = "B" Or e.KeyChar = "C" Or e.KeyChar = "D" Or e.KeyChar = "E" Or e.KeyChar = "F" Or e.KeyChar = "G" Or e.KeyChar = "H" Or _
                e.KeyChar = "K" Or e.KeyChar = "L" Or e.KeyChar = "M" Or e.KeyChar = "N" Or e.KeyChar = "O" Or e.KeyChar = "P" Or e.KeyChar = "Q" Or e.KeyChar = "R" Or _
                e.KeyChar = "S" Or e.KeyChar = "T" Or e.KeyChar = "W" Or e.KeyChar = "X" Or e.KeyChar = "Y" Or e.KeyChar = "Z" Or e.KeyChar = "a" Or e.KeyChar = "b" Or _
                e.KeyChar = "c" Or e.KeyChar = "d" Or e.KeyChar = "e" Or e.KeyChar = "f" Or e.KeyChar = "g" Or e.KeyChar = "h" Or e.KeyChar = "i" Or e.KeyChar = "j" Or _
                e.KeyChar = "k" Or e.KeyChar = "l" Or e.KeyChar = "m" Or e.KeyChar = "n" Or e.KeyChar = "o" Or e.KeyChar = "p" Or e.KeyChar = "q" Or e.KeyChar = "r" Or _
                e.KeyChar = "s" Or e.KeyChar = "t" Or e.KeyChar = "u" Or e.KeyChar = "U" Or e.KeyChar = "v" Or e.KeyChar = "w" Or e.KeyChar = "x" Or e.KeyChar = "y" Or _
                e.KeyChar = "z" Or e.KeyChar = "I" Or e.KeyChar = "J" Or e.KeyChar = "V" Then
            alphabet_sum = alphabet_sum & e.KeyChar

        Else
            word_temp = alphabet_sum
        End If
    End Sub



    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If e.KeyChar = "A" Or e.KeyChar = "B" Or e.KeyChar = "C" Or e.KeyChar = "D" Or e.KeyChar = "E" Or e.KeyChar = "F" Or e.KeyChar = "G" Or e.KeyChar = "H" Or _
                e.KeyChar = "K" Or e.KeyChar = "L" Or e.KeyChar = "M" Or e.KeyChar = "N" Or e.KeyChar = "O" Or e.KeyChar = "P" Or e.KeyChar = "Q" Or e.KeyChar = "R" Or _
                e.KeyChar = "S" Or e.KeyChar = "T" Or e.KeyChar = "W" Or e.KeyChar = "X" Or e.KeyChar = "Y" Or e.KeyChar = "Z" Or e.KeyChar = "a" Or e.KeyChar = "b" Or _
                e.KeyChar = "c" Or e.KeyChar = "d" Or e.KeyChar = "e" Or e.KeyChar = "f" Or e.KeyChar = "g" Or e.KeyChar = "h" Or e.KeyChar = "i" Or e.KeyChar = "j" Or _
                e.KeyChar = "k" Or e.KeyChar = "l" Or e.KeyChar = "m" Or e.KeyChar = "n" Or e.KeyChar = "o" Or e.KeyChar = "p" Or e.KeyChar = "q" Or e.KeyChar = "r" Or _
                e.KeyChar = "s" Or e.KeyChar = "t" Or e.KeyChar = "u" Or e.KeyChar = "U" Or e.KeyChar = "v" Or e.KeyChar = "w" Or e.KeyChar = "x" Or e.KeyChar = "y" Or _
                e.KeyChar = "z" Or e.KeyChar = "I" Or e.KeyChar = "J" Or e.KeyChar = "V" Then
            alphabet_sum = alphabet_sum & e.KeyChar

        Else
            word_temp = alphabet_sum
        End If
    End Sub

    Private Sub TextBox2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyUp
        On Error Resume Next
        If e.KeyCode = 32 Then
            alphabet_sum = ""
            word_temp = ""
            TextBox2.SelectionColor = Color.Black
        End If
    End Sub

    Private Sub TextBox2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox2.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg2.Visible = True
            pic_msg2.BringToFront()
        End If
    End Sub


    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Dim startingPoint As Integer = -1
        Dim st As Integer
        st = TextBox2.SelectionStart

        Call find_word(word_temp)

        If flag1 = True Then

            If word_temp = "" Then Exit Sub
            Application.DoEvents()
            Do
                startingPoint = TextBox2.Find(word_temp, startingPoint + 1, RichTextBoxFinds.None)
                If (startingPoint >= 0) Then
                    TextBox2.SelectionStart = startingPoint
                    TextBox2.SelectionLength = word_temp.Length
                    TextBox2.SelectionColor = Color.Brown
                End If
            Loop Until startingPoint < 0
            TextBox2.DeselectAll()
            TextBox2.SelectionStart = st
            TextBox2.SelectionColor = Color.Black
            alphabet_sum = ""
            word_temp = ""
            flag1 = False
        End If
    End Sub

    Private Sub TextBox5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyUp
        On Error Resume Next
        If e.KeyCode = 32 Then
            alphabet_sum = ""
            word_temp = ""
            TextBox5.SelectionColor = Color.Black
        End If
    End Sub

    Private Sub TextBox5_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox5.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg4.Visible = True
            pic_msg4.BringToFront()
        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Dim startingPoint As Integer = -1
        Dim st As Integer
        st = TextBox5.SelectionStart

        Call find_word(word_temp)

        If flag1 = True Then

            If word_temp = "" Then Exit Sub
            Application.DoEvents()
            Do
                startingPoint = TextBox5.Find(word_temp, startingPoint + 1, RichTextBoxFinds.None)
                If (startingPoint >= 0) Then
                    TextBox5.SelectionStart = startingPoint
                    TextBox5.SelectionLength = word_temp.Length
                    TextBox5.SelectionColor = Color.Brown
                End If
            Loop Until startingPoint < 0
            TextBox5.DeselectAll()
            TextBox5.SelectionStart = st
            TextBox5.SelectionColor = Color.Black
            alphabet_sum = ""
            word_temp = ""
            flag1 = False
        End If
    End Sub

    Private Sub pic_msg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg.Click
        pic_msg.Visible = False
    End Sub

    Private Sub TextBox3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox3.MouseUp

        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg3.Visible = True
            pic_msg3.BringToFront()
        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub pic_msg4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg4.Click
        pic_msg4.Visible = False
    End Sub

    Private Sub pic_msg3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg3.Click
        pic_msg3.Visible = False
    End Sub

    Private Sub pic_msg2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg2.Click
        pic_msg2.Visible = False
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        TextBox3.SelectedText = Chr(9)
        TextBox3.Focus()
    End Sub
End Class
