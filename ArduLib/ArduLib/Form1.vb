Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO

Public Class Form1
    '  Asemani Personal Laboratory
    ' Post.asemani@gmail.com
    Dim alphabet_sum As String
    Dim word_temp As String
    Dim flag1 As String


    Private Sub btn_new_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_new.Click
        On Error Resume Next
        Dim msg As String
        msg = MsgBox("are you sure?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "new library?")
        If msg = vbYes Then
            txt_header.Text = txt_header_temp.Text
            txt_source.Text = txt_source_temp.Text
            txt_keywords.Text = txt_key_temp.Text
            txt_example.Text = txt_exam_temp.Text
            txt_property.Text = property_temp.Text
            txt_custom.Text = "blank page file...."
            txt_exam_name.Text = "NewExample"
            txt_new_name.Text = "New"
            txtSubject.Text = "New"
            txtTitle.Text = "New"

            Label3.Text = "New"
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        On Error Resume Next

        SaveFileDialog1.Title = "Save header file to:"
        SaveFileDialog1.FileName = txt_new_name.Text
        SaveFileDialog1.Filter = "Arduino Library Header file (*.h)|*.h|All files (*.*)|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            lbl_header_file.Text = SaveFileDialog1.FileName
            If lbl_header_file.Text = "" Then Exit Sub

            System.IO.File.WriteAllText(lbl_header_file.Text, txt_header.Text) ', System.Text.Encoding.UTF8)
            MessageBox.Show("Header file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        On Error Resume Next

        SaveFileDialog1.Title = "Save source file to:"
        SaveFileDialog1.FileName = txt_new_name.Text
        SaveFileDialog1.Filter = "Arduino Library Source file (*.cpp)|*.cpp|All files (*.*)|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            lbl_source_file.Text = SaveFileDialog1.FileName
            If lbl_source_file.Text = "" Then Exit Sub
            System.IO.File.WriteAllText(lbl_source_file.Text, txt_source.Text)
            MessageBox.Show("Source file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        On Error Resume Next

        SaveFileDialog1.Title = "Save keyword file to:"
        SaveFileDialog1.FileName = "keywords"
        SaveFileDialog1.Filter = "Arduino Keywords Source file (*.txt)|*.txt|All files (*.*)|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            lbl_keywoeds_file.Text = SaveFileDialog1.FileName
            If lbl_keywoeds_file.Text = "" Then Exit Sub

            System.IO.File.WriteAllText(lbl_keywoeds_file.Text, txt_keywords.Text)
            MessageBox.Show("Keywords file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        On Error Resume Next
        SaveFileDialog1.Title = "Save example file to:"
        SaveFileDialog1.FileName = txt_exam_name.Text
        SaveFileDialog1.Filter = "Arduino Source Code file (*.ino)|*.ino|All files (*.*)|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            lbl_example_file.Text = SaveFileDialog1.FileName
            If lbl_example_file.Text = "" Then Exit Sub

            System.IO.File.WriteAllText(lbl_example_file.Text, txt_example.Text)

            MessageBox.Show("Example file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub



    Private Sub make_library()
        On Error Resume Next

        Dim dirpath As String

        FolderBrowserDialog1.ShowNewFolderButton = True
        FolderBrowserDialog1.ShowDialog()
        dirpath = FolderBrowserDialog1.SelectedPath
        If System.IO.Directory.Exists(dirpath + "\" + txt_new_name.Text) = True Then
            Dim t As String
            t = MsgBox(dirpath + "\" + txt_new_name.Text + vbCrLf + "do you want to replace it?", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "same detected!")
            If t = vbOK Then
                System.IO.Directory.Delete(dirpath + "\" + txt_new_name.Text, True)
            Else
                Exit Sub
            End If
        End If
        If dirpath <> "" Then
            MkDir(dirpath + "\" + txt_new_name.Text)
            MkDir(dirpath + "\" + txt_new_name.Text + "\Documentation")
            MkDir(dirpath + "\" + txt_new_name.Text + "\utility")
            'header
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\" + txt_new_name.Text + ".h", txt_header.Text)

            'source
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\" + txt_new_name.Text + ".cpp", txt_source.Text)
            'keywords
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\" + "keywords.txt", txt_keywords.Text)
            'property
            'System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\library.properties", txt_property.Text, System.Text.Encoding.UTF8)
            txt_property.SaveFile(dirpath + "\" + txt_new_name.Text + "\library.properties", RichTextBoxStreamType.TextTextOleObjs)
            'example

            MkDir(dirpath + "\" + txt_new_name.Text + "\examples")
            MkDir(dirpath + "\" + txt_new_name.Text + "\examples\" + txt_exam_name.Text)
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\examples\" + txt_exam_name.Text + "\" + txt_exam_name.Text + ".ino", txt_example.Text)

            MessageBox.Show("Library made successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub fixer(ByVal pat As String)
        On Error Resume Next
        Shell(Application.StartupPath + "\fixer.exe " + pat, vbNormalFocus)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next
        e.Cancel = True
        Dim msg As String
        msg = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit")
        If msg = vbYes Then
            End
        Else
            Me.Visible = True
        End If
    End Sub



    Private Sub Form1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        On Error Resume Next
        If e.KeyCode = Keys.F5 Then
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        Call load_inner_ketwords()
        Call full_color_parser(txt_header, Color.Brown)
        Call full_color_parser(txt_source, Color.Brown)
        Call full_color_parser(txt_example, Color.DarkCyan)


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

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            header_path = OpenFileDialog1.FileName
            txt_new_name.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")

            txt_header.Text = System.IO.File.ReadAllText(header_path)
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        On Error Resume Next
        Dim source_path As String

        OpenFileDialog1.Title = "Select a Source file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library Source files (*.cpp)|*.cpp|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            source_path = OpenFileDialog1.FileName


            txt_source.Text = System.IO.File.ReadAllText(source_path)
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub

    Private Sub btn_new_name_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_new_name.Click
        On Error Resume Next


        txt_header.Text = Replace(txt_header.Text, Label3.Text, txt_new_name.Text)
        txt_source.Text = Replace(txt_source.Text, Label3.Text, txt_new_name.Text)
        txt_keywords.Text = Replace(txt_keywords.Text, Label3.Text, txt_new_name.Text)
        txt_example.Text = Replace(txt_example.Text, Label3.Text, txt_new_name.Text)
        txt_property.Text = Replace(txt_property.Text, Label3.Text, txt_new_name.Text)
        txtSubject.Text = Replace(txtSubject.Text, Label3.Text, txt_new_name.Text)
        txtTitle.Text = Replace(txtTitle.Text, Label3.Text, txt_new_name.Text)
        txt_exam_name.Text = "Example"
        Label3.Text = txt_new_name.Text
        Call full_color_parser(txt_header, Color.Brown)
        Call full_color_parser(txt_source, Color.Brown)
        Call full_color_parser(txt_example, Color.DarkCyan)

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        On Error Resume Next
        Dim exam_path As String

        OpenFileDialog1.Title = "Select a Example file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library Example files (*.ino)|*.ino|PDE files (*.pde)|*.pde|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            exam_path = OpenFileDialog1.FileName


            txt_example.Text = System.IO.File.ReadAllText(exam_path)
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub


    Private Sub Label5_Click_4(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        On Error Resume Next
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
        On Error Resume Next
        Call full_color_parser(txt_source, Color.Brown)
    End Sub

    Private Sub Button13_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        On Error Resume Next
        Call full_color_parser(txt_header, Color.Brown)
    End Sub


    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        On Error Resume Next
        Call full_color_parser(txt_example, Color.DarkCyan)
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


    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_header.KeyPress

        On Error Resume Next


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
        On Error Resume Next
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

    Private Sub TextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_header.KeyUp
        On Error Resume Next
        If e.KeyCode = 32 Then
            alphabet_sum = ""
            word_temp = ""
            txt_header.SelectionColor = Color.Black
        End If
        If e.KeyCode = Keys.Tab Then
            txt_header.SelectedText = Chr(9)
        End If
        If e.KeyCode = Keys.F5 Then
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub

    Private Sub TextBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txt_header.MouseUp
        On Error Resume Next
        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg.Visible = True

        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_header.TextChanged
        On Error Resume Next
        Dim startingPoint As Integer = -1
        Dim st As Integer
        st = txt_header.SelectionStart

        Call find_word(word_temp)

        If flag1 = True Then

            If word_temp = "" Then Exit Sub
            Application.DoEvents()
            Do
                startingPoint = txt_header.Find(word_temp, startingPoint + 1, RichTextBoxFinds.None)
                If (startingPoint >= 0) Then
                    txt_header.SelectionStart = startingPoint
                    txt_header.SelectionLength = word_temp.Length
                    txt_header.SelectionColor = Color.Brown
                End If
            Loop Until startingPoint < 0
            txt_header.DeselectAll()
            txt_header.SelectionStart = st
            txt_header.SelectionColor = Color.Black
            alphabet_sum = ""
            word_temp = ""
            flag1 = False
        End If

    End Sub

    Private Sub TextBox3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_keywords.KeyUp
        On Error Resume Next
        If e.KeyCode = Keys.Tab Then
            txt_keywords.SelectedText = Chr(9)
        End If

    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_source.KeyPress
        On Error Resume Next
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



    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt_example.KeyPress
        On Error Resume Next
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

    Private Sub TextBox2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_source.KeyUp
        On Error Resume Next
        If e.KeyCode = 32 Then
            alphabet_sum = ""
            word_temp = ""
            txt_source.SelectionColor = Color.Black
        End If
        If e.KeyCode = Keys.Tab Then
            txt_source.SelectedText = Chr(9)
        End If
        If e.KeyCode = Keys.F5 Then
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub

    Private Sub TextBox2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txt_source.MouseUp
        On Error Resume Next
        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg2.Visible = True
            pic_msg2.BringToFront()
        End If
    End Sub


    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_source.TextChanged
        On Error Resume Next
        Dim startingPoint As Integer = -1
        Dim st As Integer
        st = txt_source.SelectionStart

        Call find_word(word_temp)

        If flag1 = True Then

            If word_temp = "" Then Exit Sub
            Application.DoEvents()
            Do
                startingPoint = txt_source.Find(word_temp, startingPoint + 1, RichTextBoxFinds.None)
                If (startingPoint >= 0) Then
                    txt_source.SelectionStart = startingPoint
                    txt_source.SelectionLength = word_temp.Length
                    txt_source.SelectionColor = Color.Brown
                End If
            Loop Until startingPoint < 0
            txt_source.DeselectAll()
            txt_source.SelectionStart = st
            txt_source.SelectionColor = Color.Black
            alphabet_sum = ""
            word_temp = ""
            flag1 = False
        End If
    End Sub

    Private Sub TextBox5_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_example.KeyUp
        On Error Resume Next
        If e.KeyCode = 32 Then
            alphabet_sum = ""
            word_temp = ""
            txt_example.SelectionColor = Color.Black
        End If
        If e.KeyCode = Keys.Tab Then
            txt_example.SelectedText = Chr(9)
        End If
        If e.KeyCode = Keys.F5 Then
            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub

    Private Sub TextBox5_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txt_example.MouseUp
        On Error Resume Next
        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg4.Visible = True
            pic_msg4.BringToFront()
        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_example.TextChanged
        On Error Resume Next
        Dim startingPoint As Integer = -1
        Dim st As Integer
        st = txt_example.SelectionStart

        Call find_word(word_temp)

        If flag1 = True Then

            If word_temp = "" Then Exit Sub
            Application.DoEvents()
            Do
                startingPoint = txt_example.Find(word_temp, startingPoint + 1, RichTextBoxFinds.None)
                If (startingPoint >= 0) Then
                    txt_example.SelectionStart = startingPoint
                    txt_example.SelectionLength = word_temp.Length
                    txt_example.SelectionColor = Color.DarkCyan
                End If
            Loop Until startingPoint < 0
            txt_example.DeselectAll()
            txt_example.SelectionStart = st
            txt_example.SelectionColor = Color.Black
            alphabet_sum = ""
            word_temp = ""
            flag1 = False
        End If
    End Sub

    Private Sub pic_msg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg.Click
        On Error Resume Next
        pic_msg.Visible = False
    End Sub

    Private Sub TextBox3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txt_keywords.MouseUp
        On Error Resume Next

        If e.Button = Windows.Forms.MouseButtons.Right Then
            pic_msg3.Visible = True
            pic_msg3.BringToFront()
        End If
    End Sub



    Private Sub pic_msg4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg4.Click
        On Error Resume Next
        pic_msg4.Visible = False
    End Sub

    Private Sub pic_msg3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg3.Click
        On Error Resume Next
        pic_msg3.Visible = False
    End Sub

    Private Sub pic_msg2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_msg2.Click
        On Error Resume Next
        pic_msg2.Visible = False
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        On Error Resume Next
        txt_keywords.SelectedText = Chr(9)
        txt_keywords.Focus()
    End Sub

    Private Sub Button13_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button13.MouseLeave
        On Error Resume Next
        Label6.Text = ""
    End Sub

    Private Sub Button13_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button13.MouseMove
        On Error Resume Next
        Label6.Text = "Parsing the code and indicate the keywords"
    End Sub

    Private Sub Button8_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.MouseLeave
        On Error Resume Next
        Label6.Text = ""
    End Sub

    Private Sub Button8_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button8.MouseMove
        On Error Resume Next
        Label6.Text = "import a page file and overwrite on the current page..."
    End Sub

    Private Sub Button3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.MouseLeave
        On Error Resume Next
        Label6.Text = ""
    End Sub

    Private Sub Button3_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseMove
        On Error Resume Next
        Label6.Text = "export the current page file..."
    End Sub

    Private Sub btn_publish_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btn_publish.MouseDown

    End Sub

    Private Sub btn_publish_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_publish.MouseLeave
        On Error Resume Next
        Label6.Text = ""
    End Sub

    Private Sub btn_publish_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btn_publish.MouseMove
        On Error Resume Next

        Label6.Text = "export current library completely..."
    End Sub

    Private Sub btn_new_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_new.MouseLeave
        On Error Resume Next
        Label6.Text = ""
    End Sub

    Private Sub btn_new_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btn_new.MouseMove
        On Error Resume Next
        Label6.Text = "clean the pages for new library"
    End Sub

    Private Sub btn_new_name_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_new_name.MouseLeave
        On Error Resume Next
        Label6.Text = ""
    End Sub

    Private Sub btn_new_name_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btn_new_name.MouseMove
        On Error Resume Next
        Label6.Text = "write the new library name to project"
    End Sub


    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        On Error Resume Next
        MsgBox(property_msg.Text, vbInformation, "library.properties")
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        On Error Resume Next
        SaveFileDialog1.Title = "Save custom file to:"
        SaveFileDialog1.FileName = "File"
        SaveFileDialog1.Filter = "Arduino sketch file (*.ino)|*.ino|Arduino Header file (*.h)|*.h|Arduino Source file (*.cpp)|*.cpp|Arduino Text file (*.txt)|*.txt|All files (*.*)|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            lbl_custom.Text = SaveFileDialog1.FileName
            If lbl_custom.Text = "" Then Exit Sub


                System.IO.File.WriteAllText(lbl_custom.Text, txt_custom.Text)

                MessageBox.Show("File saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        On Error Resume Next
        Dim prop_path As String

        OpenFileDialog1.Title = "Select a properties file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library UTF8 format Properties files (*.properties)|*.properties|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            prop_path = OpenFileDialog1.FileName


            txt_property.Text = System.IO.File.ReadAllText(prop_path, System.Text.Encoding.UTF8)
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        On Error Resume Next
        SaveFileDialog1.Title = "Save properties file to:"
        SaveFileDialog1.FileName = "library"
        SaveFileDialog1.Filter = "Arduino UTF8 format Properties file (*.properties)|*.properties|All files (*.*)|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            lbl_prop.Text = SaveFileDialog1.FileName
            If lbl_prop.Text = "" Then Exit Sub

            'System.IO.File.WriteAllText(lbl_prop.Text, txt_property.Text, System.Text.Encoding.UTF8)
            txt_property.SaveFile(lbl_prop.Text, RichTextBoxStreamType.TextTextOleObjs)
            MessageBox.Show("Properties file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub



    Private Sub make_new_library()
        On Error Resume Next

        Dim dirpath As String

        FolderBrowserDialog1.ShowNewFolderButton = True
        FolderBrowserDialog1.ShowDialog()
        dirpath = FolderBrowserDialog1.SelectedPath
        If System.IO.Directory.Exists(dirpath + "\" + txt_new_name.Text) = True Then
            Dim t As String
            t = MsgBox(dirpath + "\" + txt_new_name.Text + vbCrLf + "do you want to replace it?", MsgBoxStyle.Question + MsgBoxStyle.OkCancel, "same detected!")
            If t = vbOK Then
                System.IO.Directory.Delete(dirpath + "\" + txt_new_name.Text, True)
            Else
                Exit Sub
            End If
        End If
        If dirpath <> "" Then
            MkDir(dirpath + "\" + txt_new_name.Text)
            MkDir(dirpath + "\" + txt_new_name.Text + "\extras")
            MkDir(dirpath + "\" + txt_new_name.Text + "\src")
            'header
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\src\" + "\" + txt_new_name.Text + ".h", txt_header.Text)

            'source
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\src\" + "\" + txt_new_name.Text + ".cpp", txt_source.Text)
            'keywords
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\keywords.txt", txt_keywords.Text)
            'property
            'System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\library.properties", txt_property.Text, System.Text.Encoding.UTF8)

            txt_property.SaveFile(dirpath + "\" + txt_new_name.Text + "\library.properties", RichTextBoxStreamType.TextTextOleObjs)
            'example

            MkDir(dirpath + "\" + txt_new_name.Text + "\examples")
            MkDir(dirpath + "\" + txt_new_name.Text + "\examples\" + txt_exam_name.Text)
            System.IO.File.WriteAllText(dirpath + "\" + txt_new_name.Text + "\examples\" + txt_exam_name.Text + "\" + txt_exam_name.Text + ".ino", txt_example.Text)

            MessageBox.Show("Library made successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

   

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        On Error Resume Next
        Dim keyword_path As String

        OpenFileDialog1.Title = "Select a KeyWords file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library KeyWords files (*.txt)|*.txt|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            keyword_path = OpenFileDialog1.FileName

            txt_keywords.Text = System.IO.File.ReadAllText(keyword_path)

            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)
            Call full_color_parser(txt_example, Color.DarkCyan)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        On Error Resume Next
        Dim prop_path As String

        OpenFileDialog1.Title = "Select a file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino sketch file (*.ino)|*.ino|Arduino Header file (*.h)|*.h|Arduino Source file (*.cpp)|*.cpp|Arduino Text file (*.txt)|*.txt|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            prop_path = OpenFileDialog1.FileName


            txt_custom.Text = System.IO.File.ReadAllText(prop_path)
        End If
    End Sub


    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        On Error Resume Next
        Dim pdfDoc As New Document()
        Dim pdf_path As String


        SaveFileDialog1.Title = "Save Document file to:"
        SaveFileDialog1.FileName = "Document"
        SaveFileDialog1.Filter = "Arduino PDF Document file (*.pdf)|*.pdf|All files (*.*)|*.*"

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            pdf_path = SaveFileDialog1.FileName
            If pdf_path = "" Then Exit Sub

            Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(pdfDoc, New FileStream(pdf_path, FileMode.Create))
            pdfDoc.Open()
            pdfDoc.AddAuthor(txt_author.Text)
            pdfDoc.AddCreator("APL Arduino Library Builder")
            pdfDoc.AddSubject(txtSubject.Text)
            pdfDoc.AddTitle(txtTitle.Text)
            pdfDoc.Add(New Paragraph(txtText.Text))

            pdfDoc.Close()

            MessageBox.Show("Document file saved successfuly!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

   

    Private Sub txt_property_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_property.KeyUp
        On Error Resume Next
        If e.KeyCode = Keys.Tab Then
            txt_property.SelectedText = Chr(9)
        End If
    End Sub

    Private Sub txt_custom_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_custom.KeyUp
        On Error Resume Next
        If e.KeyCode = Keys.Tab Then
            txt_custom.SelectedText = Chr(9)
        End If
    End Sub

    Private Sub txtText_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtText.KeyUp
        On Error Resume Next
        If e.KeyCode = Keys.Tab Then
            txtText.SelectedText = Chr(9)
        End If
    End Sub


    
    Private Sub MakeOldVersionLibraryToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MakeOldVersionLibraryToolStripMenuItem1.Click
        On Error Resume Next
        Call make_library()
    End Sub

    Private Sub MakeNewVersionLibraryToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MakeNewVersionLibraryToolStripMenuItem1.Click
        On Error Resume Next
        Call make_new_library()
    End Sub

   
    Private Sub btn_publish_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btn_publish.MouseUp
        On Error Resume Next
        Me.popup_menu1.Show(Me.Location.X + btn_publish.Location.X + e.X, Me.Location.Y + btn_publish.Location.Y + e.Y)
    End Sub



    Private Sub btn_open_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btn_open.MouseUp
        On Error Resume Next
        Me.popup_menu2.Show(Me.Location.X + btn_publish.Location.X + e.X, Me.Location.Y + btn_publish.Location.Y + e.Y)
    End Sub


    Private Sub open_new_lib()
        On Error Resume Next
        Dim header_path As String
        Dim main_path As String
        Dim new_path As String

        OpenFileDialog1.Title = "Select a Header file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library Header files (*.h)|*.h|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            header_path = OpenFileDialog1.FileName
            txt_new_name.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")
            txtSubject.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")
            txtTitle.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")
            main_path = Mid(header_path, 0, header_path.Length - OpenFileDialog1.SafeFileName.Length)


            txt_header.Text = System.IO.File.ReadAllText(header_path) 'load header
            new_path = main_path + Replace(OpenFileDialog1.SafeFileName, ".h", ".cpp") 'load source
            txt_source.Text = "Load manually by click on the [ Import ] button below"
            txt_source.Text = System.IO.File.ReadAllText(new_path)

            

            new_path = Mid(header_path, 1, header_path.Length - OpenFileDialog1.SafeFileName.Length - 4) + "library.properties"

            txt_property.Text = "Load manually by click on the [ Import ] button below"
            txt_property.Text = System.IO.File.ReadAllText(new_path, System.Text.Encoding.UTF8)


            new_path = Mid(header_path, 1, header_path.Length - OpenFileDialog1.SafeFileName.Length - 4) + "keywords.txt"
            txt_keywords.Text = "Load manually by click on the [ Import ] button below"
            txt_keywords.Text = System.IO.File.ReadAllText(new_path)


            

            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)

            txt_example.Text = "Load manually by click on the [ Import ] button below"

        End If
    End Sub

    Private Sub open_old_lib()
        On Error Resume Next
        Dim header_path As String
        Dim main_path As String
        Dim new_path As String

        OpenFileDialog1.Title = "Select a Header file:"
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "Arduino Library Header files (*.h)|*.h|All files (*.*)|*.*"

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            header_path = OpenFileDialog1.FileName
            txt_new_name.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")
            txtSubject.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")
            txtTitle.Text = Replace(OpenFileDialog1.SafeFileName, ".h", "")
            main_path = Mid(header_path, 0, header_path.Length - OpenFileDialog1.SafeFileName.Length)


            txt_header.Text = System.IO.File.ReadAllText(header_path)

            new_path = main_path + "library.properties"
            txt_property.Text = "Load manually by click on the [ Import ] button below"
            txt_property.Text = System.IO.File.ReadAllText(new_path, System.Text.Encoding.UTF8)

            new_path = main_path + "keywords.txt"
            txt_keywords.Text = "Load manually by click on the [ Import ] button below"
            txt_keywords.Text = System.IO.File.ReadAllText(new_path)

            new_path = main_path + Replace(OpenFileDialog1.SafeFileName, ".h", ".cpp")
            txt_source.Text = "Load manually by click on the [ Import ] button below"
            txt_source.Text = System.IO.File.ReadAllText(new_path)

            Call full_color_parser(txt_header, Color.Brown)
            Call full_color_parser(txt_source, Color.Brown)

            txt_example.Text = "Load manually by click on the [ Import ] button below"

        End If
    End Sub

    Private Sub OpenOldVersionLibraryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenOldVersionLibraryToolStripMenuItem.Click
        On Error Resume Next
        Call open_old_lib()
    End Sub

    Private Sub OpenNewVersionLibraryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenNewVersionLibraryToolStripMenuItem.Click
        On Error Resume Next
        Call open_new_lib()
    End Sub

    Private Sub btn_publish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_publish.Click
        On Error Resume Next
        TabControl1.SelectTab(4)
    End Sub

   
    Private Sub btn_open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_open.Click

    End Sub
End Class
