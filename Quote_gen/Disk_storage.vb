Imports System.IO
Imports System.Text

Module Disk_storage
    'Retrieve control settings from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Public Sub Read_file_vtk5()

        Dim control_words(), words() As String
        Dim all_num, all_combo, all_check, all_text, all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        Form1.OpenFileDialog1.FileName = "Quote_select_*.vtk5"

        If Directory.Exists(Form1.dirpath_Backup) Then
            Form1.OpenFileDialog1.InitialDirectory = Form1.dirpath_Backup  'used at VTK
        Else
            Form1.OpenFileDialog1.InitialDirectory = Form1.dirpath_Home_GP  'used at home
        End If

        Form1.OpenFileDialog1.Title = "Open a VTK5"
        Form1.OpenFileDialog1.Filter = "VTK5 Files|*.vtk5"
        If Form1.OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            Dim readText As String = File.ReadAllText(Form1.OpenFileDialog1.FileName, Encoding.ASCII)
            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '---------- terugzetten numeric controls (Updated version) -----------------
            FindControlRecursive(all_num, Form1, GetType(NumericUpDown))
            words = control_words(0).Split(separators, StringSplitOptions.None)     'Split the read file content
            Restore_num_controls(words, all_num)

            '---------- terugzetten combobox controls (Updated version) -----------------
            FindControlRecursive(all_combo, Form1, GetType(ComboBox))
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            Restore_combo_controls(words, all_combo)

            '---------- terugzetten checkboxes controls (Updated version) -----------------
            FindControlRecursive(all_check, Form1, GetType(CheckBox))
            words = control_words(2).Split(separators, StringSplitOptions.None)    'Split the read file content
            Restore_checkbox_controls(words, all_check)

            '---------- terugzetten Radio button controls (Updated version) -----------------
            FindControlRecursive(all_radio, Form1, GetType(RadioButton))
            words = control_words(3).Split(separators, StringSplitOptions.None)    'Split the read file content
            Restore_radiobutton_controls(words, all_radio)

            '---------- terugzetten Text controls (Updated version) -----------------
            FindControlRecursive(all_text, Form1, GetType(TextBox))
            words = control_words(4).Split(separators, StringSplitOptions.None)    'Split the read file content
            Restore_text_controls(words, all_text)
        End If
    End Sub

    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    Public Sub Restore_num_controls(words As String(), all_num As List(Of Control))
        Dim ttt As Double

        For i = 0 To all_num.Count - 1
            Dim updown As NumericUpDown = CType(all_num(i), System.Windows.Forms.NumericUpDown)
            '============ find the stored numeric control list ====

            For j = 0 To all_num.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If updown.Name = words(j * 2 + 1) Then    '==== Found ====
                        'Debug.WriteLine("FOUND !! updown.Name= " & updown.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        If Not (Double.TryParse(words(j * 2 + 2), ttt)) Then MessageBox.Show("Numeric controls conversion problem occured")
                        If ttt <= updown.Maximum And ttt >= updown.Minimum Then
                            updown.Value = CDec(ttt)          'OK
                        Else
                            updown.Value = updown.Minimum       'NOK
                            MessageBox.Show("Numeric controls value out of outside min-max range, Minimum value is used")
                        End If
                        Exit For
                    End If
                Else
                    MessageBox.Show(updown.Name & " (num. control) was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub

    Public Sub Restore_combo_controls(words As String(), all_combo As List(Of Control))
        For i = 0 To all_combo.Count - 1
            Dim combobx As ComboBox = CType(all_combo(i), System.Windows.Forms.ComboBox)
            '============ find the stored numeric control list ====

            For j = 0 To all_combo.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If combobx.Name = words(j * 2 + 1) Then    '==== Found ====
                        'Debug.WriteLine("FOUND !! combobx.Name= " & combobx.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        If (i < words.Length - 1) Then
                            combobx.SelectedItem = words(j * 2 + 2)
                        Else
                            MessageBox.Show("Warning last combobox not found in file")
                        End If
                        Exit For
                    End If
                Else
                    MessageBox.Show(combobx.Name & " (combobox) was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub

    Public Sub Restore_checkbox_controls(words As String(), all_check As List(Of Control))
        For i = 0 To all_check.Count - 1
            Dim chbx As CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
            '============ find the stored numeric control list ====

            For j = 0 To all_check.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If chbx.Name = words(j * 2 + 1) Then    '==== Found ====
                        Debug.WriteLine("FOUND !! chbx.Name= " & chbx.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        If CBool(words(j * 2 + 2)) = True Then
                            chbx.Checked = True
                        Else
                            chbx.Checked = False
                        End If

                        Exit For
                    End If
                Else
                    MessageBox.Show(chbx.Name & " (checkbox) was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub

    Public Sub Restore_radiobutton_controls(words As String(), all_radio As List(Of Control))
        For i = 0 To all_radio.Count - 1
            Dim radiobut As RadioButton = CType(all_radio(i), System.Windows.Forms.RadioButton)
            '============ find the stored numeric control list ====
            For j = 0 To all_radio.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If radiobut.Name = words(j * 2 + 1) Then    '==== Found ====
                        'Debug.WriteLine("j= " & j.ToString & ", FOUND !! radiobut.Name= " & radiobt.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        Boolean.TryParse(words(j * 2 + 2), radiobut.Checked)
                        Exit For
                    End If
                Else
                    MessageBox.Show(radiobut.Name & " (radiobutton) was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub

    Public Sub Restore_text_controls(words As String(), all_text As List(Of Control))
        For i = 0 To all_text.Count - 1
            Dim tekst As TextBox = CType(all_text(i), System.Windows.Forms.TextBox)
            '============ find the stored numeric control list ====
            For j = 0 To all_text.Count - 1
                If (j * 2 + 2) < words.Count Then
                    If tekst.Name = words(j * 2 + 1) Then    '==== Found ====
                        Debug.WriteLine("j= " & j.ToString & ", FOUND !! tekst.Name= " & tekst.Name & ", words(j *2)= " & words(j * 2) & ", words(j*2+1)= " & words(j * 2 + 1) & ", words(j*2+2)= " & words(j * 2 + 2))
                        tekst.Text = words(j * 2 + 2)
                        Exit For
                    End If
                Else
                    MessageBox.Show(tekst.Name & " (textbox) was NOT Stored in file and is NOT updated")
                End If
            Next
        Next
    End Sub

End Module
