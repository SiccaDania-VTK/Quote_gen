Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Management
Imports Microsoft.Office.Interop.Word

Public Class Form1
    '----------- directory's-----------
    Dim dirpath_Block As String = "N:\Verkoop\Tekst\Quote_text_block\"
    Dim dirpath_Backup As String = "N:\Verkoop\Aanbiedingen\Quote_gen_backup\"
    Dim dirpath_Home_GP As String = "C:\Temp\"

    Public oWord As Word.Application
    ' see https://support.microsoft.com/en-us/help/316383/how-to-automate-word-from-visual-basic--net-to-create-a-new-document
    Private Sub Generate_word_doc()
        ' Dim oWord As Word.Application
        Dim oDoc As Word.Document
        'Dim oTable As Word.Table
        'Dim oPara1, oPara2, oPara3 As Word.Paragraph
        Dim oPara3 As Word.Paragraph
        Dim ufilename As String
        Dim pathname As String
        Dim style1 As String

        '----------- Select Word style -----------------
        style1 = "N:\VERKOOP\Tekst\Quote_text_block\VTK_Fan_Quote.dotm"

        'Start Word and open the document template. 
        oWord = CType(CreateObject("Word.Application"), Word.Application)
        oWord.Visible = True

        If File.Exists(style1) Then
            oDoc = oWord.Documents.Add(style1.Clone)
        Else
            MessageBox.Show("Dam.. Can not find " & style1)
            oDoc = oWord.Documents.Add
        End If

        '---- find ALL checkboxes controls ---
        '---- sort in Alphabetical order -------
        '---- check for checked ----
        '---- then PRINT 
        TextBox05.Clear()
        Dim all_check As New List(Of Control)
        FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the controls
        all_check = all_check.OrderBy(Function(x) x.Text).ToList()  'Alphabetical order

        For i = 0 To all_check.Count - 1
            Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
            If grbx.Checked = True Then
                oPara3 = oDoc.Content.Paragraphs.Add()
                pathname = dirpath_Block & grbx.Text.Substring(0, 4) & ".docx"
                If File.Exists(pathname) Then
                    TextBox05.Text &= "OK, file found " & pathname & vbCrLf
                    oPara3.Range.InsertFile(pathname)
                Else
                    TextBox05.Text &= "File not found " & pathname & vbCrLf
                End If
            End If
        Next

        '============ search and replace in WORD file================
        'Dim myStoryRange As Range '= oWord.ActiveDocument.Content

        Dim find_s As String = ""
        Dim rep_s As String = ""

        Find_rep(Label1.Text, TextBox01.Text)
        Find_rep(Label3.Text, TextBox07.Text)
        Find_rep(Label4.Text, TextBox08.Text)
        Find_rep(Label5.Text, TextBox09.Text)
        Find_rep(Label6.Text, TextBox02.Text)    'Cust name

        Find_rep(Label7.Text, TextBox11.Text)
        Find_rep(Label8.Text, TextBox12.Text)
        Find_rep(Label9.Text, TextBox13.Text)
        Find_rep(Label11.Text, TextBox14.Text)
        Find_rep(Label12.Text, TextBox15.Text)

        Find_rep(Label24.Text, TextBox24.Text)
        Find_rep(Label25.Text, TextBox25.Text)
        Find_rep(Label26.Text, TextBox26.Text)
        Find_rep(Label27.Text, TextBox27.Text)

        Find_rep(Label21.Text, ComboBox1.Text)
        Find_rep(Label22.Text, ComboBox2.Text)
        Find_rep(Label23.Text, ComboBox3.Text)

        '==================== backup final product===============
        ufilename = "Quote_" & TextBox01.Text & "_" & TextBox02.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".docx"

        If Directory.Exists(dirpath_Backup) Then
            ufilename = dirpath_Backup & ufilename
        Else
            ufilename = dirpath_Home_GP & ufilename
        End If
        'oWord.ActiveDocument.SaveAs(ufilename.ToString)
    End Sub

    Private Sub Find_rep(find_s As String, rep_s As String)
        '============ search and replace in WORD file================
        Dim myStoryRange As Range

        For Each myStoryRange In oWord.ActiveDocument.StoryRanges
            With myStoryRange.Find
                .Text = find_s.ToString
                .Replacement.Text = rep_s.ToString
                .Wrap = WdFindWrap.wdFindContinue
                .Execute(Replace:=Word.WdReplace.wdReplaceAll)
            End With
            Do While Not (myStoryRange.NextStoryRange Is Nothing)
                myStoryRange = myStoryRange.NextStoryRange
                With myStoryRange.Find
                    .Text = find_s.ToString
                    .Replacement.Text = rep_s.ToString
                    .Wrap = WdFindWrap.wdFindContinue
                    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                End With
            Loop
        Next myStoryRange
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Check_directories()
        Generate_word_doc()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox03.Text =
        "File naming convention" & vbCrLf & vbCrLf &
        "Text block location is " & vbTab & dirpath_Block.ToString & vbCrLf &
        "Quote backup location is " & vbTab & dirpath_Backup.ToString & vbCrLf &
        "  " & vbCrLf &
        "File-name ate the first 4 character of the checkbox name" & vbCrLf &
        "Printing squence is determined by the file_name sorted in alphabetical order" & vbCrLf &
        " "
        TextBox06.Text =
        "Quotes use font Khmer UI size 10" & vbCrLf &
        "New quotes use the local normal.dot with location" & vbCrLf &
        "C:\\users\(your user name)\appdata\roaming\microsoft\templates.." & vbCrLf
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox01.Text.Trim.Length > 0 And TextBox02.Text.Trim.Length > 0 Then
            Save_tofile()
        Else
            MessageBox.Show("Complete Quote and Customer name")
        End If
    End Sub
    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()

        Dim temp_string, user As String

        user = Trim(Environment.UserName)         'User name on the screen
        Dim filename As String = "Quote_select_" & TextBox01.Text & "_" & TextBox02.Text & DateTime.Now.ToString("_yyyy_MM_dd_") & user & ".vtkq"
        Dim all_num, all_combo, all_check, all_text As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox02.Text) Then
            TextBox02.Text = "name"
        End If

        temp_string = TextBox01.Text & ";" & TextBox02.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
            temp_string &= grbx.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save -------
        FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all textbox controls and save ----------
        FindControlRecursive(all_text, Me, GetType(System.Windows.Forms.TextBox))      'Find the control
        all_text = all_text.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_text.Count - 1
            Dim grbx As System.Windows.Forms.TextBox = CType(all_text(i), System.Windows.Forms.TextBox)
            temp_string &= grbx.Text.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        Try
            Check_directories()  'Are the directories present
            If CInt(temp_string.Length.ToString) > 5 Then      'String may be empty
                If Directory.Exists(dirpath_Backup) Then
                    File.WriteAllText(dirpath_Backup & filename, temp_string, Encoding.ASCII)      'used at VTK
                Else
                    File.WriteAllText(dirpath_Home_GP & filename, temp_string, Encoding.ASCII)     'used at home
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Line 5062, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Check_directories()
        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home_GP)) Then System.IO.Directory.CreateDirectory(dirpath_Home_GP)
            If (Not System.IO.Directory.Exists(dirpath_Block)) Then System.IO.Directory.CreateDirectory(dirpath_Block)
            If (Not System.IO.Directory.Exists(dirpath_Backup)) Then System.IO.Directory.CreateDirectory(dirpath_Backup)
        Catch ex As Exception
        End Try
    End Sub

    'Retrieve control settings and case_x_conditions from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file()
        Dim control_words(), words() As String
        Dim i As Integer
        Dim k As Integer = 0
        Dim all_num, all_combo, all_check, all_text As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "Quote_select_*"

        If Directory.Exists(dirpath_Backup) Then
            OpenFileDialog1.InitialDirectory = dirpath_Backup  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_Home_GP  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTKQ Files|*.vtkq|VTKQ file|*.vtkq"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)

            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '----- project data -----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split the read file content
            TextBox01.Text = words(0)                  'Project number
            TextBox02.Text = words(1)                  'Item no

            '---------- terugzetten combobox controls -----------------
            FindControlRecursive(all_combo, Me, GetType(ComboBox))
            all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()          'Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_combo.Count - 1
                Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.SelectedItem = words(i + 1)
                Else
                    MessageBox.Show("Warning last combobox not found in file")
                End If
            Next

            '---------- terugzetten checkbox controls -----------------
            FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(2).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last checkbox not found in file")
                End If
            Next

            '---------- terugzetten textbox controls -----------------
            FindControlRecursive(all_text, Me, GetType(System.Windows.Forms.TextBox))      'Find the control
            all_text = all_text.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(3).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_text.Count - 1
                Dim grbx As System.Windows.Forms.TextBox = CType(all_text(i), System.Windows.Forms.TextBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.Text = words(i + 1)
                Else
                    MessageBox.Show("Warning last textbox not found in file")
                End If
            Next
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Read_file()
    End Sub

    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, Button4.Enter, TabPage10.Enter
        Dim i As Integer
        Dim k As Integer = 0
        Dim all_check As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        TextBox04.Clear()

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Text).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
            If grbx.Checked = True Then
                TextBox04.Text &= grbx.Text & vbCrLf
            End If
        Next
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, CheckBox179.CheckStateChanged, CheckBox178.CheckStateChanged, CheckBox176.CheckStateChanged, CheckBox175.CheckStateChanged, CheckBox174.CheckStateChanged, CheckBox173.CheckStateChanged, CheckBox26.CheckStateChanged, CheckBox25.CheckStateChanged, CheckBox24.CheckStateChanged, CheckBox23.CheckStateChanged, CheckBox22.CheckStateChanged, CheckBox118.CheckStateChanged, CheckBox201.CheckedChanged, CheckBox200.CheckedChanged, CheckBox199.CheckedChanged, CheckBox198.CheckedChanged, CheckBox197.CheckedChanged, CheckBox196.CheckedChanged, CheckBox195.CheckedChanged, CheckBox194.CheckedChanged, CheckBox88.CheckStateChanged, CheckBox144.CheckStateChanged, CheckBox143.CheckStateChanged, CheckBox141.CheckStateChanged, CheckBox156.CheckedChanged, CheckBox155.CheckedChanged, CheckBox39.CheckedChanged, CheckBox38.CheckedChanged, CheckBox31.CheckedChanged, CheckBox30.CheckedChanged, CheckBox29.CheckedChanged, CheckBox162.CheckedChanged, CheckBox161.CheckedChanged, CheckBox167.CheckedChanged, CheckBox160.CheckedChanged, CheckBox213.CheckedChanged, CheckBox183.CheckedChanged, CheckBox11.CheckedChanged
        Check_combinations()
    End Sub
    Private Sub Check_combinations()
        Dim no_checked As Integer = 0
        Dim kctrl As System.Windows.Forms.Control

        '---------- Groupbox 3 (Bearings)---------------
        no_checked = 0
        For Each kctrl In GroupBox3.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox3.BackColor = CType(IIf(no_checked > 1 Or no_checked = 0, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 4 (Electrical)---------------
        no_checked = 0
        For Each kctrl In GroupBox4.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox4.BackColor = CType(IIf(no_checked > 1, Color.Red, SystemColors.Window), Color)


        '---------- Groupbox 19 (Casing)---------------
        no_checked = 0
        For Each kctrl In GroupBox19.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox19.BackColor = CType(IIf(no_checked > 1 Or no_checked = 0, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 21 (Vane type)---------------
        no_checked = 0
        For Each kctrl In GroupBox21.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox21.BackColor = CType(IIf(no_checked > 1 Or no_checked = 0, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 29 (Coupling)---------------
        no_checked = 0
        For Each kctrl In GroupBox29.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox29.BackColor = CType(IIf(no_checked > 1 Or no_checked = 0, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 30 (Coupling guard)---------------
        no_checked = 0
        For Each kctrl In GroupBox30.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.checked Then no_checked += 1
            End If
        Next
        GroupBox30.BackColor = CType(IIf(no_checked > 1 Or no_checked = 0, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 32 (Motor)---------------
        no_checked = 0
        For Each kctrl In GroupBox32.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox32.BackColor = CType(IIf(no_checked > 1 Or no_checked = 0, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 34 (VSD)---------------
        no_checked = 0
        For Each kctrl In GroupBox34.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox34.BackColor = CType(IIf(no_checked > 1, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 40 (Motor efficiency)---------------
        no_checked = 0
        For Each kctrl In GroupBox40.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox40.BackColor = CType(IIf(no_checked > 1, Color.Red, SystemColors.Window), Color)

        '---------- Groupbox 42 (Vib sensor)---------------
        no_checked = 0
        For Each kctrl In GroupBox42.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox42.BackColor = CType(IIf(no_checked > 1, Color.Red, SystemColors.Window), Color)


        '---------- Groupbox 44 (Electrical)---------------
        no_checked = 0
        For Each kctrl In GroupBox44.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        GroupBox44.BackColor = CType(IIf(no_checked > 1, Color.Red, SystemColors.Window), Color)

    End Sub
End Class
