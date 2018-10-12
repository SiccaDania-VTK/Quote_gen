﻿Imports System.IO
Imports System.Text
Imports System.Math
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Management
'Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class Form1
    'Keep the application object and the workbook object global, so you can  
    'retrieve the data in Button2_Click that was set in Button6_Click.
    Dim objApp As Excel.Application
    Dim objBook As Excel._Workbook

    '----------- directory's-----------
    Dim dirpath_Block As String = "N:\Verkoop\Tekst\Quote_text_block\"
    Dim dirpath_Backup As String = "N:\Verkoop\Aanbiedingen\Quote_gen_backup\"
    Dim dirpath_Home_GP As String = "C:\Temp\"

    Public Shared Flight_dia() As String =   'tbv screw diameter selectie
      {"280", "330", "400", "500", "630", "800", "1000", "1200", "1400"}

    Public flight_pitch() As String = {"variable", "1/2x Diam.", "3/4x Diam.", "1x Diam."}
    Public atex_zone() As String = {"0", "1", "2", "20", "21", "22"}
    Public atex_group() As String = {"IIA", "IIB", "IIC"}
    Public atex_temp() As String = {"T1", "T2", "T3", "T4", "T5", "T6"}
    Public drive_make() As String = {"SEW", "Nord", "Bauer", "Flender"}


    Public Shared steel() As String =
  {"16M03;                  EN10028-2 UNS;          16M03;                          1.5415;     Plate",
   "Aluminium D54S;         DIN1745-1;              AA5083 AIMo45Mn-H116;           3.3547;     Max 70c",
   "Corten A / B;           EN10155 UNS;            S355J2G1W;                      1.8962/63;  Plate",
   "Duplex (Avesta-2205);   EN 10088-1 UfllW;       X2CrNiMoN22-5-3 saisna;         1.4462;     Plate",
   "Hastelloy-C22;          DIN Nr: ASTM UNS;       NiCr21Mo14W 2277 B575 N06022;   2.4602;     Plate",
   "HSLA S690Q;             EN10149-2 UNS;          S690Q;                          1.8974;     Plate",
   "Inconel (Alloy) 600;    Nicrofer 7216H;         NiCr15Fe, Alloy 600 ;           2.4816;     Plate",
   "P265GH (HII);           EN10028-2 UNS;          P265GH(HII);                    1.0425;     Plate",
   "P355NH;                 EN10028-3;              P355NH;                         1.0565;     Plate",
   "S235JR;                 EN10025 UNS;            S235JR;                         1.0038;     Struc-Steel",
   "S355J2;                 EN10025-2;              S355J2;                         1.0570;     Shaft-mat",
   "SS 304L;                EN10088-2;              X2CrNi19-11, S30403;            1.4306;     Plate",
   "SS 316L;                EN10088-2;              X2CrNiMo17-12-2, S31603;        1.4404;     Plate",
   "SS 316TI;               EN10088-2;              X6CrNiMoTi17-12-2, S31635;      1.4571;     Plate",
   "SuperDuplex;            --;                     X2CrNiMoN22-5-3 saisna;         1.4501;     Plate",
   "Titanium-ür 2;          ASTM UNS niN;           B265/348-Gr2;                   3.7035;     Plate"}


    Public oWord As Word.Application
    Private stringSplitOptons As Object
    ' see https://support.microsoft.com/en-us/help/316383/how-to-automate-word-from-visual-basic--net-to-create-a-new-document

    Private Sub Generate_word_doc()
        Dim oDoc As Word.Document
        Dim oPara3 As Word.Paragraph
        Dim ufilename As String
        Dim pathname As String
        Dim style1 As String

        '----------- Select Word style -----------------
        style1 = "N:\VERKOOP\Tekst\Quote_text_block\VTK_Fan_Quote.dotm"

        'Start Word and open the document template. 
        oWord = CType(CreateObject("Word.Application"), Word.Application)
        oWord.Visible = False
        oWord.ScreenUpdating = False
        ProgressBar1.Visible = True

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
            If ProgressBar1.Value > 99 Then ProgressBar1.Value = 1
            ProgressBar1.Value += 1
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

        '---------- Fan----------------
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
        Find_rep(Label13.Text, TextBox16.Text)
        Find_rep(Label14.Text, TextBox17.Text)
        Find_rep(Label15.Text, TextBox18.Text)
        Find_rep(Label16.Text, TextBox19.Text)
        Find_rep(Label17.Text, TextBox20.Text)
        Find_rep(Label18.Text, TextBox21.Text)
        Find_rep(Label19.Text, TextBox22.Text)
        Find_rep(Label20.Text, TextBox23.Text)

        Find_rep(Label21.Text, ComboBox1.Text)
        Find_rep(Label22.Text, ComboBox2.Text)
        Find_rep(Label23.Text, ComboBox3.Text)

        Find_rep(Label24.Text, TextBox24.Text)
        Find_rep(Label25.Text, TextBox25.Text)
        Find_rep(Label26.Text, TextBox26.Text)
        Find_rep(Label27.Text, TextBox27.Text)

        '---------- General------------------
        Find_rep(Label40.Text, ComboBox11.Text)     '_Mat_impellar
        Find_rep(Label42.Text, ComboBox12.Text)     '_Mat_casing
        Find_rep(Label47.Text, ComboBox13.Text)     '_mat_shaft

        Find_rep("_Comments", TextBox4.Text)
        Find_rep("_Comments2", TextBox4.Text)

        '---------- Conveyor----------------
        Find_rep(Label41.Text, NumericUpDown7.Value.ToString)   'Length
        Find_rep(Label43.Text, ComboBox7.Text)                  'Diameter flight
        Find_rep(Label32.Text, ComboBox8.Text)                  'Fligh pitch
        Find_rep(Label29.Text, NumericUpDown1.Value.ToString)

        Find_rep(Label53.Text, NumericUpDown3.Value.ToString)
        Find_rep(Label51.Text, NumericUpDown4.Value.ToString)
        Find_rep(Label50.Text, NumericUpDown5.Value.ToString)
        Find_rep(Label39.Text, NumericUpDown8.Value.ToString)   'Inspect doors
        Find_rep(Label46.Text, ComboBox8.Text)                  'Make drive
        Find_rep(Label45.Text, NumericUpDown6.Value.ToString)   'Speed
        Find_rep(Label38.Text, NumericUpDown2.Value.ToString)   'Flight thick

        '==================== backup final product===============
        ufilename = "Quote_" & TextBox01.Text & "_" & TextBox02.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".docx"

        If Directory.Exists(dirpath_Backup) Then
            ufilename = dirpath_Backup & ufilename
        Else
            ufilename = dirpath_Home_GP & ufilename
        End If
        ProgressBar1.Visible = False
        oWord.Visible = True
        oWord.ScreenUpdating = True
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

        ListBox1.Items.Add("Cyclone" & vbTab & vbTab & "1100")
        ListBox1.Items.Add("Filter" & vbTab & vbTab & "1500")
        ListBox1.Items.Add("Heater" & vbTab & vbTab & "2100")
        ListBox1.Items.Add("Demper" & vbTab & vbTab & "2700")
        ListBox1.Items.Add("Ringduct" & vbTab & vbTab & "3000")
        ListBox1.Items.Add("Piping" & vbTab & vbTab & "3100")
        ListBox1.Items.Add("Supports" & vbTab & vbTab & "3500")
        ListBox1.Items.Add("Valve" & vbTab & vbTab & "3600")

        ListBox2.Items.Add("Fan" & vbTab & vbTab & "4000")
        ListBox2.Items.Add("Conveyor" & vbTab & vbTab & "4400")
        ListBox2.Items.Add("Dewatering screw" & vbTab & "4500")
        ListBox2.Items.Add("Mixer" & vbTab & vbTab & "5600")
        ListBox2.Items.Add("Inwerprad" & vbTab & "6000")
        ListBox2.Items.Add("Disintegrator" & vbTab & "6100")
        ListBox2.Items.Add("Sluice" & vbTab & vbTab & "6200")
        ListBox2.Items.Add("Flap valve" & vbTab & "6300")
        ListBox2.Items.Add("Metal trap" & vbTab & "6400")
        ListBox2.Items.Add("Mill" & vbTab & vbTab & "6500")
        ListBox2.Items.Add("Sieve" & vbTab & vbTab & "6600")
        ListBox2.Items.Add("Pump" & vbTab & vbTab & "7000")

        ListBox3.Items.Add("Hopper" & vbTab & vbTab & "5900")
        ListBox3.Items.Add("Tank" & vbTab & vbTab & "7100")
        ListBox3.Items.Add("Struc. steel" & vbTab & "8000")
        ListBox3.Items.Add("Others" & vbTab & vbTab & "9000")

        Combo_init_atex()
        Combo_init_dia()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox01.Text.Trim.Length > 0 And TextBox07.Text.Trim.Length > 0 Then
            Save_tofile()
        Else
            MessageBox.Show("Complete Quote number and Customer tag" & vbCrLf & "Then the file can be saved")
        End If
    End Sub
    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()
        Dim temp_string, user As String
        user = Trim(Environment.UserName)         'User name on the screen
        Dim filename As String = "Quote_select_" & TextBox01.Text & "_" & TextBox07.Text & DateTime.Now.ToString("_yyyy_MM_dd_") & user & ".vtkq"
        Dim all_num, all_combo, all_check, all_text As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox02.Text) Then
            TextBox02.Text = "name"
        End If

        temp_string = TextBox01.Text & ";" & TextBox02.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(System.Windows.Forms.ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As System.Windows.Forms.ComboBox = CType(all_combo(i), System.Windows.Forms.ComboBox)
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
            MessageBox.Show("Line 204, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Check_directories()
        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home_GP)) Then System.IO.Directory.CreateDirectory(dirpath_Home_GP)
            If (Not System.IO.Directory.Exists(dirpath_Block)) Then System.IO.Directory.CreateDirectory(dirpath_Block)
            If (Not System.IO.Directory.Exists(dirpath_Backup)) Then System.IO.Directory.CreateDirectory(dirpath_Backup)
        Catch ex As Exception
            MessageBox.Show("Line 214, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    'Retrieve control settings from file
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

        ProgressBar1.Visible = True
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

            ProgressBar1.Value = 10

            '----- project data -----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split the read file content
            TextBox01.Text = words(0)                  'Project number
            TextBox02.Text = words(1)                  'Item no
            ProgressBar1.Value = 20

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
            ProgressBar1.Value = 30

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
            ProgressBar1.Value = 50

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
        ProgressBar1.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Read_file()
        MessageBox.Show("oo")
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
        Dim cnt As Integer = 1
        Dim all_check As New List(Of Control)

        TextBox04.Multiline = True
        TextBox04.Clear()

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Text).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
            If grbx.Checked = True Then
                TextBox04.Text &= grbx.Text & Environment.NewLine
            End If
        Next
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, CheckBox179.CheckStateChanged, CheckBox178.CheckStateChanged, CheckBox176.CheckStateChanged, CheckBox175.CheckStateChanged, CheckBox174.CheckStateChanged, CheckBox173.CheckStateChanged, CheckBox26.CheckStateChanged, CheckBox25.CheckStateChanged, CheckBox24.CheckStateChanged, CheckBox23.CheckStateChanged, CheckBox22.CheckStateChanged, CheckBox118.CheckStateChanged, CheckBox201.CheckedChanged, CheckBox200.CheckedChanged, CheckBox199.CheckedChanged, CheckBox198.CheckedChanged, CheckBox197.CheckedChanged, CheckBox196.CheckedChanged, CheckBox195.CheckedChanged, CheckBox194.CheckedChanged, CheckBox88.CheckStateChanged, CheckBox141.CheckStateChanged, CheckBox156.CheckedChanged, CheckBox155.CheckedChanged, CheckBox39.CheckedChanged, CheckBox38.CheckedChanged, CheckBox31.CheckedChanged, CheckBox30.CheckedChanged, CheckBox29.CheckedChanged, CheckBox162.CheckedChanged, CheckBox161.CheckedChanged, CheckBox167.CheckedChanged, CheckBox160.CheckedChanged, CheckBox213.CheckedChanged, CheckBox183.CheckedChanged, CheckBox11.CheckedChanged, CheckBox90.CheckedChanged, CheckBox33.CheckedChanged, CheckBox142.CheckedChanged, CheckBox133.CheckedChanged, CheckBox131.CheckedChanged, CheckBox216.CheckedChanged, CheckBox124.CheckedChanged, CheckBox123.CheckedChanged, CheckBox121.CheckedChanged, CheckBox6.CheckedChanged, CheckBox3.CheckedChanged, CheckBox9.CheckedChanged, CheckBox12.CheckedChanged, CheckBox8.CheckedChanged, CheckBox5.CheckedChanged, CheckBox165.CheckedChanged, CheckBox164.CheckedChanged, CheckBox18.CheckedChanged, CheckBox17.CheckedChanged, CheckBox16.CheckedChanged, CheckBox15.CheckedChanged, CheckBox1.CheckedChanged
        Check_combinations()
    End Sub
    'Check for groupbox checked > 1 or 0
    Private Sub Check_group1(ggg As System.Windows.Forms.GroupBox)
        Dim no_checked As Integer = 0

        For Each kctrl In ggg.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        ggg.BackColor = CType(IIf(no_checked > 1 Or no_checked = 0, Color.Red, SystemColors.Window), Color)
    End Sub
    'Check for groupbox > 1 checked
    Private Sub Check_group2(ggg As System.Windows.Forms.GroupBox)
        Dim no_checked As Integer = 0

        For Each kctrl In ggg.Controls
            If (kctrl.GetType() Is GetType(System.Windows.Forms.CheckBox)) Then
                Dim grbx As System.Windows.Forms.CheckBox = CType(kctrl, System.Windows.Forms.CheckBox)
                If grbx.Checked Then no_checked += 1
            End If
        Next
        ggg.BackColor = CType(IIf(no_checked > 1, Color.Red, SystemColors.Window), Color)
    End Sub
    Private Sub Check_combinations()
        Check_group1(GroupBox3) '(Bearings)
        Check_group2(GroupBox4) '(Electrical)
        Check_group1(GroupBox18) '(Arrangement)
        Check_group1(GroupBox19) '(Casing)
        Check_group1(GroupBox21) '(Vane type)
        Check_group1(GroupBox23) '(Seal type)
        Check_group1(GroupBox29) '(Coupling)
        Check_group1(GroupBox30) '(Coupling guard)
        Check_group1(GroupBox32) '(Motor)
        Check_group2(GroupBox34) '(VSD)
        Check_group2(GroupBox40) '(Motor efficiency)
        Check_group2(GroupBox41) '(Vibration isolation)
        Check_group2(GroupBox42) '(Vib sensor)
        Check_group2(GroupBox43) '(Bearing options temp)
        Check_group2(GroupBox44) '(Electrical)
        Check_group1(GroupBox45) '(Paint)
        Check_group2(GroupBox46) '(Space heater)
    End Sub

    Private Sub Combo_init_dia()
        ComboBox7.Items.Clear()
        ComboBox8.Items.Clear()
        ComboBox9.Items.Clear()

        '-------Fill combobox------------------
        For hh = 0 To Flight_dia.Length - 1                'Fill combobox 
            ComboBox7.Items.Add(Flight_dia(hh))
        Next hh

        '-------Fill combobox------------------
        For hh = 0 To flight_pitch.Length - 1               'Fill combobox 
            ComboBox8.Items.Add(flight_pitch(hh))
        Next hh

        '-------Fill combobox------------------
        For hh = 0 To drive_make.Length - 1               'Fill combobox 
            ComboBox9.Items.Add(drive_make(hh))
        Next hh

        ComboBox7.SelectedIndex = 2
        ComboBox8.SelectedIndex = 1
        ComboBox9.SelectedIndex = 0
    End Sub

    Private Sub Combo_init_atex()
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox11.Items.Clear()
        ComboBox12.Items.Clear()
        ComboBox13.Items.Clear()

        '-------Fill combobox, zone------------------
        For hh = 0 To atex_zone.Length - 1                'Fill combobox 
            ComboBox1.Items.Add(atex_zone(hh))
            ComboBox4.Items.Add(atex_zone(hh))
            ComboBox10.Items.Add(atex_zone(hh))
        Next hh

        '-------Fill combobox, temp------------------
        For hh = 0 To atex_temp.Length - 1                'Fill combobox 
            ComboBox3.Items.Add(atex_temp(hh))
            ComboBox5.Items.Add(atex_temp(hh))
        Next hh

        '-------Fill combobox, group------------------
        For hh = 0 To atex_group.Length - 1                'Fill combobox 
            ComboBox2.Items.Add(atex_group(hh))
            ComboBox6.Items.Add(atex_group(hh))
        Next hh

        '-------Fill combobox, materials--------------
        Dim words() As String
        Dim separators() As String = {";"}

        For hh = 0 To steel.Length - 1            'Fill combobox 
            words = steel(hh).Split(separators, StringSplitOptions.None)
            ComboBox11.Items.Add(LTrim(words(0)))
            ComboBox12.Items.Add(LTrim(words(0)))
            ComboBox13.Items.Add(LTrim(words(0)))
        Next hh

        ComboBox1.SelectedIndex = 2     'Zone
        ComboBox3.SelectedIndex = 2     'Temp
        ComboBox2.SelectedIndex = 1     'group

        ComboBox4.SelectedIndex = 4     'Zone dust
        ComboBox10.SelectedIndex = 5    'Zone dust
        ComboBox5.SelectedIndex = 2     'Temp
        ComboBox6.SelectedIndex = 1     'group

        ComboBox11.SelectedIndex = 5     'Steel impeller (domex)
        ComboBox12.SelectedIndex = 10    'Steel casing
        ComboBox13.SelectedIndex = 0     'Steel shaft
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Dim objBooks As Excel.Workbooks
        Dim objSheets As Excel.Sheets
        Dim objSheet As Excel._Worksheet
        Dim range As Excel.Range
        Dim z As Integer = 0
        Dim temp As String()

        Try
            'Create an array.
            temp = TextBox04.Text.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

            ' Create a new instance of Excel and start a new workbook.
            objApp = New Excel.Application()
            objBooks = objApp.Workbooks
            objBook = objBooks.Add
            objSheets = objBook.Worksheets

            objSheet = CType(objSheets(1), Excel._Worksheet)

            'Get the range where the starting cell has the address
            'm_sStartingCell and its dimensions are m_iNumRows x m_iNumCols.
            range = objSheet.Range("A1", Reflection.Missing.Value)
            range = range.Resize(temp.Length + 20, 2)

            Dim saRet(100, 1) As String
            saRet(0, 0) = "VTK Quote summary"
            saRet(1, 0) = Label1.Text
            saRet(1, 1) = TextBox01.Text  'Project
            saRet(2, 0) = Label3.Text
            saRet(2, 1) = TextBox07.Text  'Tag
            saRet(3, 0) = Label6.Text
            saRet(3, 1) = TextBox02.Text  '_Cust_name
            saRet(4, 0) = Label7.Text
            saRet(4, 1) = TextBox11.Text  '_Cust_ref
            saRet(5, 0) = Label8.Text
            saRet(5, 1) = TextBox12.Text  '_Cust_proj
            saRet(6, 0) = Label9.Text
            saRet(6, 1) = TextBox13.Text  '_Contact
            saRet(7, 0) = Label11.Text
            saRet(7, 1) = TextBox14.Text  '_fan_modelnr
            saRet(8, 0) = Label12.Text
            saRet(8, 1) = TextBox15.Text  '_Model
            saRet(9, 0) = Label13.Text
            saRet(9, 1) = TextBox16.Text  '_fan type
            '----------------------
            saRet(10, 0) = Label26.Text
            saRet(10, 1) = TextBox26.Text  'Orientation
            saRet(11, 0) = Label27.Text
            saRet(11, 1) = TextBox27.Text  'Orientation
            '---------- ATEX---------
            saRet(12, 0) = Label21.Text
            saRet(12, 1) = ComboBox1.SelectedItem.ToString
            saRet(13, 0) = Label22.Text
            saRet(13, 1) = ComboBox2.SelectedItem.ToString
            saRet(14, 0) = Label23.Text
            saRet(14, 1) = ComboBox3.SelectedItem.ToString
            '---------------
            saRet(15, 0) = Label25.Text
            saRet(15, 1) = TextBox25.Text  'T_design

            z = 16
            For Each Line As String In temp
                If Line.Length > 0 Then
                    saRet(z, 0) = Line.Substring(0, 4)
                    saRet(z, 1) = Line.Remove(0, 6)
                End If
                z += 1
            Next
            range.Value = saRet 'Set the range value to the array.
            range.ColumnWidth = 30

        Catch ex As Exception

        Finally
            'Return control of Excel to the user.
            objApp.Visible = True
            objApp.UserControl = True

            'Clean up a little.
            range = Nothing
            objSheet = Nothing
            objSheets = Nothing
            objBooks = Nothing
        End Try
    End Sub
    'Steel selection is changed
    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        Dim separators() As String = {";"}

        Dim words() As String = steel(ComboBox11.SelectedIndex).Split(separators, StringSplitOptions.None)
        TextBox29.Text = LTrim(words(1))
        TextBox30.Text = LTrim(words(3))
    End Sub
    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        Dim separators() As String = {";"}

        Dim words() As String = steel(ComboBox12.SelectedIndex).Split(separators, StringSplitOptions.None)
        TextBox31.Text = LTrim(words(1))
        TextBox32.Text = LTrim(words(3))
    End Sub
    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        Dim separators() As String = {";"}

        Dim words() As String = steel(ComboBox13.SelectedIndex).Split(separators, StringSplitOptions.None)
        TextBox33.Text = LTrim(words(1))
        TextBox34.Text = LTrim(words(3))
    End Sub
End Class
