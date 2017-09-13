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


    Private Sub Impeller_stress_to_word()
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara3 As Word.Paragraph
        Dim ufilename As String

        'Start Word and open the document template. 
        oWord = CType(CreateObject("Word.Application"), Word.Application)
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add()
        oPara1.Range.Text = "VTK SALES"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = 16
        oPara1.Range.Font.Bold = CInt(True)
        oPara1.Format.SpaceAfter = 2                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = 11
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = CInt(False)
        oPara2.Range.Text = "Quotation for customer " & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 9
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        oTable.Cell(1, 1).Range.Text = "Project Name"
        oTable.Cell(1, 2).Range.Text = TextBox1.Text
        oTable.Cell(2, 1).Range.Text = "Project number "
        oTable.Cell(2, 2).Range.Text = TextBox1.Text
        oTable.Cell(3, 1).Range.Text = "Author "
        oTable.Cell(3, 2).Range.Text = Environment.UserName
        oTable.Cell(4, 1).Range.Text = "Date "
        oTable.Cell(4, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        oTable.Cell(5, 1).Range.Text = "Fan type"
        oTable.Cell(5, 2).Range.Text = Label1.Text

        oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        'oPara2.Range.InsertBreak()                               'New page

        '===================== inserting blocks========================
        '--- insert block #1
        oPara3 = oDoc.Content.Paragraphs.Add()
        If CheckBox1.Checked Then oPara3.Range.InsertFile(dirpath_Block & "QG_Gen_GB_001.docx")

        '--- insert block #2
        oPara3 = oDoc.Content.Paragraphs.Add()
        If CheckBox2.Checked Then oPara3.Range.InsertFile(dirpath_Block & "QG_Gen_GB_002.docx")

        '--- insert block #3
        oPara3 = oDoc.Content.Paragraphs.Add()
        If CheckBox3.Checked Then oPara3.Range.InsertFile(dirpath_Block & "QG_Gen_GB_003.docx")

        '==================== backup final product===============
        ufilename = "Quote_" & TextBox1.Text & "_" & TextBox2.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".docx"

        If Directory.Exists(dirpath_Backup) Then
            ufilename = dirpath_Backup & ufilename
        Else
            ufilename = dirpath_Home_GP & ufilename
        End If
        oWord.ActiveDocument.SaveAs(ufilename.ToString)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Check_directories()
        Impeller_stress_to_word()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox3.Text =
        "File naming convention" & vbCrLf & vbCrLf &
        "Text block location is " & vbTab & dirpath_Block.ToString & vbCrLf &
        "Quote backup location is " & vbTab & dirpath_Backup.ToString & vbCrLf & vbCrLf &
        "File name General data is " & vbTab & "QG_Gen_GB_001.docx " & vbCrLf &
        "File name Fan is " & vbTab & vbTab & "QG_fan_GB_001.docx " & vbCrLf &
        "File name Cyclone is " & vbTab & vbTab & "QG_Cyc_GB_001.docx " & vbCrLf &
        "File name Dryer is " & vbTab & vbTab & "QG_Dry_GB_001.docx " & vbCrLf &
        "File name Blower is " & vbTab & vbTab & "QG_Blo_GB_001.docx " & vbCrLf &
        "File name Othere is " & vbTab & vbTab & "QG_Oth_GB_001.docx " & vbCrLf &
        "File name Commercial is " & vbTab & "QG_Com_GB_001.docx " & vbCrLf &
        " "
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text.Trim.Length > 0 And TextBox2.Text.Trim.Length > 0 Then
            Save_tofile()
        Else
            MessageBox.Show("Complete Quote and item number")
        End If
    End Sub
    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()

        Dim temp_string, user As String

        user = Trim(Environment.UserName)         'User name on the screen
        Dim filename As String = "Quote_select_" & TextBox1.Text & "_" & TextBox2.Text & DateTime.Now.ToString("_yyyy_MM_dd_") & user & ".vtkq"
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox2.Text) Then
            TextBox2.Text = "name"
        End If

        temp_string = TextBox1.Text & ";" & TextBox2.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"


        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
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
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
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

            '----- retrieve case condition-----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split the read file content
            TextBox1.Text = words(0)                  'Project number
            TextBox2.Text = words(1)                  'Item no

            '---------- terugzetten checkbox controls -----------------
            'FindControlRecursive(all_check, Me, GetType(CheckBox))
            FindControlRecursive(all_check, Me, GetType(System.Windows.Forms.CheckBox))      'Find the control
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As System.Windows.Forms.CheckBox = CType(all_check(i), System.Windows.Forms.CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last checkbox not found in file")
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
End Class
