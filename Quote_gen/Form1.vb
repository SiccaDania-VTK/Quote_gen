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
        If CheckBox1.Checked Then oPara3.Range.InsertFile(dirpath_Block & "quote_block_0001.docx")

        '--- insert block #2
        oPara3 = oDoc.Content.Paragraphs.Add()
        If CheckBox2.Checked Then oPara3.Range.InsertFile(dirpath_Block & "quote_block_0002.docx")

        '--- insert block #3
        oPara3 = oDoc.Content.Paragraphs.Add()
        If CheckBox3.Checked Then oPara3.Range.InsertFile(dirpath_Block & "quote_block_0003.docx")

        '==================== backup final product===============
        ufilename = "Quote_" & TextBox1.Text & "_" & TextBox2.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".docx"

        If Directory.Exists(dirpath_Backup) Then
            ufilename = dirpath_Backup & ufilename
        Else
            ufilename = dirpath_Home_GP & ufilename
        End If
        oWord.ActiveDocument.SaveAs(ufilename)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home_GP)) Then System.IO.Directory.CreateDirectory(dirpath_Home_GP)
            If (Not System.IO.Directory.Exists(dirpath_Block)) Then System.IO.Directory.CreateDirectory(dirpath_Block)
            If (Not System.IO.Directory.Exists(dirpath_Backup)) Then System.IO.Directory.CreateDirectory(dirpath_Backup)
        Catch ex As Exception
        End Try
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

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub
End Class
