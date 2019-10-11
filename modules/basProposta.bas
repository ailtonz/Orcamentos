Attribute VB_Name = "basProposta"
Sub Gerar_Proposta()



'Dim wdApp As Word.Application, wdDoc As Word.Document
''Dim tdate As Date
''Dim contract As String
'Dim Arquivo As String: Arquivo = "Proposta - Springer Brasil - 2014.docx"
'
''contract = "Verizon FIOS"
''tdate = Date
'
'On Error Resume Next
'
'Set wdApp = GetObject(, "Word.Application")
'
'If Err.Number <> 0 Then 'Word isn't already running
'    Set wdApp = CreateObject("Word.Application")
'End If
'
'On Error GoTo 0
'
'Set wdDoc = wdApp.Documents.Open(ActiveWorkbook.Path & "\db\" & Arquivo, ReadOnly:=True)
'
'wdApp.Visible = True
''Writing Variables from Excel to the Checklist word doc.
'wdDoc.Bookmarks("N_CONTROLE").Range.Text = ActiveSheet.Name
'wdDoc.Bookmarks("CLIENTE").Range.Text = Range("C4").Value
'wdDoc.Bookmarks("RESPONSAVEL").Range.Text = Range("C5").Value
'wdDoc.Bookmarks("PROJETO").Range.Text = Range("C6").Value
'wdDoc.Bookmarks("JOURNAL").Range.Text = Range("C9").Value
'wdDoc.Bookmarks("AUTOR").Range.Text = Range("C10").Value
'wdDoc.Bookmarks("PUBLISHER").Range.Text = Range("C8").Value
'
'wdDoc.Bookmarks("FORMATO").Range.Text = Range("C29").Value
'wdDoc.Bookmarks("N_PAGINAS").Range.Text = Range("C27").Value
'
''wdDoc.Bookmarks("IDIOMA").Range.Text = Range("C17").Value
''wdDoc.Bookmarks("VOLUME").Range.Text = Range("").Value
''wdDoc.Bookmarks("PRC_VENDA").Range.Text = Range("").Value
''wdDoc.Bookmarks("PRC_TOTAL").Range.Text = Range("").Value
'
'wdDoc.Bookmarks("G_CONTAS").Range.Text = Range("C3").Value
''wdDoc.Bookmarks("TELEFONE").Range.Text = Range("I2").Value
''wdDoc.Bookmarks("CELULAR_01").Range.Text = Range("I2").Value
''wdDoc.Bookmarks("CELULAR_02").Range.Text = Range("I2").Value
''wdDoc.Bookmarks("ID_NEXTEL").Range.Text = Range("I2").Value
'
'wdDoc.SaveAs pathDesktopAddress & "\" & Now() & "_" & Arquivo
'wdDoc.Close
'
'wdApp.Application.Quit
End Sub


Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = TextToUse
    ActiveDocument.Bookmarks.Add BookmarkToUpdate, BMRange
End Sub
