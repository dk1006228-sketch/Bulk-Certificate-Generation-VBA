Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateCertificatesPDF()

    Dim wdApp As Object
    Dim wdDoc As Object

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long

    Dim TemplatePath As String
    Dim OutputFolder As String

    Dim EmployeeID As String
    Dim CandidateName As String
    Dim CompletionDate As String
    Dim CertificateID As String

    Dim PDFName As String

    'Select Word Template
    With Application.FileDialog(msoFileDialogFilePicker)

        .Title = "Select Certificate Template"
        .Filters.Clear
        .Filters.Add "Word Files", "*.docx;*.doc"

        If .Show <> -1 Then Exit Sub

        TemplatePath = .SelectedItems(1)

    End With

    'Select Output Folder
    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Output Folder"

        If .Show <> -1 Then Exit Sub

        OutputFolder = .SelectedItems(1)

    End With

    Set ws = ActiveSheet

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Open Word
    On Error Resume Next

    Set wdApp = GetObject(, "Word.Application")

    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If

    On Error GoTo 0

    wdApp.Visible = False

    'Loop Through Records
    For i = 2 To LastRow

        EmployeeID = Trim(ws.Cells(i, 1).Value)
        CandidateName = Trim(ws.Cells(i, 2).Value)
        CompletionDate = Trim(ws.Cells(i, 3).Text)
        CertificateID = Trim(ws.Cells(i, 4).Value)

        If EmployeeID <> "" Then

            Set wdDoc = wdApp.Documents.Open(TemplatePath)

            'Replace Values
            ReplaceInDocument wdDoc, "<<Candidate Name>>", CandidateName
            ReplaceInDocument wdDoc, "<<Employee ID>>", EmployeeID
            ReplaceInDocument wdDoc, "<<Certificate ID>>", CertificateID
            ReplaceInDocument wdDoc, "<<Completion Date>>", CompletionDate

            PDFName = OutputFolder & "\" & EmployeeID & "_Certificate.pdf"

            wdDoc.ExportAsFixedFormat _
                OutputFileName:=PDFName, _
                ExportFormat:=17

            wdDoc.Close False

        End If

    Next i

    wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing

    MsgBox "All certificates generated successfully!", vbInformation

End Sub


Private Sub ReplaceInDocument(ByVal wdDoc As Object, _
                              ByVal FindText As String, _
                              ByVal ReplaceText As String)

    Dim shp As Object
    Dim rngStory As Object

    'Replace in normal document text
    For Each rngStory In wdDoc.StoryRanges

        With rngStory.Find

            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = FindText
            .Replacement.Text = ReplaceText
            .Wrap = 1

            .Execute Replace:=2

        End With

    Next rngStory

    'Replace inside Text Boxes / Shapes
    For Each shp In wdDoc.Shapes

        On Error Resume Next

        If shp.TextFrame.HasText Then

            shp.TextFrame.TextRange.Text = _
            VBA.Replace(shp.TextFrame.TextRange.Text, _
                        FindText, ReplaceText)

        End If

        On Error GoTo 0

    Next shp

End Sub

