Sub AcquérirValeursExternes()

    Dim startTagOpen  As Range
    Dim startTagClose As Range
    Dim startTag As Range
    Dim nextStartTagOpen  As Range
    Dim stopTag As Range
    Dim tagContent As Range
    Dim referenceReplaced As Integer
    Dim startTagText As String
    Dim excelFileName As String
    Dim worksheetName As String
    Dim cellCoordinate As String
    Dim parameters() As String
    Dim excelFile As Workbook
    Dim path As String
    Dim importedValue As String
    
    referenceReplaced = 0
    
    ' Init ranges
    Set stopTag = ActiveDocument.Range(0, 0)
    
    Do
        ' Reset error handling
        On Error GoTo 0
        
        ' Define startTag
        Set startTagOpen = ActiveDocument.Range(stopTag.End, ActiveDocument.Range.End)
        startTagOpen.TextRetrievalMode.IncludeHiddenText = True
        startTagOpen.Find.Execute FindText:="{REF"
        
        ' Validate stop tag or end SubRoutine
        If startTagOpen <> "{REF" Then
            Exit Do
        End If
        
        Set startTagClose = ActiveDocument.Range(startTagOpen.Start + 1, ActiveDocument.Range.End)
        startTagClose.TextRetrievalMode.IncludeHiddenText = True
        startTagClose.Find.Execute FindText:="}"
        Set startTag = ActiveDocument.Range(startTagOpen.Start, startTagClose.End)
        
        ' Define stopTag
        Set stopTag = ActiveDocument.Range(startTagClose.Start, ActiveDocument.Range.End)
        stopTag.TextRetrievalMode.IncludeHiddenText = True
        stopTag.Find.Execute FindText:="{FINREF}"
        
        ' Validate that we don't find another "{REF" tag before the next "{FINREF}"
        Set nextStartTagOpen = ActiveDocument.Range(startTagClose.Start, ActiveDocument.Range.End)
        nextStartTagOpen.TextRetrievalMode.IncludeHiddenText = True
        nextStartTagOpen.Find.Execute FindText:="{REF"
        
        If nextStartTagOpen <> "{REF" Then
            Set nextStartTagOpen = ActiveDocument.Range(ActiveDocument.Range.End - 1, ActiveDocument.Range.End)
        End If
        
        ' Validate stop tag or skip to next iteration
        If stopTag <> "{FINREF}" Or (stopTag.Start > nextStartTagOpen.Start) Then
            MsgBox "Marqueur de fin ({FINREF}) manquant pour " & startTag
            Set stopTag = ActiveDocument.Range(startTagClose.End, startTagClose.End + 1)
            GoTo Continue
        End If
        
        ' Reveal tags (for parsing purposes)
        startTag.Font.Hidden = False
        stopTag.Font.Hidden = False
        
        ' Clear content between reference tags
        Set tagContent = ActiveDocument.Range(startTag.End, stopTag.Start)
        tagContent.TextRetrievalMode.IncludeHiddenText = True
        If stopTag.Start > startTag.End Then
            tagContent.Delete
        End If
        
        ' Parse content of startTag
        startTagText = startTag.Text
        parameters = Split(Mid(startTagText, 6, Len(startTagText) - 6), ",")
        excelFileName = CleanInputString(parameters(0))
        worksheetName = CleanInputString(parameters(1))
        cellCoordinate = UCase(CleanInputString(parameters(2)))
        
        ' Open Excel file
        On Error GoTo UnableToOpenExcel
        path = ""
        If InStr(excelFileName, ":") = 0 Then
            path = Application.ActiveDocument.path & "\"
        End If
        
        Set excelFile = Workbooks.Open(path & excelFileName)
        
        ' Extract value from excel file
        importedValue = excelFile.Worksheets(worksheetName).Range(cellCoordinate).Value
        excelFile.Close
        
        ' Fill space between tags
        tagContent.Text = importedValue
        
        ' Check if value is empty
        If Len(importedValue) = 0 Then
            MsgBox "Aucune valeur à extraire de " & excelFileName & " " & worksheetName & " " & cellCoordinate, vbExclamation
        End If
        
        ' Hide reference tag
        startTag.Font.Hidden = True
        stopTag.Font.Hidden = True
        tagContent.Font.Hidden = False
        referenceReplaced = referenceReplaced + 1
        
Continue:
    Loop

    ' Normal end of Subroutine
    MsgBox referenceReplaced & " réferences remplacées"
    Exit Sub
    
UnableToOpenExcel:
    MsgBox "Impossible d'ouvrir " & excelFileName & " " & worksheetName & " " & cellCoordinate, vbExclamation
    GoTo Continue
    
Exit Sub
    
    
End Sub

Sub AfficherRéférences()
    Dim startTagOpen  As Range
    Dim startTagClose As Range
    Dim startTag As Range
    Dim stopTag As Range
    Dim tagContent As Range
        
    ' Init ranges
    Set stopTag = ActiveDocument.Range(0, 0)
    
    Do
        ' Find startTag
        Set startTagOpen = ActiveDocument.Range(stopTag.End, ActiveDocument.Range.End)
        startTagOpen.TextRetrievalMode.IncludeHiddenText = True
        startTagOpen.Find.Execute FindText:="{REF"
        
        ' Validate stop tag or end SubRoutine
        If startTagOpen <> "{REF" Then
            Exit Do
        End If
        
        Set startTagClose = ActiveDocument.Range(startTagOpen.Start + 1, ActiveDocument.Range.End)
        startTagClose.TextRetrievalMode.IncludeHiddenText = True
        startTagClose.Find.Execute FindText:="}"
        Set startTag = ActiveDocument.Range(startTagOpen.Start, startTagClose.End)
        
        ' Find stopTag
        Set stopTag = ActiveDocument.Range(startTagClose.Start, ActiveDocument.Range.End)
        stopTag.TextRetrievalMode.IncludeHiddenText = True
        stopTag.Find.Execute FindText:="{FINREF}"
        
        
        ' Validate that we don't find another "{REF" tag before the next "{FINREF}"
        Set nextStartTagOpen = ActiveDocument.Range(startTagClose.Start, ActiveDocument.Range.End)
        nextStartTagOpen.TextRetrievalMode.IncludeHiddenText = True
        nextStartTagOpen.Find.Execute FindText:="{REF"
        
        If nextStartTagOpen <> "{REF" Then
            Set nextStartTagOpen = ActiveDocument.Range(ActiveDocument.Range.End - 1, ActiveDocument.Range.End)
        End If
        
        ' Validate stop tag or skip to next iteration
        If stopTag <> "{FINREF}" Or (stopTag.Start > nextStartTagOpen.Start) Then
            Set stopTag = ActiveDocument.Range(startTagClose.End, startTagClose.End + 1)
            GoTo ContinueShowRefTag
        End If
        
        startTag.Font.Hidden = False
        stopTag.Font.Hidden = False
ContinueShowRefTag:
    Loop
End Sub

Function CleanInputString(inputString As String) As String
    Dim result As String
    
    ' Remove double-quotes
    result = Replace(inputString, Chr(147), "")
    result = Replace(result, Chr(34), "")
    result = Replace(result, ChrW(8221), "")
    
    ' Remove leading and trailing spaces
    CleanInputString = Trim(result)
End Function
