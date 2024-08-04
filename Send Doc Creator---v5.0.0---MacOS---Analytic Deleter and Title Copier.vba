' SendDoc Creator v5.0.0---MacOS---Analytic Deleter and Title Copier
' https://github.com/KSXia/Verbatim-Send-Doc-Creator/tree/MacOS---Analytic-Deleter-and-Title-Copier
' Updated on 2024-08-03
' Thanks to Truf for providing the original version of the macro!
Sub CreateSendDoc()
    Dim StylesToDelete As Variant
    
    ' ---SET THE STYLES TO DELETE <<HERE>>!---
    ' Add the names of styles that you want to delete to this list in the StylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas.
    ' WARNING: There MUST be at least one style listed in the StylesToDelete array. It MUST NOT be empty!
    StylesToDelete = Array("Analytic", "Analytics", "Undertag")
    
    ' ---INITIAL CHECKS---
    Dim OriginalDoc As Document
    ' Assign the original document to a variable
    Set OriginalDoc = ActiveDocument
    
    ' Check if the original document has previously been saved
    If OriginalDoc.Path = "" Then
        ' If the original document has not been previously saved:
        MsgBox "The current document must be saved at least once.", Title:="Error in Creating Send Doc"
    
    Else
        ' If the original document has been previously saved:
        ' Assign the original document name to a variable
        Dim OriginalDocName As String
        OriginalDocName = OriginalDoc.Name
        
        Dim SendDoc As Document
        
        ' If the doc has been previously saved, create a copy of it to be the send doc
        Set SendDoc = Documents.Add(OriginalDoc.FullName)
        
        GreatestStyleIndex = UBound(StylesToDelete) - LBound(StylesToDelete)
        
        ' ---INITIAL SETUP---
        ' Disable error prompts in case one of the styles set to be deleted isn't present
        On Error Resume Next
        
        ' Disable screen updating for faster execution
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        ' ---STYLE DELETION---
        Dim CurrentStyleNumber As Integer
        For CurrentStyleNumber = 0 to GreatestStyleIndex Step +1
            Dim StyleToDelete As Style
             
            ' Specify the style to be deleted and delete it
            Set StyleToDelete = SendDoc.Styles(StylesToDelete(CurrentStyleNumber))
            
            ' Use Find and Replace to remove text with the specified style and delete it
            With SendDoc.Content.Find
                .ClearFormatting
                .Style = StyleToDelete
                .Replacement.ClearFormatting
                .Replacement.Text = ""
                .Format = True
                ' Disabling checks for the find process for optimization
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                .MatchPrefix = False
                .MatchSuffix = False
                ' Delete all text with the style to delete
                .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
            End With
        Next CurrentStyleNumber
        
        ' Reenable error prompts
        On Error GoTo 0
        
        ' ---SEND DOCUMENT TITLE COPIER---
        Dim ClipboardText As DataObject
        
        ' Set a variable to be the name of the send doc
        Dim SendDocName As String
        SendDocName = Left(OriginalDocName, Len(OriginalDocName) - 5) & " [S]"
        
        ' Put the name of the send doc into the clipboard
        Set ClipboardText = New DataObject
        ClipboardText.SetText SendDocName
        ClipboardText.PutInClipboard
        
        ' ---FINAL PROCESSES---
        ' Enable screen updating and alerts
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
    End If
End Sub