' ---Send Doc Creator v3.0.0---
' Updated on 2024-08-23.
' https://github.com/KSXia/Verbatim-Send-Doc-Creator
' Thanks to Truf for creating and providing the original "Create Send Doc" macro this macro is based on! You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/
Sub CreateSendDoc()
	Dim CopyFormattedTitle As Boolean
	Dim AutomaticallySaveSendDoc As Boolean
	Dim AutomaticallyCloseSavedSendDoc As Boolean
	Dim DeleteStyles As Boolean
	Dim StylesToDelete() As Variant
	Dim DeleteLinkedCharacterStyles As Boolean
	Dim LinkedCharacterStylesToDelete() As Variant
	Dim SendDocNamePrefix As String
	Dim SendDocNameSuffix As String
	
	' ---USER CUSTOMIZATION---
	' <<CUSTOMIZE THE SAVING MECHANISMS HERE!>>
	' If CopyFormattedTitle is set to True, this macro will copy the formatted name of the send doc into the clipboard.
	CopyFormattedTitle = False
	
	' If AutomaticallySaveSendDoc is set to True, this macro will automatically save the send doc.
	' WARNING: This feature to automatically save the send doc has LIMITED COMPATIBILITY! It might not work on MacOS.
	AutomaticallySaveSendDoc = False
	
	' If this macro is set to automatically save the send doc and AutomaticallyCloseSavedSendDoc is set to True, the send doc will automatically be closed after it is saved.
	AutomaticallyCloseSavedSendDoc = True
	
	' <<SET THE STYLES TO DELETE HERE!>>
	' Add the names of styles that you want to delete to the list in the StylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no styles will be deleted.
	StylesToDelete = Array("Undertag", "Analytic")
	
	' If DeleteStyles is set to True, the styles listed in the StylesToDelete array will be deleted. If DeleteStyles is set to False, the styles listed in the StylesToDelete array will not be deleted.
	' If you want to disable the deletion of the styles listed in the StylesToDelete array, set DeleteStyles to False.
	DeleteStyles = True
	
	' <<SET THE LINKED CHARACTER STYLES TO DELETE HERE!>>
	' A linked style will either apply the style to the entire paragraph or a selection of words depending on what you have selected. If you have clicked on a paragraph and have selected no text or have selected the entire paragraph, it will apply the paragraph variant of the style. If you have selected a subset of the paragraph, it will apply the character variant of the style to your selection. The options in this section control whether this macro will delete the instances of character variants of linked styles and which linked styles this macro will operate on.
	
	' If DeleteLinkedCharacterStyles is set to True, the character variants of the linked styles listed in the LinkedCharacterStylesToDelete array will be deleted. If DeleteLinkedCharacterStyles is set to False, they will not be deleted.
	DeleteLinkedCharacterStyles = True
	
	' Add the names of linked styles that you want to delete the character variant of to the list in the LinkedCharacterStylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no character variants of linked styles will be deleted.
	LinkedCharacterStylesToDelete = Array("Analytic")
	
	' <<SET HOW THE SEND DOC IS NAMED HERE!>>
	' Set SendDocNamePrefix to the prefix you want to add to the send doc name.
	' Make sure there are quotation marks around the prefix you want to insert into the send doc name!
	' If you do not want to insert a prefix into the send doc name, put nothing in-between the quotation marks. If you do this, you MUST have a suffix for the send doc name.
	SendDocNamePrefix = ""
	
	' Set SendDocNameSuffix to the suffix you want to add to the send doc name.
	' Make sure there are quotation marks around the suffix you want to insert into the send doc name!
	' If you do not want to insert a suffix into the send doc name, put nothing in-between the quotation marks. If you do this, you MUST have a prefix for the send doc name.
	SendDocNameSuffix = " [S]"
	
	' ---CHECK VALIDITY OF USER CONFIGURATION---
	' Check if there is either a prefix or suffix for the send doc name
	If SendDocNamePrefix = "" And SendDocNameSuffix = "" Then
		' If there is neither a prefix nor suffix for the send doc name:
		MsgBox "You have not set a suffix or prefix to add to the send doc name. Please set one in the macro settings and try again.", Title:="Error in Creating Send Doc"
		Exit Sub
	End If
	
	' ---INITIAL VARIABLE SETUP---
	Dim OriginalDoc As Document
	' Assign the original document to a variable
	Set OriginalDoc = ActiveDocument
	
	' Check if the original document has previously been saved
	If OriginalDoc.Path = "" Then
		' If the original document has not been previously saved:
		MsgBox "The current document must be saved at least once. Please save the current document and try again.", Title:="Error in Creating Send Doc"
		Exit Sub
	End If
	
	' Assign the original document name to a variable
	Dim OriginalDocName As String
	OriginalDocName = OriginalDoc.Name
	
	' ---INITIAL GENERAL SETUP---
	' Disable screen updating for faster execution
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---VARIABLE SETUP---
	Dim SendDoc As Document
	
	' If the doc has been previously saved, create a copy of it to be the send doc
	Set SendDoc = Documents.Add(OriginalDoc.FullName)
	
	Dim GreatestStyleIndex As Integer
	GreatestStyleIndex = UBound(StylesToDelete) - LBound(StylesToDelete)
	
	Dim GreatestLinkedCharacterStyleIndex As Integer
	GreatestLinkedCharacterStyleIndex = UBound(LinkedCharacterStylesToDelete) - LBound(LinkedCharacterStylesToDelete)
	
	' ---STYLE DELETION SETUP---
	' Disable error prompts in case one of the styles set to be deleted isn't present
	On Error Resume Next
	
	' ---PRE-PROCESSING FOR STYLE DELETION---
	' Use Find and Replace to replace paragraph marks in the character variants of linked styles set for deletion with paragraph marks in Tag style.
	' This ensures all paragraph marks in lines or paragraphs that have character variants of linked styles set to be delted are in Tag style so they do not get deleted in the style deletion stage of this macro.
	' Otherwise, lines ending in character variants of linked styles set to be delted may have their paragraph mark deleted and have the following line be merged into them, which can mess up the formatting of the line.
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleNameToProcessIndex As Integer
		For CurrentLinkedCharacterStyleNameToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) = LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) & " Char"
		Next CurrentLinkedCharacterStyleNameToProcessIndex
		
		Dim CurrentLinkedCharacterStyleToProcessIndex As Integer
		For CurrentLinkedCharacterStyleToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToProcess As Style
			
			Set LinkedCharacterStyleToProcess = SendDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToProcessIndex))
			
			With SendDoc.Content.Find
				.ClearFormatting
				.Text = "^p"
				.Style = LinkedCharacterStyleToProcess
				.Replacement.ClearFormatting
				.Replacement.Text = "^p"
				.Replacement.Style = "Tag Char"
				.Format = True
				' Ensure various checks are disabled to have the search properly function
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToProcessIndex
	End If
	
	' ---STYLE DELETION---
	If DeleteStyles = True Then
		Dim CurrentStyleToDeleteIndex As Integer
		For CurrentStyleToDeleteIndex = 0 to GreatestStyleIndex Step 1
			Dim StyleToDelete As Style
			
			' Specify the style to be deleted
			Set StyleToDelete = SendDoc.Styles(StylesToDelete(CurrentStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the specified style and delete it
			With SendDoc.Content.Find
				.ClearFormatting
				.Style = StyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentStyleToDeleteIndex
	End If
	
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleToDeleteIndex As Integer
		For CurrentLinkedCharacterStyleToDeleteIndex = 0 to GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToDelete As Style
			
			' Specify the linked style to delete the character variants of
			Set LinkedCharacterStyleToDelete = SendDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the character variants of the specified linked style and delete it
			With SendDoc.Content.Find
				.ClearFormatting
				.Style = LinkedCharacterStyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToDeleteIndex
	End If
	
	' ---POST STYLE DELETION PROCESSES---
	' Re-enable error prompts
	On Error GoTo 0
	
	' ---SEND DOCUMENT TITLE COPIER---
	If CopyFormattedTitle = True Then
		Dim ClipboardText As DataObject
		
		' Set a variable to be the name of the send doc
		Dim SendDocName As String
		SendDocName = SendDocNamePrefix & Left(OriginalDocName, Len(OriginalDocName) - 5) & SendDocNameSuffix
		
		' Put the name of the send doc into the clipboard
		Set ClipboardText = New DataObject
		ClipboardText.SetText SendDocName
		ClipboardText.PutInClipboard
	End If
	
	' ---SAVING THE SEND DOC---
	If AutomaticallySaveSendDoc = True Then
		Dim SavePath As String
		SavePath = OriginalDoc.Path & "\" & SendDocNamePrefix & Left(OriginalDocName, Len(OriginalDocName) - 5) & SendDocNameSuffix & ".docx"
		SendDoc.SaveAs2 Filename:=SavePath, FileFormat:=wdFormatDocumentDefault
		
		If AutomaticallyCloseSavedSendDoc = True Then
			SendDoc.Close SaveChanges:=wdSaveChanges
			MsgBox "The send doc is saved at " & SavePath, Title="Successfully Created and Saved Send Doc"
		End If
	End If
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub
