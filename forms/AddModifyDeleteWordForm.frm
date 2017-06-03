VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddModifyDeleteWordForm 
   Caption         =   "Add/Modify/Delete Word in Dictionary"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8340.001
   OleObjectBlob   =   "AddModifyDeleteWordForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddModifyDeleteWordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dictionaryLanguage As String

Private Sub AddOptionButton_Click()
    DoOperationCommandButton.Caption = "Add Word"
    CurrentWordTextBox.Visible = True
    CurrentWordTextBox.Value = ""
    WordsListBox.Visible = True
    CurrentSelectedTextLabel.Visible = False
    CurrentWordLabel.Visible = False
End Sub

Private Sub DeleteOptionButton_Click()
    DoOperationCommandButton.Caption = "Delete Word"
    CurrentWordTextBox.Visible = False
    WordsListBox.Visible = True
    CurrentSelectedTextLabel.Visible = True
    CurrentWordLabel.Visible = True
End Sub

Private Sub DoOperationCommandButton_Click()
    If AddOptionButton.Value = True Then
        SpellCheck.addWordToDictionary word:=CurrentWordTextBox.Value, language:=dictionaryLanguage
    End If
    If DeleteOptionButton.Value = True Then
        SpellCheck.removeWordFromDictionary word:=CurrentWordTextBox.Value, language:=dictionaryLanguage
    End If
    If EditOptionButton.Value = True Then
        For i = 0 To WordsListBox.ListCount - 1
            If WordsListBox.Selected(i) Then
                SpellCheck.editWordFromDictionary oldWord:=WordsListBox.List(i), newWord:=CurrentWordTextBox.Value, language:=dictionaryLanguage
                Exit For
            End If
        Next i
    End If
    
    WordsListBox.Clear
    Set c = SpellCheck.getAllWords(dictionaryLanguage)
    For Each Item In c
        WordsListBox.AddItem (Item)
    Next Item
End Sub

Private Sub EditOptionButton_Click()
    DoOperationCommandButton.Caption = "Edit Word"
    CurrentWordTextBox.Visible = True
    WordsListBox.Visible = True
    CurrentSelectedTextLabel.Visible = True
    CurrentWordLabel.Visible = True
    For i = 0 To WordsListBox.ListCount - 1
            If WordsListBox.Selected(i) Then
                CurrentWordLabel.Caption = WordsListBox.List(i)
                If EditOptionButton.Value = True Then
                    CurrentWordTextBox.Value = WordsListBox.List(i)
                End If
                Exit For
            End If
        Next i
End Sub

Private Sub UserForm_Initialize()
    Dim c As Collection
    dictionaryLanguage = SpellCheckingForm.dictionaryLanguage
    SpellCheck.currentSelectedLanguage = dictionaryLanguage
    Set c = SpellCheck.getAllWords(dictionaryLanguage)
    
    For Each Item In c
        WordsListBox.AddItem Item
    Next Item
    
End Sub

Private Sub WordsListBox_Click()
    For i = 0 To WordsListBox.ListCount - 1
            If WordsListBox.Selected(i) Then
                CurrentWordLabel.Caption = WordsListBox.List(i)
                If EditOptionButton.Value = True Then
                    CurrentWordTextBox.Value = WordsListBox.List(i)
                End If
                Exit For
            End If
    Next i
End Sub
