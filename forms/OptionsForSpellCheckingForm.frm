VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionsForSpellCheckingForm 
   Caption         =   "Options For Spell checking"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8625.001
   OleObjectBlob   =   "OptionsForSpellCheckingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionsForSpellCheckingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public currentWordSelection As Selection
Public selectedLanguage As String

Private Sub AddWordOptionButton_Click()
    CommandButton.Caption = "Add word"
End Sub

Private Sub CommandButton_Click()
    Dim i As Integer
    If AddWordOptionButton.Value = True Then
        Selection.Font.Underline = wdUnderlineNone
        SpellCheck.addToDictionary currentWordSelection.Text
    End If
    
    If ReplaceWordOptionButton.Value = True Then
        Selection.Font.Underline = wdUnderlineNone
        For i = 0 To RecommendedWordsListBox.ListCount - 1
            If RecommendedWordsListBox.Selected(i) Then
                SpellCheck.replacedWord = RecommendedWordsListBox.List(i)
                currentWordSelection.Text = SpellCheck.replacedWord
                Exit For
            End If
        Next i
        SpellCheck.wordReplaced = True
    End If
    Unload Me
End Sub

Private Sub IgnoreSpellCheckOptionButton_Click()
    CommandButton.Caption = "Ignore Spell Check"
End Sub

Private Sub ReplaceWordOptionButton_Click()
    CommandButton.Caption = "Replace word"
End Sub

Private Sub UserForm_Initialize()
    Dim c As Collection
    Set currentWordSelection = SpellCheck.currentWordsIterationSelection
    Set c = SpellCheck.getClosestWords(currentWordSelection.Range.Text)
    For Each itemm In c
        RecommendedWordsListBox.AddItem (itemm)
    Next itemm
    CurrentWordLabel.Caption = currentWordSelection.Text
End Sub

