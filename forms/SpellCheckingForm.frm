VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SpellCheckingForm 
   Caption         =   "Spell Checking"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205.001
   OleObjectBlob   =   "SpellCheckingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SpellCheckingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public documentSelection As Selection
Public dictionaryLanguage As String
Private Const ENGLISH As String = "English"
Private Const ROMANIAN As String = "Romanian"

Private Sub CurrentSelectionCommandButton_Click()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Current selection is invalid or no selection has been made. Please select a portion of the text and run the Macro again.", vbOKOnly
        Unload Me
        Exit Sub
    End If
    Set documentSelection = Selection
    DictionaryLanguageSelectionForm.Show
    Unload Me
End Sub

Private Sub EntireDocumentCommandButton_Click()
    Selection.WholeStory
    Set documentSelection = Selection
    DictionaryLanguageSelectionForm.Show
    Unload Me
End Sub

Private Sub EnglishDictionaryCommandButton_Click()
    dictionaryLanguage = ENGLISH
    AddModifyDeleteWordForm.Show
    Unload Me
End Sub

Private Sub RomanianDictionaryCommandButton_Click()
    dictionaryLanguage = ROMANIAN
    AddModifyDeleteWordForm.Show
    Unload Me
End Sub
