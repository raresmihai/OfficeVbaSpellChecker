VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DictionaryLanguageSelectionForm 
   Caption         =   "Dictionary Selection"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   OleObjectBlob   =   "DictionaryLanguageSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DictionaryLanguageSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ENGLISH As String = "English"
Private Const ROMANIAN As String = "Romanian"


Public selectedLanguage As String
Public currentWordSelection As Selection


Private Sub StartCheckCommandButton_Click()
    
    If EnglishOptionButton.Value = True Then
        selectedLanguage = ENGLISH
    End If
    If RomanianOptionButton.Value = True Then
        selectedLanguage = ROMANIAN
    End If
    Set currentWordSelection = Selection
    SpellCheck.SpellCheck language:=selectedLanguage, currentSelection:=currentWordSelection
    
    Unload Me
End Sub
