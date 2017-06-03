Attribute VB_Name = "SpellCheck"

Public dict As New Scripting.Dictionary
Public objExcel As Object
Public exWb As Excel.Workbook
Public dictWorksheet As Excel.Worksheet

Public currentWordsIterationSelection As Selection
Public currentSelectedLanguage As String

Public wordReplaced As Boolean
Public replacedWord As String


Public Sub Main()
    SpellCheckingForm.Show
End Sub

Public Sub SpellCheck(ByVal language As String, ByVal currentSelection As Selection)
    OpenDictionary (language)
    SetDictionary
    currentSelectedLanguage = language
    IterateWords selectedRange:=currentSelection.Range
'Dim c As Collection
'Set c = getClosestWords("NouNo")
'MsgBox (c.Count)

'Dim r As Range
'Set r = Selection.Range
    exWb.Save
    exWb.Close
End Sub

Public Function getAllWords(ByVal language As String) As Collection
    Dim allWords As New Collection
    
    OpenDictionary (language)
    SetDictionary
    exWb.Close
    
    For Each Key In dict.Keys
        allWords.Add (Key)
    Next Key
    
    Set getAllWords = allWords
End Function

Public Sub addWordToDictionary(word As String, language As String)
    word = LCase(word)
    OpenDictionary (language)
    SetDictionary
    If Not dict.Exists(word) Then
        Dim LastRow As Long
        LastRow = dictWorksheet.Range("A" & dictWorksheet.Rows.Count).End(xlUp).Row + 1
        dictWorksheet.Cells(LastRow, 1).Value = word
        dictWorksheet.UsedRange.Sort Key1:=Worksheets(currentSelectedLanguage).Range("A1"), Order1:=xlAscending
        SetDictionary
        exWb.Save
    End If
    exWb.Close
End Sub

Public Sub removeWordFromDictionary(word As String, language As String)
    word = LCase(word)
    OpenDictionary (language)
    SetDictionary
    If dict.Exists(word) Then
        dict.Remove word
        MsgBox "deleted!!"
        For Each Cell In dictWorksheet.Cells
            If Cell.Value = word Then
                Cell.Delete
                Exit For
            End If
        Next Cell
        dictWorksheet.UsedRange.Sort Key1:=Worksheets(currentSelectedLanguage).Range("A1"), Order1:=xlAscending
        SetDictionary
        exWb.Save
    End If
    exWb.Close
End Sub

Public Sub editWordFromDictionary(oldWord As String, newWord As String, language As String)
    oldWord = LCase(oldWord)
    OpenDictionary (language)
    SetDictionary
    If dict.Exists(oldWord) Then
        For Each Cell In dictWorksheet.Cells
            If Cell.Value = oldWord Then
                Cell.Value = newWord
                Exit For
            End If
        Next Cell
        dictWorksheet.UsedRange.Sort Key1:=Worksheets(currentSelectedLanguage).Range("A1"), Order1:=xlAscending
        SetDictionary
        exWb.Save
    End If
    exWb.Close
End Sub

Public Function getClosestWords(ByVal word As String) As Collection
    Set getClosestWords = New Collection
    Dim closestWords As New Collection
    Dim distance As Long
    Dim min As Long
    min = Len(word)
    'word = LCase(word)
    
    'Calculate the minimum distance and add those words in the collection
    For Each Key In dict.Keys
        distance = Levenshtein(word, Key)
        If distance < min Then
            min = distance
            Set closestWords = New Collection
            closestWords.Add (Key)
        Else
            If distance = min Then
                closestWords.Add (Key)
            End If
        End If
    Next Key
    
    'Get only 5 closest words (that have the minimum distance)
    If closestWords.Count > 5 Then
        Dim addedWords As Integer
        addedWords = 0
        Dim ob As Variant
        For Each ob In closestWords
            getClosestWords.Add (ob)
            addedWords = addedWords + 1
            If addedWords = 5 Then
                Exit Function
            End If
        Next ob
    Else
        Set getClosestWords = closestWords
    End If
    
End Function

Public Sub addToDictionary(word As String)
    word = LCase(word)
    If Not dict.Exists(word) Then
        Dim LastRow As Long
        LastRow = dictWorksheet.Range("A" & dictWorksheet.Rows.Count).End(xlUp).Row + 1
        dictWorksheet.Cells(LastRow, 1).Value = word
        dictWorksheet.UsedRange.Sort Key1:=Worksheets(currentSelectedLanguage).Range("A1"), Order1:=xlAscending
        SetDictionary
    End If
End Sub

'Iterate through the selection, word by word, by selecting, underlying, and opening a form for words that are not in the dict
Private Sub IterateWords(ByVal selectedRange As Range)
    wordReplaced = False
    Dim currentChar As String, j As Integer, word As String
    For i = 1 To Len(selectedRange)
        j = 0
        currentChar = selectedRange.Characters(i)
        On Error GoTo GetOut
        word = ""
        While IsLetter(currentChar) Or IsNumeric(currentChar)
            word = word & currentChar
            j = j + 1
            currentChar = selectedRange.Characters(i + j)
        Wend
        If j > 0 Then 'We have a word
            If isBadSpelling(word) Then 'Word not in the dictionary
                selectedRange.Characters(i).Select 'Select the first letter of the word
                Selection.MoveEnd Count:=(j - 1) 'Select the whole word
                Selection.Font.Underline = wdUnderlineWords 'Underline the word
                Set currentWordsIterationSelection = Selection
                OptionsForSpellCheckingForm.Show 'Show the form with options(Ignore,Replace,AddWordToDict)
                If wordReplaced = True Then
                    wordReplaced = False
                    j = Len(replacedWord) - 1
                End If
            End If
        End If
        i = i + j 'Move to the end of the word
        selectedRange.Characters(i).Select
    Next i
    
GetOut:
    MsgBox "Spell Checking Completed!"
End Sub

Private Sub OpenDictionary(dictName As String)
    Dim path As String
    path = "D:\Facultate\2016-2017\Semestrul II\MS-OFFICE\Proiect\Dictionary.xls"
    'Set objExcel = CreateObject("Excel.Application")
    Set exWb = Workbooks.Open(path)
    Set dictWorksheet = exWb.Worksheets(dictName)
End Sub

Private Sub SetDictionary()
    Set dict = New Scripting.Dictionary
    For Each cel In dictWorksheet.UsedRange.Columns("A").Cells
        If Not dict.Exists(cel.Text) Then
             dict.Add cel.Text, cel.Row
        End If
    Next cel
End Sub

Private Function isBadSpelling(word As String) As Boolean
    If dict.Exists(LCase(word)) Then
        isBadSpelling = False
    Else
        isBadSpelling = True
    End If
End Function

Private Function isWord(word As String)
    Dim currentChar As String
    For i = 1 To Len(word)
        currentChar = Mid(word, i, 1)
        If IsLetter(currentChar) = False And IsNumeric(currentChar) = False Then
            isWord = False
            Exit Function
        End If
    Next i
    isWord = True
End Function

Private Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function


Private Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

    Dim i As Long, j As Long
    Dim string1_length As Long
    Dim string2_length As Long
    Dim distance() As Long
    
    string1_length = Len(string1)
    string2_length = Len(string2)
    ReDim distance(string1_length, string2_length)
    
    For i = 0 To string1_length
        distance(i, 0) = i
    Next
    
    For j = 0 To string2_length
        distance(0, j) = j
    Next
    
    For i = 1 To string1_length
        For j = 1 To string2_length
            If Asc(Mid$(string1, i, 1)) = Asc(Mid$(string2, j, 1)) Then
                distance(i, j) = distance(i - 1, j - 1)
            Else
                distance(i, j) = min3 _
                (distance(i - 1, j) + 1, _
                 distance(i, j - 1) + 1, _
                 distance(i - 1, j - 1) + 1)
            End If
        Next
    Next
    
    Levenshtein = distance(string1_length, string2_length)

End Function

Private Function min3(ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    min3 = v1
    If min3 > v2 Then
        min3 = v2
    End If
    If min3 > v3 Then
        min3 = v3
    End If
End Function

