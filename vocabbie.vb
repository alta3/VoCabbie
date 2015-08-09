Sub readvocab(skipword As Integer)
    
    Dim VocabEntry() As String
    Dim iFileNum As Integer
    Dim i As Integer
    Dim w As Integer
    Dim pattern As String
    Dim vocabWord() As String
    Dim Chapter() As String
    Dim Definition() As String
    Dim vocabRegexSplit As Object
    Dim strInput As String
    Dim b_found As Boolean
    Dim chapternum As String
    Dim wordcount As Integer
    Dim myword As Integer

' Location of vocab file:
    VocabFileName = "C:\Users\sfeeser\Cloud\OpenStack\Vocab.txt"
    
' does vocab file exist?
    If Len(Dir$(VocabFileName)) = 0 Then
        MsgBox "Vocab.txt file does not exist!"
        Exit Sub
    End If

 ' -------- Set up REGEX -------- '
    Set vocabRegexSplit = New RegExp 'Create the RegExp object
    With vocabRegexSplit
        .Global = True
        .pattern = "([a-zA-Z0-9]*)\s+\[(\d*)\]\s+(.*)"
        .IgnoreCase = False         ' False is default
    End With

' Initialize file handle number and open file
    iFileNum = FreeFile()
    Open VocabFileName For Input As iFileNum
    
' -------- Load vocab text file into an array-------- '
    i = 0
    Do Until EOF(iFileNum)
        ReDim Preserve VocabEntry(i)
        Line Input #iFileNum, VocabEntry(i)
        i = i + 1
    Loop
    Close iFileNum
 
' -------- Load chapter number-------- '
    chapternum = chapternumber.Text
 ' -------- Parse array to contain only this chapter's vocabulary words --------
    w = 0
    For i = 0 To UBound(VocabEntry)
      strInput = VocabEntry(i)
       If chapternum = vocabRegexSplit.Replace(strInput, "$2") Then
        ReDim Preserve vocabWord(w)
        ReDim Preserve Chapter(w)
        ReDim Preserve Definition(w)
        vocabWord(w) = vocabRegexSplit.Replace(strInput, "$1")
        Chapter(w) = vocabRegexSplit.Replace(strInput, "$2")
        Definition(w) = vocabRegexSplit.Replace(strInput, "$3")
        w = w + 1
      End If
    Next
    
' -------- Set indexes and error check  --------
    myword = wordcounter.Value - 1
    If w = 0 Then   ' No matches!
        ReDim Preserve vocabWord(w)
        ReDim Preserve Chapter(w)
        ReDim Preserve Definition(w)
        vocabWord(w) = "---"
        Chapter(w) = chapternum
        Definition(w) = "Please add at least ONE WORD"
        myword = 0
    ElseIf UBound(vocabWord) = 0 Then
      myword = 0
    ElseIf myword = 0 Then
      Select Case skipword
      Case Is = 1
        myword = 1
      Case Is = -1
        myword = UBound(vocabWord)
      Case Is = 0
        myword = 0
      End Select
    ElseIf (myword > 0) And (myword < UBound(vocabWord)) Then
      Select Case skipword
      Case Is = 1
        myword = myword + 1
      Case Is = -1
        myword = myword - 1
      Case Is = 0
        myword = 0
      End Select
    ElseIf (myword > 0) And (myword >= UBound(vocabWord)) Then
      Select Case skipword
      Case Is = 1
        myword = 0
      Case Is = -1
        myword = myword - 1
      Case Is = 0
        myword = 0
      End Select
    End If
    
 ' ------------- Populate screen  -----------
    wordcounter.Value = (myword + 1)
    wordquan.Value = (UBound(vocabWord) + 1)
    If myword > UBound(vocabWord) Then
      myword = UBound(vocabWord)
    End If
    vocabWordbox.Value = vocabWord(myword)
    vocabbox.Value = Definition(myword)
    Set vocabRegexSplit = Nothing
End Sub
Private Sub chapternumber_Change()
    ' Call readvocab(0)
End Sub
Private Sub wordcounter_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call readvocab(0)
End Sub
Private Sub wordquan_Change()

End Sub
Private Sub wordsurfer_SpinUp()
    Call readvocab(1)
End Sub
Private Sub wordsurfer_SpinDown()
    Call readvocab(-1)
End Sub
