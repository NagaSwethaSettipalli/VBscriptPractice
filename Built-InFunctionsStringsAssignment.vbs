'Given Text: MyText = "The quick brown fox jumps over the lazy dog"
'find total no of words
'extract the word jump and dispay a message
'display a message with the reverse of word "quick"
'display a message by removing all white spaces from the variable MyText
'Find the length of MyText variable

Option Explicit

Dim MyText, arrWords, word, extractedWord, extractedWord1
Dim revWord , newText

MyText = "The quick brown fox jumps over the lazy dog"

'1. Find total no of words in MyText
arrWords = Split(MyText)
For Each word in arrWords
    WScript.Echo word
Next
 
'2. Extract word jump and display a message
extractedWord = Mid(MyText, 21, 4)
MsgBox " The word after extraction is : " & extractedWord 

'3. display a message with the reverse of word "quick"
extractedWord1 = Mid(MyText, 5, 5 )
MsgBox " The word after extraction is : " & extractedWord1
revWord = StrReverse(extractedWord1)
MsgBox " the reverse of extracted word is : " & revWord

'4.display a message by removing all white spaces from the variable MyText
newText = Replace (MyText, " ", "")
MsgBox " New text after removing white spaces is : " & newText

'5. Find the length of MyText variable
 MsgBox " The length of MyText variable is : " & Len(MyText)