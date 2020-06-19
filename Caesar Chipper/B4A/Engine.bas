B4A=true
Group=Default Group
ModulesStructureVersion=1
Type=Class
Version=7.3
@EndOfDesignText@
Sub Class_Globals
	Dim hasil As String 
	Dim abjad(1 To 26) As String 
End Sub


Public Sub getIdxFromStr(str As String, geser As Int, enkrip As Boolean)
Dim idx As Int
Dim hasil As String
Select Case str
Case "a"
idx = 1
Case "b"
idx = 2
Case "c"
idx = 3
Case "d"
idx = 4
Case "e"
idx = 5
Case "f"
idx = 6
Case "g"
idx = 7
Case "h"
idx = 8
Case "i"
idx = 9
Case "j"
idx = 10
Case "k"
idx = 11
Case "l"
idx = 12
Case "m"
idx = 13
Case "n"
idx = 14
Case "o"
idx = 15
Case "p"
idx = 16
Case "q"
idx = 17
Case "r"
idx = 18
Case "s"
idx = 19
Case "t"
idx = 20
Case "u"
idx = 21
Case "v"
idx = 22
Case "w"
idx = 23
Case "x"
idx = 24
Case "y"
idx = 25
Case "z"
idx = 26
Case " "
idx = 0
End Select


abjad(1) = "a"
abjad(2) = "b"
abjad(3) = "c"
abjad(4) = "d"
abjad(5) = "e"
abjad(6) = "f"
abjad(7) = "g"
abjad(8) = "h"
abjad(9) = "i"
abjad(10) = "j"
abjad(11) = "k"
abjad(12) = "l"
abjad(13) = "m"
abjad(14) = "n"
abjad(15) = "o"
abjad(16) = "p"
abjad(17) = "q"
abjad(18) = "r"
abjad(19) = "s"
abjad(20) = "t"
abjad(21) = "u"
abjad(22) = "v"
abjad(23) = "w"
abjad(24) = "x"
abjad(25) = "y"
abjad(26) = "z"

If enkrip = True Then
    If idx + geser > 26 Then
        hasil = abjad((idx - 26) + geser)
    Else
        hasil = abjad(idx + geser)
            If idx = 0 Then
                hasil = " "
            End If
    End If
Else

If idx - geser < 1 Then
    hasil = abjad((idx + 26) - geser)
        If idx = 0 Then
         hasil = " "
        End If
Else
    hasil = abjad(idx - geser)
End If
End If

Return hasil

'Form1.txthasil.Text = Form1.txthasil.Text & hasil
End Sub

Public Sub getCaesarText(str As String, geser As Int) As String

For i = 1 To str.Length 
    getIdxFromStr (str.SubString2(i, 1), geser, True)
Next
Return hasil
End Sub

Public Sub DecryptCaesarText(str As String, geser As Int) As String
Dim hasil As String
	For i = 1 To str.Length
		getIdxFromStr (str.SubString2(i, 1), geser, False)
	Next
Return hasil
End Sub


'Initializes the object. You can add parameters to this method if needed.
Public Sub Initialize

End Sub