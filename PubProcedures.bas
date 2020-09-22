Attribute VB_Name = "PubProcedures"
Public First
Public secondg
Public thirdg
Public fourthg
Public Function dec2bin(mynum As Variant) As String
Dim loopcounter As Integer
If mynum >= 2 ^ 31 Then
dec2bin = "Too big"
Exit Function
End If
Do
If (mynum And 2 ^ loopcounter) = 2 ^ loopcounter Then
dec2bin = "1" & dec2bin
Else
dec2bin = "0" & dec2bin
End If
loopcounter = loopcounter + 1
Loop Until 2 ^ loopcounter > mynum
End Function


Public Sub converb()
Dim firstbin As Integer
firstbin = First
Dim secondBin As Integer
secondBin = secondg
Dim thirdBin As Integer
thirdBin = thirdg
Dim fourthBin As Integer
fourthBin = fourthg
Dim firstbinary As Integer
firstbinary = dec2bin(firstbin)
MsgBox firstbinary
End Sub
