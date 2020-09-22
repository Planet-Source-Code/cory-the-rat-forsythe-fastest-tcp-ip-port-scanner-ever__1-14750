Attribute VB_Name = "Module1"
''load final outputs into string buffer
Public firstbinary As String
Public secondbinary As String
Public thirdbinary As String
Public fourthbinary As String

Public Sub dec2bin(firstnum As String, secondnum As String, thirdnum As String, fourthnum As String)
If firstnum = "0" Then firstbinary = "0"
If firstnum = "1" Then firstbinary = "1"
If secondnum = "0" Then secondbinary = "0"
If secondnum = "1" Then secondbinary = "1"
End Sub
