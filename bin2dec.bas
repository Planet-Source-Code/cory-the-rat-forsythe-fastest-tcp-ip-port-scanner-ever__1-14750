Attribute VB_Name = "bin2dec"
Public Function Bin(BinaryValue As Variant) As Variant
    Dim Bit As Integer
    Dim Value As Integer
    Dim Counter As Integer


    For Counter = 1 To Len(BinaryValue)
        Bit = Mid(BinaryValue, Counter, 1)
        


        If Bit = 1 Then
            Value = Value + 2 ^ (Len(BinaryValue) - Counter)
        End If
    Next
    Bin = Value
End Function
        

