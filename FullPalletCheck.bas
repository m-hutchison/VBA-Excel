Attribute VB_Name = "Module5"
'Function that checks to see its a full pallet or not
'Function: FP(CtnType, Pallet Qty)

Function FP(CtnType, PQty)
If CtnType = "" And PQty = "" Then
    FP = ""
Else
If CtnType = "1" Then
    If PQty >= 205 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "2" Then
    If PQty >= 144 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "3" Then
    If PQty >= 120 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "A" Then
    If PQty >= 96 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "B" Then
    If PQty >= 72 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "C" Then
    If PQty >= 65 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "D" Then
    If PQty >= 60 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "E" Then
    If PQty >= 48 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "F" Then
    If PQty >= 40 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "G" Then
    If PQty >= 36 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "H" Then
    If PQty >= 32 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "J" Then
    If PQty >= 30 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "K" Then
    If PQty >= 28 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "L" Then
     If PQty >= 24 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "M" Then
    If PQty >= 24 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "N" Then
    If PQty >= 18 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "O" Then
    If PQty >= 16 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "P" Then
    If PQty >= 14 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "R" Then
    If PQty >= 12 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "S" Then
    If PQty >= 10 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "T" Then
    If PQty >= 8 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "U" Then
    If PQty >= 6 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "V" Then
    If PQty >= 5 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "W" Then
    If PQty >= 4 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "X" Then
    If PQty >= 3 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "Y" Then
    If PQty >= 2 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
Else
If CtnType = "Z" Then
    If PQty >= 1 Then
        FP = "Yes"
    Else
        FP = "No"
    End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Function


