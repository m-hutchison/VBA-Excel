Attribute VB_Name = "Module2"
'Function that works out approx how many pallets there are
'Function: PQ(CtnType, Pallet Qty)

Function PQ(CtnType, PQty)
If CtnType = "1" Then
    PQ = Application.RoundUp(PQty / 205, 0)
Else
If CtnType = "2" Then
    PQ = Application.RoundUp(PQty / 144, 0)
Else
If CtnType = "3" Then
    PQ = Application.RoundUp(PQty / 120, 0)
Else
If CtnType = "A" Then
    PQ = Application.RoundUp(PQty / 96, 0)
Else
If CtnType = "B" Then
    PQ = Application.RoundUp(PQty / 72, 0)
Else
If CtnType = "C" Then
    PQ = Application.RoundUp(PQty / 65, 0)
Else
If CtnType = "D" Then
    PQ = Application.RoundUp(PQty / 60, 0)
Else
If CtnType = "E" Then
    PQ = Application.RoundUp(PQty / 48, 0)
Else
If CtnType = "F" Then
    PQ = Application.RoundUp(PQty / 40, 0)
Else
If CtnType = "G" Then
    PQ = Application.RoundUp(PQty / 36, 0)
Else
If CtnType = "H" Then
    PQ = Application.RoundUp(PQty / 32, 0)
Else
If CtnType = "J" Then
    PQ = Application.RoundUp(PQty / 30, 0)
Else
If CtnType = "K" Then
    PQ = Application.RoundUp(PQty / 28, 0)
Else
If CtnType = "L" Then
    PQ = Application.RoundUp(PQty / 24, 0)
Else
If CtnType = "M" Then
    PQ = Application.RoundUp(PQty / 20, 0)
Else
If CtnType = "N" Then
    PQ = Application.RoundUp(PQty / 18, 0)
Else
If CtnType = "O" Then
    PQ = Application.RoundUp(PQty / 16, 0)
Else
If CtnType = "P" Then
    PQ = Application.RoundUp(PQty / 14, 0)
Else
If CtnType = "R" Then
    PQ = Application.RoundUp(PQty / 12, 0)
Else
If CtnType = "S" Then
    PQ = Application.RoundUp(PQty / 10, 0)
Else
If CtnType = "T" Then
    PQ = Application.RoundUp(PQty / 8, 0)
Else
If CtnType = "U" Then
    PQ = Application.RoundUp(PQty / 6, 0)
Else
If CtnType = "V" Then
    PQ = Application.RoundUp(PQty / 5, 0)
Else
If CtnType = "W" Then
    PQ = Application.RoundUp(PQty / 4, 0)
Else
If CtnType = "X" Then
    PQ = Application.RoundUp(PQty / 3, 0)
Else
If CtnType = "Y" Then
    PQ = Application.RoundUp(PQty / 2, 0)
Else
If CtnType = "Z" Then
    PQ = Application.RoundUp(PQty / 1, 0)
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


