Attribute VB_Name = "Module1"
Function PZ(Zone, Cat, AGrade, CtnType)
If Zone = "TB" Then
    PZ = "SHP"
Else
If Zone = "SH" And Cat = "Fragrance" Then
    PZ = "DG"
Else
'Bathroom Zoning
If Zone = "SH" And Cat = "Bathroom" And AGrade = "Yes" Then
    PZ = "SBP"
Else
If Zone = "SH" And Cat = "Bathroom" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SBS"
Else
If Zone = "SH" And Cat = "Bathroom" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SBQ"
Else

'Bed and Linen Zoning
If Zone = "SH" And Cat = "Bed Linen" And AGrade = "Yes" Then
    PZ = "SLP"
Else
If Zone = "SH" And Cat = "Bed Linen" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SLS"
Else
If Zone = "SH" And Cat = "Bed Linen" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SLQ"
Else

'Pillows and Quilts Zoning
If Zone = "SH" And Cat = "Pillows" Then
    PZ = "SHP"
Else
If Zone = "SH" And Cat = "Quilts" Then
    PZ = "SHP"
Else

''' Zoning everything else - Baby Shapes, Bed Accessories, Decorate
''' Tableware, Outerwear, Other, Merchandise Material & Sundry, Sleepwear

'Baby Shapes Zoning
If Zone = "SH" And Cat = "Baby Shapes" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Baby Shapes" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Baby Shapes" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

'Bed Accessories Zoning
If Zone = "SH" And Cat = "Bed Accessories" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Bed Accessories" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Bed Accessories" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

'Decorate Zoning
If Zone = "SH" And Cat = "Decorate" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Decorate" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Decorate" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

'Tablewear Zoning
If Zone = "SH" And Cat = "Tableware" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Tableware" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Tableware" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

'Outerwear Zoning
If Zone = "SH" And Cat = "Outerwear" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Outerwear" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Outerwear" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

'Other Zoning
If Zone = "SH" And Cat = "Other" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Other" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Other" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

'Merch Material Sundry Zoning
If Zone = "SH" And Cat = "Merchandise Material & Sundry" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Merchandise Material & Sundry" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Merchandise Material & Sundry" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

'Sleepwear Zoning
If Zone = "SH" And Cat = "Sleepwear" And AGrade = "Yes" Then
    PZ = "SOP"
Else
If Zone = "SH" And Cat = "Sleepwear" And AGrade = "No" And _
    (CtnType = "0" Or CtnType = "1" Or CtnType = "2" Or CtnType = "3" Or _
    CtnType = "4" Or CtnType = "5" Or CtnType = "6" Or CtnType = "7" Or _
    CtnType = "8" Or CtnType = "9" Or CtnType = "A" Or CtnType = "B" Or _
    CtnType = "C" Or CtnType = "D" Or CtnType = "E" Or CtnType = "F" Or _
    CtnType = "G" Or CtnType = "H" Or CtnType = "J" Or CtnType = "K" Or _
    CtnType = "L" Or CtnType = "M") Then
        PZ = "SOS"
Else
If Zone = "SH" And Cat = "Sleepwear" And AGrade = "No" And _
    (CtnType = "N" Or CtnType = "O" Or CtnType = "P" Or CtnType = "R" Or _
    CtnType = "S" Or CtnType = "T" Or CtnType = "U" Or CtnType = "V" Or _
    CtnType = "W" Or CtnType = "X" Or CtnType = "Y" Or CtnType = "Z") Then
        PZ = "SOQ"
Else

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
End If
End If
End If
End If
End If

End Function



