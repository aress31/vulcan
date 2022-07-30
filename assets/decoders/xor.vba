Public Function kUG(HoR As Variant)
    For i = 0 To UBound(HoR)
        HoR(i) = HoR(i) - CAESAR
    Next i
End Function