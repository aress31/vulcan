
Public Function kUG(ByRef srK() As Variant, kAm As String)
    For i = 0 To UBound(srK)
        j = i Mod Len(kAm)
        BpG = Asc(srK(i)) Xor Asc(Mid(kAm, j + 1, 1))
        srK(i) = Chr(BpG)
    Next i
End Function