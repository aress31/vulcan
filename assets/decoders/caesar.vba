
Public Function kUG(ByRef srK() As Variant, kAm As Integer)
    For i = 0 To UBound(srK)
        BpG = Asc(srK(i)) - kAm
        
        If BpG < 0 Then
            BpG = (Asc(srK(i)) - kAm) + 256
        End If
 
        srK(i) = Chr(BpG)
    Next i
End Function