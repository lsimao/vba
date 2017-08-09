'Testing module

Private Sub t_Inflate()
    Dim B() As Byte, s As String, d() As Byte, i As Long
    'If infgen = 0 And False Then
    '    infgen = FreeFile
    '    Open ThisWorkbook.Path & "\infgen.txt" For Append Shared As infgen
    'End If
    'showout = False
    B = buffer.FromFile(ThisWorkbook.path & "\pickletools.py.def")
    'b = buffer.FromDecimalStringArray("250 255 159 1 47 248 63 42 63 172 229 1 2 12 0 209 255 31 225") 'http://stackoverflow.com/questions/13924422/deflatestream-compress-decompress-inconsitency
    'Debug.Print Buffer.ToBitString(b)
    dbg.SetCounter
    For i = 0 To 9
        d = Inflate(B, LBound(B))
    Next
    Debug.Print "Inflate speed: "; i * (UBound(B) - LBound(B)) / dbg.GetCounter / 1024; "KB/s"
End Sub
