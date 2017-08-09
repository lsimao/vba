'Provides ZLib buffer handing

Public Function ZLib_Decompress(buffer() As Byte, Optional ByRef position As Long) As Byte()
    Dim ret() As Byte, cs() As Byte
    If buffer(position) <> &H78 Then Err.Raise 57000, "ZLib_Decompress", "Unknown compression method!"
    If (buffer(position) * 256& + buffer(position + 1)) Mod 31 <> 0 Then Err.Raise 57002, "ZLib.Decompress", "Checksum failed!"
    If buffer(position + 1) And &H20 Then position = position + 4 'DICT unexpected, but ignored!
    position = position + 2
    ret = Deflate.Inflate(buffer, position)
    cs = Adler32(ret)
    If buffer(position) <> cs(0) Or buffer(position + 1) <> cs(1) Or buffer(position + 2) <> cs(2) Or buffer(position + 3) <> cs(3) Then Err.Raise 57002, "ZLib.Decompress", "Checksum failed!"
    position = position + 4
    ZLib_Decompress = ret
End Function

Public Function Adler32(buffer() As Byte, Optional ByVal position As Long, Optional ByVal Size As Long = -1) As Byte()
    Dim s1 As Long, s2 As Long, ub As Long, ret(0 To 3) As Byte
    ub = IIf(Size < 0, UBound(buffer), position + Size - 1)
    s1 = 1
    For position = position To ub
        s1 = (s1 + buffer(position)) Mod 65521
        s2 = (s2 + s1) Mod 65521
    Next
    ret(0) = s2 \ &H100 And &HFF
    ret(1) = s2 And &HFF
    ret(2) = s1 \ &H100 And &HFF
    ret(3) = s1 And &HFF
    Adler32 = ret
End Function
