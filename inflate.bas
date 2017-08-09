Option Explicit
'Inflate (Deflate decompression) according https://tools.ietf.org/html/rfc1951

'TODO:
'      pReadBits: MORE TESTING NEEDED!!! 1. last byte full read?, 2. pReadBit vs bit by bit, 3...
'      check unused variables
'      test fixed codes
'      test uncompressed block
'      normalize error codes
'      separate static and dynamic block processing
'      convert all two power to zArrays

Private Declare Sub RtlMoveMemory Lib "Kernel32" (trg As Any, src As Any, ByVal Size As Long)

'z* names for performance improvement
Private zBitsMask(0 To 30) As Long '(&H01, &H03, &H07, &H0F, &H1F, &H3F, &H7F, &HFF, ...
Private z1BitMask(0 To 30) As Long '(&H01, &H02, &H04, &H08, &H10, &H20, &H40, &H80, ...
Private zhZero(1 To 15) '((0), (0,0), (0,0,0,0), ...)
'Tables
Private hlCLen_map(0 To 18) As Long, aLit_Bits(257 To 285) As Long, aLit_Add(257 To 285) As Long, aDist_Bits(0 To 29) As Long, aDist_Add(0 To 29) As Long
'Fixed alphabet
Private hcLitF, hcDistF

Private Const BUFFER_GROW_SIZE = 1048576 '1MB

'Returns decompressed buffer from Deflate compressed data
' - Buffer: compressed byte buffer
' - Position: buffer starting index (zero if not specified)
'             returns position after compressed data
Function Inflate(buffer() As Byte, Optional ByRef position As Long) As Byte()
    'Input/output buffer control
    Dim pBit As Long, oBuf() As Byte, oSize As Long, oByte As Long
    'Alphabets
    Dim hcLit(), hcDist(), hcCLen()
    'Alphabets upper bounds and counts
    Dim ubLen As Long, cLen As Long
    'Auxiliar variables
    Dim bFinal As Long, bType As Long, i As Long, lit As Long, dist As Long, length As Long, lens() As Integer
    
    pInit
    Do
        bFinal = pReadBits(buffer, position, pBit, 1)
        bType = pReadBits(buffer, position, pBit, 2)
        If bType = 0 Then
            'no compression
            'skip any remaining bits in current partially processed byte
            position = position - (pBit > 0)
            pBit = 0
            'read LEN
            length = buffer(position) + buffer(position + 1) * &H100&
            'check NLEN
            If (buffer(position + 2) + buffer(position + 3) * &H100& Xor &HFFFF&) <> length Then Err.Raise 57002, "Inflate", "Bad block data!"
            position = position + 4
            'Check input buffer
            If UBound(buffer) < position + length - 1 Then Err.Raise 57003, "Deflate.Inflate", "Not enough data!"
            If length Then 'Avoid unnecessary processing
                If oSize - oByte < length Then oSize = oByte + BUFFER_GROW_SIZE: ReDim Preserve oBuf(0 To oSize - 1)
                RtlMoveMemory oBuf(oByte), buffer(position), length
                position = position + length
                oByte = oByte + length
            End If
        ElseIf bType = 3 Then
            Err.Raise 57001, "Deflate.Inflate", "Invalid block!"
        Else 'BType=1 or BType=2
            If bType = 1 Then
                'compressed with fixed Huffman codes
                hcLit = hcLitF
                hcDist = hcDistF
            ElseIf bType = 2 Then
                'compressed with dynamic Huffman codes
                cLen = pReadBits(buffer, position, pBit, 5) + 257 'count HLIT
                ubLen = pReadBits(buffer, position, pBit, 5) + cLen 'upper bound HLIT+HDIST
                length = pReadBits(buffer, position, pBit, 4) + 3 'upper bound HCLEN
                ReDim lens(0 To 18)
                For i = 0 To length
                    lens(hlCLen_map(i)) = pReadBits(buffer, position, pBit, 3)
                Next
                hcCLen = pHTBuild(lens, 18, 7)
                ReDim lens(0 To ubLen)
                length = 0
                i = 0
                Do While i <= ubLen
                    lit = pHTDecode(buffer, position, pBit, hcCLen)
                    If lit <= 15 Then
                        lens(i) = lit
                        If lens(i) > length Then length = lens(i)
                        i = i + 1
                    ElseIf lit = 16 Then
                        For i = i To i + pReadBits(buffer, position, pBit, 2) + 2
                            lens(i) = lens(i - 1)
                        Next
                    Else
                        i = i + pReadBits(buffer, position, pBit, 3 - 4 * (lit = 18)) + 3 - 8 * (lit = 18)
                    End If
                Loop
                hcLit = pHTBuild(lens, cLen - 1, length)
                hcDist = pHTBuild(lens, ubLen - cLen, length, cLen)
            End If
            Do
                lit = pHTDecode(buffer, position, pBit, hcLit)
                If lit < 256& Then
                    If oByte >= oSize Then oSize = oSize + BUFFER_GROW_SIZE: ReDim Preserve oBuf(0 To oSize)
                    oBuf(oByte) = lit
                    oByte = oByte + 1
                ElseIf lit > 256& Then
                    length = aLit_Add(lit)
                    If aLit_Bits(lit) Then length = length + pReadBits(buffer, position, pBit, aLit_Bits(lit))
                    dist = pHTDecode(buffer, position, pBit, hcDist)
                    If aDist_Bits(dist) Then
                        dist = aDist_Add(dist) + pReadBits(buffer, position, pBit, aDist_Bits(dist))
                    Else
                        dist = aDist_Add(dist)
                    End If
                    If oByte + length > oSize Then oSize = oSize + BUFFER_GROW_SIZE: ReDim Preserve oBuf(0 To oSize)
                    For oByte = oByte To oByte + length - 1
                        oBuf(oByte) = oBuf(oByte - dist)
                    Next
                End If
            Loop Until lit = 256&
        End If
    Loop Until bFinal
    position = position - (pBit > 0) 'Skip remaining bits
    If oByte Then ReDim Preserve oBuf(0 To oByte - 1) 'Trim output buffer
    Inflate = oBuf
End Function

Private Function pHTDecode(buffer, pByte As Long, pBit As Long, htCodes) As Integer
    Dim code As Long, l As Long
    For l = 1 To 15 'Max len possible
        code = code * 2 - ((buffer(pByte) And z1BitMask(pBit)) <> 0)
        pBit = (pBit + 1) And 7
        If pBit = 0 Then pByte = pByte + 1
        If htCodes(l)(code) Then pHTDecode = htCodes(l)(code) - 1: Exit Function
    Next
    Err.Raise 57004, "Deflate.Inflate", "Invalid data!"
End Function

Private Function pHTBuild(htLen, ByVal max_code As Long, ByVal max_len As Long, Optional ByVal Index As Long)
    Dim htCode(), bl_count(0 To 15) As Long, code As Long, next_code(0 To 15) As Long, i As Long
    For i = 0 To max_code
        bl_count(htLen(i + Index)) = bl_count(htLen(i + Index)) + 1
    Next
    bl_count(0) = 0
    For i = 1 To max_len
        code = (code + bl_count(i - 1)) * 2
        next_code(i) = code
    Next
    htCode = zhZero
    For i = 0 To max_code
        If htLen(i + Index) Then
            htCode(htLen(i + Index))(next_code(htLen(i + Index))) = i + 1
            next_code(htLen(i + Index)) = next_code(htLen(i + Index)) + 1
        End If
    Next
    pHTBuild = htCode
End Function

'Max bits read at once: 13 (Dist extra bits). Huffman decoding uses inline reading!
Private Function pReadBits(buffer, pByte As Long, pBit As Long, ByVal Size As Long) As Long
    Dim ret As Long
    'Read first byte:
    ret = buffer(pByte) \ z1BitMask(pBit)
    pBit = pBit + Size
    If pBit < 8 Then pReadBits = ret And zBitsMask(Size): Exit Function
    'Not enough, read second byte:
    Dim bw As Long 'bits written
    bw = 8 - pBit + Size
    pBit = Size - bw
    ret = (zBitsMask(pBit) And buffer(pByte + 1)) * z1BitMask(bw) + ret
    If pBit < 8 Then pByte = pByte + 1: pReadBits = ret: Exit Function
    'Not enough, read third and last byte:
    bw = bw + 8
    pBit = pBit - 8
    pByte = pByte + 2
    pReadBits = (zBitsMask(pBit) And buffer(pByte)) * z1BitMask(bw) + ret
End Function

Private Sub pInit()
    Dim i As Long, a_b() As String, a_a() As String, a16() As Integer
    If hlCLen_map(0) = 16 Then Exit Sub 'Still init'd
    
    'Tables:
    'Literal/Length alphabet extra bits
    a_b = Split("0 0 0 0 0 0 0 0 1 1 1 1 2 2 2 2 3 3 3 3 4 4 4 4 5 5 5 5 0")
    a_a = Split("3 4 5 6 7 8 9 10 11 13 15 17 19 23 27 31 35 43 51 59 67 83 99 115 131 163 195 227 258")
    For i = 257 To 285
        aLit_Bits(i) = a_b(i - 257)
        aLit_Add(i) = a_a(i - 257)
    Next
    'Distance alphabet extra bits
    a_b = Split("0 0 0 0 1 1 2 2 3 3 4 4 5 5 6 6 7 7 8 8 9 9 10 10 11 11 12 12 13 13")
    a_a = Split("1 2 3 4 5 7 9 13 17 25 33 49 65 97 129 193 257 385 513 769 1025 1537 2049 3073 4097 6145 8193 12289 16385 24577")
    For i = 0 To 29
        aDist_Bits(i) = a_b(i)
        aDist_Add(i) = a_a(i)
    Next
    'Code length read order
    a_a = Split("16 17 18 0 8 7 9 6 10 5 11 4 12 3 13 2 14 1 15")
    For i = 0 To 18
        hlCLen_map(i) = a_a(i)
    Next
    
    'Empty Huffman tree
    For i = 1 To 15
        ReDim a16(0 To 2 ^ i - 1)
        zhZero(i) = a16
    Next
    
    'Static Huffman tree
    hcLitF = zhZero
    hcDistF = zhZero
    For i = 0 To 143: hcLitF(8)(48 + i) = i + 1: Next
    For i = 144 To 255: hcLitF(9)(400 + i - 144) = i + 1: Next
    For i = 256 To 279: hcLitF(7)(i - 256) = i + 1: Next
    For i = 280 To 287: hcLitF(8)(192 + i - 280) = i + 1: Next
    For i = 0 To 29: hcDistF(5)(i) = i + 1: Next
    
    'Bit masks
    For i = 0 To 30
        z1BitMask(i) = 2 ^ i
        zBitsMask(i) = 2 ^ i - 1
    Next

End Sub
