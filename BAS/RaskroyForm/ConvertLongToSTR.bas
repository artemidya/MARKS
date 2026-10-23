Private Function ConvertLongToSTR(ByRef s As Long) As String
    Dim aa As Byte, bb As Byte, cc As Byte, dd As Byte
    aa = Int(s / 16777216)
    bb = Int((s - CDbl(aa) * 16777216) / 65536)
    cc = Int((s - CDbl(aa) * 16777216 - CDbl(bb) * 65536) / 256)
    dd = Int((s - CDbl(aa) * 16777216 - CDbl(bb) * 65536 - CDbl(cc) * 256))
    ConvertLongToSTR = Chr(aa) & Chr(bb) & Chr(cc) & Chr(dd)
End Function
