Module Module1
    Sub Main() ' stub for testing...
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        'Console.WriteLine(Encode("01234567"))
        'Console.WriteLine(Encode("0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789"))
        'Console.WriteLine(Encode("AC-42"))
        'Console.WriteLine(Encode("The quick brown fox jumped over the lazy dog. The quick brown fox jumped over the lazy dog. The quick brown fox jumped over the lazy dog.The quick brown fox jumped over the lazy dog. The quick brown fox jumped over the lazy dog. The quick brown fox jumped over the lazy dog. The quick brown fox jumped over the lazy dog. The quick brown fox jumped over the lazy dog.The quick brown fox jumped over the lazy dog. The quick brown fox jumped over the lazy dog."))
        Console.WriteLine(Encode("THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG. THE QUICK BROWN FOX JUMPED OVER THE LAZY DOG.", 1, 1))
        'Console.WriteLine(Encode("2901234123573"))
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine("Press any key to continue...")
        Console.ReadLine()
    End Sub

    Public Function ImageExists(ByRef URL As String) As Boolean
        'Create a web request
        Dim m_Req As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create(URL)

        'Get a web response
        Try
            Dim HttpWResp As System.Net.HttpWebResponse = CType(m_Req.GetResponse(), System.Net.HttpWebResponse)
            If HttpWResp.StatusCode = System.Net.HttpStatusCode.OK Then
                Return True
            Else
                Return False
            End If
        Catch ex As System.Net.WebException
            If ex.Status = System.Net.WebExceptionStatus.ProtocolError Then
                Return False
            End If
        End Try
        Return Nothing
    End Function

    Public Function Encode(ByVal content As [String], ByVal qrcodeVersion As Byte, ByVal ErrorCorrectionLevel As Integer) As [String]
        Return Encode(content, qrcodeVersion, ErrorCorrectionLevel, System.Text.Encoding.ASCII)
    End Function

    Public Function Encode(ByVal content As [String], ByVal qrcodeVersion As Byte, ByVal ErrorCorrectionLevel As Integer, ByVal encoding As System.Text.Encoding) As [String]
        Dim qr As [String] = ""
        Dim matrix As Boolean()() = calQrcode(encoding.GetBytes(content), qrcodeVersion, ErrorCorrectionLevel)
        For i As Integer = 0 To matrix.Length \ 2 - 1
            qr += Chr(0)
            For j As Integer = 0 To matrix.Length - 1
                If (matrix(j)(i * 2) And matrix(j)(i * 2 + 1)) Then
                    qr += "█"
                ElseIf matrix(j)(i * 2) Then
                    qr += "▀"
                ElseIf matrix(j)(i * 2 + 1) Then
                    qr += "▄"
                Else
                    qr += " "
                End If
            Next
            qr += Chr(0) & vbLf
        Next
        qr += Chr(0)
        For j As Integer = 0 To matrix.Length - 1
            If (matrix(j)(matrix.Length - 1) And 1) <> 0 Then
                qr += "▀"
            Else
                qr += " "
            End If
        Next
        qr += Chr(0)
        Return qr
    End Function

    ' ErrorCorrectionLevel index {0=>M, 1=>L, 2=>H, 3=>Q}
    ' qrcodeVersion: minimum qr code version required, automatically increased if required by data
    Private Function calQrcode(ByVal qrcodeRawData As Byte(), ByVal qrcodeVersion As Byte, ByVal ErrorCorrectionLevel As Integer) As Boolean()()
        Dim DataLength As UInteger = qrcodeRawData.Length

        'Enqueue the data for sequential access
        Dim DataQueue As New Collections.Generic.Queue(Of Byte)(qrcodeRawData)

        If DataLength <= 0 Then
            Dim ret As Boolean()() = New Boolean()() {New Boolean() {False}}
            Return ret
        End If

        ' Determine encoding mode based on received character set - use only single mode
        Dim qrcodeEncodeMode As Byte ' 0=Numeric, 1=Alphanumeric, 2=8-bit Byte Encoding
        Dim AlphaNumericSpecialChars As Char() = New Char() {" ", "$", "%", "*", "+", "-", ".", "/", ":"}
        For i As Integer = 0 To DataLength - 1
            If (qrcodeRawData(i) >= Asc("0") AndAlso qrcodeRawData(i) <= Asc("9")) Then Continue For ' optimise for Numeric mode, then Alphanumeric
            qrcodeEncodeMode = 1
            If (qrcodeRawData(i) < Asc("A") OrElse qrcodeRawData(i) > Asc("Z")) AndAlso (Array.IndexOf(AlphaNumericSpecialChars, Chr(qrcodeRawData(i))) = 0) Then
                qrcodeEncodeMode = 2
                Exit For
            End If
        Next

        Static characterCountBits As Byte()() = New Byte(2)() {New Byte(2) {10, 12, 14}, New Byte(2) {9, 11, 13}, New Byte(2) {8, 16, 16}}

        Dim DataBits As Integer
        Dim DigitGroups As Integer
        Dim RemainingDigits As Integer

        ' TODO: repeated code, use array constant to store encode mode parameters and condense this logic
        If qrcodeEncodeMode = 0 Then
            DigitGroups = DataLength \ 3
            RemainingDigits = DataLength Mod 3
            DataBits = 4 + 10 * DigitGroups + 7 * (RemainingDigits >> 1) + 4 * (RemainingDigits And 1)
        ElseIf qrcodeEncodeMode = 1 Then
            DigitGroups = DataLength \ 2
            RemainingDigits = DataLength Mod 2
            DataBits = 4 + 11 * DigitGroups + 6 * (RemainingDigits And 1)
        Else
            DigitGroups = DataLength
            RemainingDigits = 0
            DataBits = 4 + 8 * DigitGroups
        End If

        Dim ModulesAcross As Byte
        Dim AlignmentPatternsAcross As Byte
        Dim AlignmentPatternCount As Byte
        Dim DataECCodewords As Integer
        Dim ECCodewords As Integer

        Static Blocks As Integer()() = New Integer()() { _
            New Integer() {0, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49}, _
            New Integer() {0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25}, _
            New Integer() {0, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81}, _
            New Integer() {0, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68} _
        }

        Static ECratio As Double() = New Double() {0.1855, 0.1, 0.32, 0.2699}

        Static Exceptions As Integer()() = New Integer()() { _
            New Integer() {0}, _
            New Integer() {0, 7, 10, 15}, _
            New Integer() {0, 17}, _
            New Integer() {0, 13, 22} _
        }

        Dim usedDataCodewords As Integer

        'calculate the required version
        Dim versionGroup As Byte
        For qrcodeVersion = qrcodeVersion To 40
            ModulesAcross = 17 + 4 * qrcodeVersion
            If qrcodeVersion > 1 Then
                AlignmentPatternsAcross = (qrcodeVersion \ 7) + 2
                AlignmentPatternCount = AlignmentPatternsAcross ^ 2 - 3
            End If

            DataECCodewords = ModulesAcross ^ 2 - 191 ' modules less 3 8x8 Position Detection Patterns and 31-bit Format Information, plus 32-bit Timing Pattern overlap with Position Detection Patterns
            DataECCodewords -= AlignmentPatternCount * 25 ' Alignment Patterns,
            DataECCodewords -= 2 * ModulesAcross ' Timing Patterns,
            If AlignmentPatternsAcross > 2 Then
                DataECCodewords += (AlignmentPatternsAcross - 2) * 10 ' Timing Patterns overlap Alignment Patterns
            End If
            If qrcodeVersion > 6 Then
                DataECCodewords -= 36 ' Version Information
            End If

            DataECCodewords \= 8 ' Convert to Data/EC Blocks, discarding Remainder Bits
            If qrcodeVersion >= Exceptions(ErrorCorrectionLevel).Length Then
                ECCodewords = Math.Round(DataECCodewords / Blocks(ErrorCorrectionLevel)(qrcodeVersion) * ECratio(ErrorCorrectionLevel)) * 2
            Else
                ECCodewords = Exceptions(ErrorCorrectionLevel)(qrcodeVersion)
            End If

            If qrcodeVersion = 10 Or qrcodeVersion = 27 Then
                versionGroup += 1
            End If
            usedDataCodewords = Math.Ceiling((DataBits + characterCountBits(qrcodeEncodeMode)(versionGroup)) / 8)
            If usedDataCodewords + ECCodewords * Blocks(ErrorCorrectionLevel)(qrcodeVersion) <= DataECCodewords Then
                Exit For
            End If
        Next

        ' Log and Antilog tables for bytewise modulo 285 arithmetic
        Dim Log As Integer() = New Integer(256) {}
        Dim ALog As Integer() = New Integer(256) {}
        ALog(0) = 1
        For i As Integer = 1 To 255
            ALog(i) = ALog(i - 1) * 2
            If ALog(i) > 255 Then ALog(i) = ALog(i) Xor 285
            Log(ALog(i)) = i
        Next

        Dim G As Byte() = New Byte(ECCodewords) {}
        G(0) = 1
        For i As Integer = 0 To ECCodewords - 1
            For j As Integer = i To 0 Step -1
                G(j + 1) = G(j + 1) Xor G(j)
                G(j) = ALog((Log(G(j)) + i) Mod 255)
            Next
        Next

        Dim NumBlocks1 As Integer = Blocks(ErrorCorrectionLevel)(qrcodeVersion) - DataECCodewords Mod Blocks(ErrorCorrectionLevel)(qrcodeVersion)
        Dim NumBlocks2 As Integer = DataECCodewords Mod Blocks(ErrorCorrectionLevel)(qrcodeVersion)
        Dim BlockSize As Integer = DataECCodewords \ Blocks(ErrorCorrectionLevel)(qrcodeVersion)
        Dim DataCodewords As Integer = DataECCodewords \ Blocks(ErrorCorrectionLevel)(qrcodeVersion) - ECCodewords

        Dim BitOffset As UInteger
        Dim dataBitstream As Byte() = New Byte(DataECCodewords - ECCodewords - 1) {} ' give this a name - DataCodewords

        writeBits(dataBitstream, BitOffset, 1 << qrcodeEncodeMode, 4) ' mode indicator
        writeBits(dataBitstream, BitOffset, DataLength, characterCountBits(qrcodeEncodeMode)(versionGroup)) ' character count bits
        ' TODO: repeated code.
        If qrcodeEncodeMode = 0 Then
            For i As Integer = 0 To DigitGroups - 1 ' write 3 digits as 10-bit
                Dim DataVal As UInteger = (DataQueue.Dequeue() - Asc("0")) * 100 + (DataQueue.Dequeue() - Asc("0")) * 10 + (DataQueue.Dequeue() - Asc("0"))
                writeBits(dataBitstream, BitOffset, DataVal, 10)
            Next
            If RemainingDigits = 2 Then ' write remainder as 4 or 7 bits
                Dim DataVal As UInteger = (DataQueue.Dequeue() - Asc("0")) * 10 + (DataQueue.Dequeue() - Asc("0"))
                writeBits(dataBitstream, BitOffset, DataVal, 7)
            ElseIf RemainingDigits = 1 Then
                Dim DataVal As UInteger = DataQueue.Dequeue() - Asc("0")
                writeBits(dataBitstream, BitOffset, DataVal, 4)
            End If
        ElseIf qrcodeEncodeMode = 1 Then ' Alphanumeric mode
            For i As Integer = 0 To DigitGroups - 1 ' write 2 characters as 11-bit
                Dim DataVal As UInteger() = New UInteger(1) {}
                For j As Integer = 0 To 1
                    If (DataQueue.Peek() >= Asc("0") AndAlso DataQueue.Peek() <= Asc("9")) Then
                        DataVal(j) = DataQueue.Dequeue() - Asc("0")
                    ElseIf DataQueue.Peek() >= Asc("A") AndAlso DataQueue.Peek() <= Asc("Z") Then
                        DataVal(j) = 10 + DataQueue.Dequeue() - Asc("A")
                    Else
                        DataVal(j) = 36 + Array.IndexOf(AlphaNumericSpecialChars, Chr(DataQueue.Dequeue()))
                    End If
                Next
                writeBits(dataBitstream, BitOffset, 45 * DataVal(0) + DataVal(1), 11)
            Next
            If RemainingDigits = 1 Then ' write remaining character as 6-bit
                Dim DataVal As UInteger
                If (DataQueue.Peek() >= Asc("0") AndAlso DataQueue.Peek() <= Asc("9")) Then
                    DataVal = DataQueue.Dequeue() - Asc("0")
                ElseIf DataQueue.Peek() >= Asc("A") AndAlso DataQueue.Peek() <= Asc("Z") Then
                    DataVal = 10 + DataQueue.Dequeue() - Asc("A")
                Else
                    DataVal = 36 + Array.IndexOf(AlphaNumericSpecialChars, Chr(DataQueue.Dequeue()))
                End If
                writeBits(dataBitstream, BitOffset, DataVal, 6)
            End If
        Else ' 8-bit Byte mode
            For i As Integer = 0 To qrcodeRawData.Length - 1
                Dim DataVal As UInteger = DataQueue.Dequeue()
                writeBits(dataBitstream, BitOffset, DataVal, 8)
            Next
        End If

        ' add terminator and padding bits
        Static dataFill As Byte() = New Byte() {236, 17}
        If DataECCodewords - ECCodewords - usedDataCodewords > 1 Then
            writeBits(dataBitstream, BitOffset, 0, 4)
            Dim Offset As Integer = Math.Ceiling(BitOffset / 8)
            For i As Integer = 0 To DataECCodewords - ECCodewords - Offset - 1
                dataBitstream(Offset + i) = dataFill(i Mod 2)
            Next
        End If

        Dim ECCQueue As New Collections.Generic.Queue(Of Byte)(dataBitstream)

        Dim DataBlocks As Byte()() = New Byte(NumBlocks1 + NumBlocks2 - 1)() {}
        Dim ErrorBlocks As Byte()() = New Byte(NumBlocks1 + NumBlocks2 - 1)() {}

        ' Calculate the error correction codewords and generate ouput data blocks
        ' TODO: repeated code.
        For i As Integer = 0 To NumBlocks1 - 1 ' Numblocks1 blocks of size BlockSize (DataCodewords+ECCodewords)
            DataBlocks(i) = New Byte(DataCodewords - 1) {}
            ErrorBlocks(i) = New Byte(ECCodewords) {} ' Extra byte used in calculations
            For j As Integer = 0 To DataCodewords - 1
                DataBlocks(i)(j) = ECCQueue.Dequeue()
                Dim tmp As Byte = ErrorBlocks(i)(0) Xor DataBlocks(i)(j)
                If tmp <> 0 Then
                    For k As Integer = 0 To ECCodewords - 1
                        ErrorBlocks(i)(k) = ALog((Log(tmp) + Log(G(ECCodewords - k - 1))) Mod 255)
                        ErrorBlocks(i)(k) = ErrorBlocks(i)(k) Xor ErrorBlocks(i)(k + 1)
                    Next
                Else
                    For k As Integer = 0 To ECCodewords - 2
                        ErrorBlocks(i)(k) = ErrorBlocks(i)(k + 1)
                    Next
                    ErrorBlocks(i)(ECCodewords - 1) = 0
                End If
            Next
        Next
        For i As Integer = 0 To NumBlocks2 - 1 ' Numblocks2 blocks of size BlockSize+1 (DataCodewords+ECCodewords+1)
            DataBlocks(NumBlocks1 + i) = New Byte(DataCodewords) {}
            ErrorBlocks(NumBlocks1 + i) = New Byte(ECCodewords) {} ' Extra byte used in calculations
            For j As Integer = 0 To DataCodewords
                DataBlocks(NumBlocks1 + i)(j) = ECCQueue.Dequeue()
                Dim tmp As Byte = ErrorBlocks(NumBlocks1 + i)(0) Xor DataBlocks(NumBlocks1 + i)(j)
                If tmp <> 0 Then
                    For k As Integer = 0 To ECCodewords - 1
                        ErrorBlocks(NumBlocks1 + i)(k) = ALog((Log(tmp) + Log(G(ECCodewords - k - 1))) Mod 255)
                        ErrorBlocks(NumBlocks1 + i)(k) = ErrorBlocks(NumBlocks1 + i)(k) Xor ErrorBlocks(NumBlocks1 + i)(k + 1)
                    Next
                Else
                    For k As Integer = 0 To ECCodewords - 2
                        ErrorBlocks(NumBlocks1 + i)(k) = ErrorBlocks(NumBlocks1 + i)(k + 1)
                    Next
                    ErrorBlocks(NumBlocks1 + i)(ECCodewords - 1) = 0
                End If
            Next
        Next

        Dim qrcodeData As Byte() = New Byte(DataECCodewords) {} ' Incl. remainder bits
        Dim qrcodeDataCounter As Integer

        For i As Integer = 0 To DataCodewords - 1
            For j As Integer = 0 To NumBlocks1 + NumBlocks2 - 1
                qrcodeData(qrcodeDataCounter) = DataBlocks(j)(i)
                qrcodeDataCounter += 1
            Next
        Next
        For i As Integer = NumBlocks1 To NumBlocks1 + NumBlocks2 - 1
            qrcodeData(qrcodeDataCounter) = DataBlocks(i)(DataCodewords)
            qrcodeDataCounter += 1
        Next
        For i As Integer = 0 To ECCodewords - 1
            For j As Integer = 0 To NumBlocks1 + NumBlocks2 - 1
                qrcodeData(qrcodeDataCounter) = ErrorBlocks(j)(i)
                qrcodeDataCounter += 1
            Next
        Next

        Dim AlignmentPatternSpacing As Byte = Math.Ceiling(Math.Round((4 + 4 * qrcodeVersion) / (qrcodeVersion \ 7 + 1)) / 2) * 2

        Dim ModuleMatrix As Byte()() = New Byte(ModulesAcross - 1)() {}
        For i As Integer = 0 To ModulesAcross - 1
            ModuleMatrix(i) = New Byte(ModulesAcross - 1) {}
        Next

        ' Position Detection Patterns
        For i As Integer = 0 To 7
            For j As Integer = 0 To 7
                If (i Mod 6 = 0 And j < 7) Or (j Mod 6 = 0 And i < 7) Then
                    ModuleMatrix(i)(j) = 3
                    ModuleMatrix(ModulesAcross - i - 1)(j) = 3
                    ModuleMatrix(j)(ModulesAcross - i - 1) = 3
                ElseIf i > 1 And i < 5 And j > 1 And j < 5 Then
                    ModuleMatrix(i)(j) = 3
                    ModuleMatrix(ModulesAcross - i - 1)(j) = 3
                    ModuleMatrix(j)(ModulesAcross - i - 1) = 3
                Else
                    ModuleMatrix(i)(j) = 2
                    ModuleMatrix(ModulesAcross - i - 1)(j) = 2
                    ModuleMatrix(j)(ModulesAcross - i - 1) = 2
                End If
            Next
        Next

        ' Timing Patterns
        For i As Integer = 8 To ModulesAcross - 9 Step 2
            ModuleMatrix(i)(6) = 3
            ModuleMatrix(6)(i) = 3
            ModuleMatrix(i + 1)(6) = 2
            ModuleMatrix(6)(i + 1) = 2
        Next

        ' Alignment Patterns
        For k As Integer = -2 To 2
            For l As Integer = -2 To 2
                Dim Value As Byte
                If Math.Abs(k) = 2 Or Math.Abs(l) = 2 Or (k = 0 And l = 0) Then
                    Value = 3
                Else
                    Value = 2
                End If
                For i As Integer = 1 To AlignmentPatternsAcross - 2
                    ModuleMatrix(k + ModulesAcross - 7 - i * AlignmentPatternSpacing)(l + 6) = Value
                    ModuleMatrix(l + 6)(k + ModulesAcross - 7 - i * AlignmentPatternSpacing) = Value
                Next
                For i As Integer = 0 To AlignmentPatternsAcross - 2
                    For j As Integer = 0 To AlignmentPatternsAcross - 2
                        ModuleMatrix(k + ModulesAcross - 7 - i * AlignmentPatternSpacing)(l + ModulesAcross - 7 - j * AlignmentPatternSpacing) = Value
                        ModuleMatrix(l + ModulesAcross - 7 - j * AlignmentPatternSpacing)(k + ModulesAcross - 7 - i * AlignmentPatternSpacing) = Value
                    Next
                Next
            Next
        Next

        'Format Information Placeholder
        For i As Integer = 0 To 7
            ModuleMatrix(8)(ModulesAcross - 1 - i) = 2
            ModuleMatrix(ModulesAcross - 1 - i)(8) = 2
            ModuleMatrix(8)(i + i \ 6) = 2
            ModuleMatrix(i + i \ 6)(8) = 2
        Next

        'Version Information Placeholder
        If qrcodeVersion >= 7 Then
            For i As Integer = 0 To 2
                For j As Integer = 0 To 5
                    ModuleMatrix(j)(ModulesAcross - 9 - i) = 2
                    ModuleMatrix(ModulesAcross - 9 - i)(j) = 2
                Next
            Next
        End If

        ' Fill with data
        BitOffset = 0
        For i As Integer = (ModulesAcross \ 4) - 1 To 0 Step -1
            For j As Integer = ModulesAcross * 2 - 1 To 1 - ModulesAcross * 2 Step -2
                For k As Integer = 3 To 2 Step -1
                    Dim x As Integer = i * 4 + k + Math.Sign(j)
                    Dim y As Integer = Math.Abs(j \ 2)
                    If x < 7 Then x -= 1 ' skip column 6 (timing pattern)
                    If ModuleMatrix(x)(y) = 0 Then
                        If ((qrcodeData(BitOffset \ 8) << (BitOffset Mod 8)) And 128) > 0 Then ModuleMatrix(x)(y) = 1
                        BitOffset += 1
                    End If
                Next
            Next
        Next

        'Apply mask(s)
        Dim MaskedMatrix As Byte()() = New Byte(ModulesAcross - 1)() {}
        For i As Integer = 0 To ModulesAcross - 1
            MaskedMatrix(i) = New Byte(ModulesAcross - 1) {}
        Next

        For j As Integer = 0 To ModulesAcross - 1
            For i As Integer = 0 To ModulesAcross - 1
                If ModuleMatrix(j)(i) >> 1 = 0 Then
                    Dim DataBit As Byte = ModuleMatrix(j)(i) And 1
                    MaskedMatrix(j)(i) = DataBit Xor (CByte(((i * j) Mod 3 + (i + j) Mod 2) Mod 2 = 0) And 1) ' 111
                    MaskedMatrix(j)(i) <<= 1
                    MaskedMatrix(j)(i) += DataBit Xor (CByte(((i * j) Mod 2 + (i * j) Mod 3) Mod 2 = 0) And 1) ' 110
                    MaskedMatrix(j)(i) <<= 1
                    MaskedMatrix(j)(i) += DataBit Xor (CByte((i * j) Mod 2 + (i * j) Mod 3 = 0) And 1) ' 101
                    MaskedMatrix(j)(i) <<= 1
                    MaskedMatrix(j)(i) += DataBit Xor (CByte((i \ 2 + j \ 3) Mod 2 = 0) And 1) ' 100
                    MaskedMatrix(j)(i) <<= 1
                    MaskedMatrix(j)(i) += DataBit Xor (CByte((i + j) Mod 3 = 0) And 1) ' 011
                    MaskedMatrix(j)(i) <<= 1
                    MaskedMatrix(j)(i) += DataBit Xor (CByte(j Mod 3 = 0) And 1) ' 010
                    MaskedMatrix(j)(i) <<= 1
                    MaskedMatrix(j)(i) += DataBit Xor (CByte(i Mod 2 = 0) And 1) ' 001
                    MaskedMatrix(j)(i) <<= 1
                    MaskedMatrix(j)(i) += DataBit Xor (CByte((i + j) Mod 2 = 0) And 1) ' 000
                Else
                    MaskedMatrix(j)(i) = (ModuleMatrix(j)(i) And 1) * &HFF
                End If
            Next
        Next

        'Format Information
        For j As Integer = 7 To 0 Step -1
            Dim formatInformation As UInteger = ((ErrorCorrectionLevel << 3) Or j) << 10
            Dim remainder As UInteger = formatInformation
            While remainder > 1023
                Dim degree As Byte = 0
                While (2048 << degree) < remainder
                    degree += 1
                End While
                remainder = remainder Xor (1335 << degree)
            End While
            formatInformation += remainder
            formatInformation = formatInformation Xor &H5412

            For i As Integer = 0 To 7
                MaskedMatrix(ModulesAcross - 1 - i)(8) += (formatInformation >> i) And 1
                If j <> 0 Then MaskedMatrix(ModulesAcross - 1 - i)(8) <<= 1
                MaskedMatrix(8)(i + i \ 6) += (formatInformation >> i) And 1
                If j <> 0 Then MaskedMatrix(8)(i + i \ 6) <<= 1
            Next
            For i As Integer = 0 To 6
                MaskedMatrix(8)(ModulesAcross - 7 + i) += (formatInformation >> (i + 8)) And 1
                If j <> 0 Then MaskedMatrix(8)(ModulesAcross - 7 + i) <<= 1
                MaskedMatrix(i + i \ 6)(8) += (formatInformation >> (14 - i)) And 1
                If j <> 0 Then MaskedMatrix(i + i \ 6)(8) <<= 1
            Next
        Next
        MaskedMatrix(8)(ModulesAcross - 8) = &HFF

        'Version Information
        If qrcodeVersion > 6 Then
            Dim versionInformation As UInteger = CUInt(qrcodeVersion) << 12
            Dim remainder As UInteger = versionInformation
            While remainder > 4095
                Dim degree As Byte = 0
                While (8192 << degree) < remainder
                    degree += 1
                End While
                remainder = remainder Xor (7973 << degree)
            End While
            versionInformation += remainder

            For i As Integer = 0 To 2
                For j As Integer = 0 To 5
                    MaskedMatrix(j)(ModulesAcross - 11 + i) = ((versionInformation >> i + j * 3) And 1) * &HFF
                    MaskedMatrix(ModulesAcross - 11 + i)(j) = ((versionInformation >> i + j * 3) And 1) * &HFF
                Next
            Next
        End If

        ' Scoring of masking results
        Dim penaltyScore As Integer() = New Integer(7) {}
        ' proportion of light and dark
        Dim darkcount As Integer() = New Integer(7) {}

        For i As Integer = 0 To ModulesAcross - 1
            ' remember the last 5 runs for each row/col for each bit
            Dim runs As Integer()()() = New Integer(7)()() {}
            Dim current As Byte() = New Byte(1) {}
            Dim previous As Byte() = New Byte(1) {MaskedMatrix(i)(0), MaskedMatrix(0)(i)}

            Dim tmp As Byte = MaskedMatrix(i)(0)
            For k As Integer = 0 To 7 ' is it quicker to dim outside loop and just zero the array here? hard to say, question of scope and compiler optimisation
                runs(k) = New Integer(1)() {New Integer(4) {1, 0, 0, 0, 0}, New Integer(4) {1, 0, 0, 0, 0}}
                darkcount(k) += tmp And 1
                tmp >>= 1
            Next

            ' Count runs of adjacent modules in each row/col of same colour. store previous 4 runs to identify pattern 1:1:3:1:1 dark:light:...
            For j As Integer = 1 To ModulesAcross - 1
                tmp = MaskedMatrix(i)(0)
                For k As Integer = 0 To 7
                    darkcount(k) += tmp And 1
                    tmp >>= 1
                Next

                current(0) = MaskedMatrix(i)(j)
                current(1) = MaskedMatrix(j)(i)
                For m As Integer = 0 To 1
                    Dim changed As Byte = current(m) Xor previous(m)
                    tmp = current(m)
                    For k As Integer = 0 To 7
                        runs(k)(m)(0) += 1
                        If ((changed And 1) = 1) OrElse (j = ModulesAcross - 1) Then
                            If runs(k)(m)(0) > 5 Then penaltyScore(k) += runs(k)(m)(0) - 2
                            ' 1:1:3:1:1 dark:light:dark:light:dark
                            If (tmp And 1) = 1 AndAlso runs(k)(m)(0) = runs(k)(m)(1) AndAlso 3 * runs(k)(m)(0) = runs(k)(m)(2) AndAlso runs(k)(m)(0) = runs(k)(m)(3) AndAlso runs(k)(m)(0) = runs(k)(m)(4) Then penaltyScore(k) += 40
                            For l As Integer = 0 To 3
                                runs(k)(m)(l + 1) = runs(k)(m)(l)
                            Next
                            runs(k)(m)(0) = 1
                        End If
                        changed >>= 1
                        tmp >>= 1
                    Next
                    previous(m) = current(m)
                Next

                ' Count blocks of 2x2
                If i > 0 Then
                    Dim blockAnd As Byte = (MaskedMatrix(i)(j) And MaskedMatrix(i - 1)(j) And MaskedMatrix(i)(j - 1) And MaskedMatrix(i - 1)(j - 1))
                    Dim blockOr As Byte = (MaskedMatrix(i)(j) Or MaskedMatrix(i - 1)(j) Or MaskedMatrix(i)(j - 1) Or MaskedMatrix(i - 1)(j - 1))
                    For k As Integer = 0 To 7
                        If ((blockAnd And 1) = 1) OrElse ((blockOr And 1) = 0) Then penaltyScore(k) += 3
                        blockAnd >>= 1
                        blockOr >>= 1
                    Next
                End If
            Next
        Next

        Dim maskPatternReference As Byte
        For i As Integer = 0 To 7
            penaltyScore(i) += 10 * Math.Abs(Math.Round(((0.5 - (darkcount(i) / (ModulesAcross ^ 2))) * 20)))
            If penaltyScore(i) < penaltyScore(maskPatternReference) Then maskPatternReference = i
        Next

        Dim MaskedModuleMatrix As Boolean()() = New Boolean(ModulesAcross - 1)() {}
        For i As Integer = 0 To ModulesAcross - 1
            MaskedModuleMatrix(i) = New Boolean(ModulesAcross - 1) {}
        Next

        For i As Integer = 0 To ModulesAcross - 1
            For j As Integer = 0 To ModulesAcross - 1
                MaskedModuleMatrix(i)(j) = (MaskedMatrix(i)(j) >> maskPatternReference) And 1
            Next
        Next

        Return MaskedModuleMatrix
    End Function

    ' write up to 24 bits contained in dataVal to byte array bit buffer
    Private Sub writeBits(ByRef outBuffer As Byte(), ByRef bitOffset As UInteger, ByVal dataVal As UInteger, ByVal bitCount As Byte)
        dataVal <<= 8 - (bitOffset Mod 8) + 24 - bitCount ' shift left by number of bits left in last byte in buffer. first bit of data is highest
        For i As Integer = 0 To Math.Ceiling((bitCount + (bitOffset Mod 8)) / 8) - 1
            If (bitOffset \ 8) + i < outBuffer.Length Then outBuffer((bitOffset \ 8) + i) = outBuffer((bitOffset \ 8) + i) Or ((dataVal >> 8 * (3 - i)) And &HFF)
        Next
        bitOffset += bitCount
    End Sub
End Module
