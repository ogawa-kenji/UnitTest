Imports System.Text

Public Class StringUtility
    Private Shared _SpRegex As RegularExpressions.Regex =
                  New RegularExpressions.Regex("^[^\uD800-\uDBFF\uDC00-\uDFFF]+$")

    Public Shared Function ByteLeft(value As String, len As Integer) As String
        If len < 0 Then
            Return String.Empty
        End If
        Dim byteSJIS As Byte() = System.Text.Encoding.GetEncoding("Shift-JIS").GetBytes(value)
        Return System.Text.Encoding.Default.GetString(byteSJIS, 0, len)
    End Function

    Public Shared Function ByteRight(value As String, len As Integer) As String
        If len < 0 Then
            Return String.Empty
        End If
        Dim byteSJIS As Byte() = System.Text.Encoding.GetEncoding("Shift-JIS").GetBytes(value)
        Return System.Text.Encoding.Default.GetString(byteSJIS, byteSJIS.Length - len, len)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="start">開始バイト位置 1～</param>
    ''' <param name="len">区切バイト数</param>
    ''' <returns></returns>
    Public Shared Function ByteMid(ByVal value As String, ByVal start As Integer, ByVal len As Integer) As String
        If start < 0 OrElse len < 0 Then
            Return String.Empty
        End If
        Dim byteSJIS As Byte() = System.Text.Encoding.GetEncoding("Shift-JIS").GetBytes(value)
        Return System.Text.Encoding.Default.GetString(byteSJIS, start - 1, len)
    End Function

    ''' <summary>
    ''' 文字列の文字数
    ''' サロゲートペア・結合文字も1文字をカウントする
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Shared Function StringCount(ByVal value As String) As Integer
        Dim si As New System.Globalization.StringInfo(value)
        'サロゲートペア・結合文字も1文字とカウントする
        Return si.LengthInTextElements()
    End Function

    ''' <summary>
    ''' シフトJIS判定
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsShiftJisCode(ByVal str As String) As Boolean
        ' 全ての文字がシフトJIS表現可能かつJIS2004追加文字コードではない

        ' 文字列なしはTrue
        If String.IsNullOrEmpty(str) Then
            Return True
        End If

        If _SpRegex.IsMatch(str) Then
            ' マッチする=サロゲートペアで使用しているコード範囲は存在しない

        Else
            ' マッチしない=サロゲートペアで使用しているコード範囲は存在する
            Return False

        End If

        Dim byteSJIS As Byte() = Nothing
        Dim chkChar() As Char = Nothing
        Dim byteCode As Integer = 0
        ' 1文字ずつチェックを実施
        For Each chrBuf As Char In str

            byteSJIS = System.Text.Encoding.GetEncoding("Shift-JIS").GetBytes(chrBuf)
            chkChar = System.Text.Encoding.GetEncoding("Shift-JIS").GetChars(byteSJIS)
            If chkChar(0) <> chrBuf Then
                ' Char→Byte()→Charに変換した場合に元に戻らない場合はShift-JIS未定義文字なのでFalse
                Return False
            End If

            byteCode = 0
            If byteSJIS.Length = 1 Then
                byteCode = byteSJIS(0) And &HFF
            ElseIf byteSJIS.Length = 2 Then
                byteCode = (((byteSJIS(0) And &HFF) << 8) Or (byteSJIS(1) And &HFF))
            End If

            ' JIS2004で追加された10文字の場合はFalseを返却
            ' &H879F 「倶」の厳密異体字
            ' &H889E 「剥」の厳密異体字
            ' &H9873 「叱」の慣例異体字
            ' &H989E 「呑」の厳密異体字
            ' &HEAA5 「嘘」の厳密異体字
            ' &HEFF8 「妍」の厳密異体字
            ' &HEFF9 「屏」の厳密異体字
            ' &HEFFA 「并」の厳密異体字
            ' &HEFFB 「痩」の厳密異体字
            ' &HEFFC 「繋」の厳密異体字
            Select Case byteCode
                Case &H879F, &H889E, &H9873, &H989E, &HEAA5, &HEFF8, &HEFF9, &HEFFA, &HEFFB, &HEFFC

                    Return False

                Case Else

            End Select

        Next

        ' 全文字OKならTrueを返却
        Return True

    End Function

End Class
