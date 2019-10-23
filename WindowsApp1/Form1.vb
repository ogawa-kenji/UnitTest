Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim si As New System.Globalization.StringInfo()
        si.String = "𩸽鷗"

        Dim textEmurator As System.Globalization.TextElementEnumerator = System.Globalization.StringInfo.GetTextElementEnumerator("𩸽鷗")

        While textEmurator.MoveNext
            Debug.WriteLine(textEmurator.GetTextElement())
            Debug.WriteLine(Char.IsSurrogate(textEmurator.GetTextElement()))
        End While

        'Debug.WriteLine(StringUtility.IsShiftJisCode("叱"))
        Debug.WriteLine(StringUtility.IsShiftJisCode("瘦"))
        Debug.WriteLine(ChrW(&H2E21))



    End Sub
End Class
