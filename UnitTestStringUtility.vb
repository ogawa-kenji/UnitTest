Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports WindowsApp1
<TestClass()> Public Class UnitTestStringUtility

    <TestMethod()> Public Sub Test_StringUtility_ByteLeft()

        'バイト区切り(左)テスト
        Assert.AreEqual(StringUtility.ByteLeft("1234", 2), "12")
        Assert.AreEqual(StringUtility.ByteLeft("1234", -1), "")
        Assert.AreEqual(StringUtility.ByteLeft("12345", 5), "12345")
        Assert.AreEqual(StringUtility.ByteLeft("あいうえ", 6), "あいう")
        Assert.AreEqual(StringUtility.ByteLeft("一二三", 4), "一二")

    End Sub

End Class
