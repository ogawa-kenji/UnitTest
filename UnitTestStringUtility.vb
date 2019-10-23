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

    <TestMethod()> Public Sub Test_StringUtility_ByteRight()

        'バイト区切り(右)テスト
        Assert.AreEqual(StringUtility.ByteRight("1234", 2), "34")
        Assert.AreEqual(StringUtility.ByteRight("1234", -1), "")
        Assert.AreEqual(StringUtility.ByteRight("12345", 5), "12345")
        Assert.AreEqual(StringUtility.ByteRight("あいうえ", 6), "いうえ")
        Assert.AreEqual(StringUtility.ByteRight("一二三", 4), "二三")

    End Sub

    <TestMethod()> Public Sub Test_StringUtility_ByteMid()

        'バイト区切り(右)テスト
        Assert.AreEqual(StringUtility.ByteMid("1234", 2, 2), "23")
        Assert.AreEqual(StringUtility.ByteMid("1234", -1, 1), "")
        Assert.AreEqual(StringUtility.ByteMid("1234", 1, -1), "")
        Assert.AreEqual(StringUtility.ByteMid("12345", 1, 5), "12345")
        Assert.AreEqual(StringUtility.ByteMid("あいうえ", 5, 4), "うえ")
        Assert.AreEqual(StringUtility.ByteMid("一二三", 3, 4), "二三")

    End Sub

    <TestMethod()> Public Sub Test_StringUtility_StringCount()

        'バイト区切り(右)テスト
        Assert.AreEqual(StringUtility.StringCount("1234"), 4)
        Assert.AreEqual(StringUtility.StringCount("あいうえお"), 5)
        Assert.AreEqual(StringUtility.StringCount("一二三"), 3)
        Assert.AreEqual(StringUtility.StringCount("𩸽鷗"), 2)
        Assert.AreEqual(StringUtility.StringCount("äőë"), 3)

    End Sub

End Class
